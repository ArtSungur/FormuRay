using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using FormuRay.Core;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel.Audit
{
    public sealed class FormulaExplorerService : IDisposable
    {
        private readonly ExcelAuditContext context;
        private readonly AuditFillSession auditedFill = new AuditFillSession();
        private readonly AuditFillSession selectedFill = new AuditFillSession();
        private readonly Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> functionValues = new Dictionary<string, string>();
        public ExplorerDocument Current { get; private set; }
        public ExcelAuditContext Context { get { return context; } }
        public bool IsNavigating { get; private set; }

        public FormulaExplorerService(X.Application application) { context = new ExcelAuditContext(application); }

        public ExplorerDocument Refresh(ExplorerMode mode)
        {
            Clear();
            values.Clear();
            functionValues.Clear();
            var document = new ExplorerDocument { Mode = mode };
            Current = document;
            X.Range source = mode == ExplorerMode.Precedents ? context.Application.ActiveCell as X.Range
                : context.Application.Selection as X.Range ?? context.Application.ActiveCell as X.Range;
            if (source == null || !(context.Application.ActiveSheet is X.Worksheet))
            {
                document.Status = mode == ExplorerMode.Precedents ? "Select a formula cell first." : "Select a cell or range first.";
                return document;
            }
            if (mode == ExplorerMode.Precedents) source = (X.Range)source.Cells[1, 1];
            document.Source = ExcelAuditContext.Location(source);
            if (Convert.ToDouble(source.CountLarge) == 1 && Convert.ToBoolean(source.HasFormula))
            {
                document.Formula = Convert.ToString(source.Formula);
                document.FormulaLocal = Convert.ToString(source.FormulaLocal);
                document.LocalDecimalSeparator = context.Application.DecimalSeparator;
            }
            if (mode == ExplorerMode.Precedents)
            {
                if (document.Formula.Length > 0) BuildPrecedents(document, source);
                document.Status = document.Components.Count > 0 ? document.Components.Count + " formula component(s) found."
                    : document.Formula.Length > 0 ? "No supported references found." : "Active cell does not contain a formula.";
            }
            else
            {
                BuildDependents(document, source);
                document.Status = document.Components.Count > 0 ? document.Components.Count + " dependent formula cell(s) found." : "No dependent formulas found.";
            }
            document.Components.Insert(0, new ExplorerComponent { Kind = "Audited", Display = document.Source.Address,
                Reference = FormulaReferenceParser.QuoteSheet(document.Source.Sheet) + "!" + document.Source.Address,
                Target = document.Source, Value = context.Value(source), Start = -1, Occurrence = 1 });
            auditedFill.Apply(source, AuditFillSession.Green, 10000, false);
            return document;
        }

        private void BuildPrecedents(ExplorerDocument document, X.Range source)
        {
            document.Expression = FormulaReferenceParser.ParseExpression(document.Formula, document.Source.Workbook);
            // Resolve names through the existing safety/scope policy, against the same
            // invariant formula used by the model. UI never reparses the formula.
            var names = new ExplorerDocument { Source = document.Source, Formula = document.Formula };
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int scanned = 0;
            AddNames(names, source.Worksheet.Names, source, true, seen, ref scanned);
            AddNames(names, ((X.Workbook)source.Worksheet.Parent).Names, source, false, seen, ref scanned);
            foreach (var node in document.Expression.Components)
            {
                ExplorerComponent component = null;
                if (node.Reference != null)
                {
                    var reference = node.Reference;
                    var range = context.ResolveReference(reference, source.Worksheet);
                    try
                    {
                        component = new ExplorerComponent { Kind = node.Kind, Display = node.Element, Reference = reference.Text,
                            Occurrence = reference.Occurrence, Target = range == null ? null : ExcelAuditContext.Location(range), Value = CachedValue(range) };
                    }
                    finally { ComReference.Release(range); }
                }
                else if (node.Kind == "Name")
                    component = names.Components.FirstOrDefault(c => c.Start >= node.Start && c.Start < node.Start + node.Length
                        && string.Equals(c.Display, node.Element, StringComparison.OrdinalIgnoreCase));
                if (component == null) component = new ExplorerComponent { Kind = node.Kind, Display = node.Element, Value = node.Value, NonNavigable = true };
                if (node.Kind == "Name") component.Kind = "Name";
                component.Start = node.Start; component.Length = node.Length; component.Formula = document.Formula;
                // Spans are invariant A1 spans. Display that exact formula for precedents.
                component.FormulaLocal = "";
                node.Component = component;
                document.Components.Add(component);
            }
        }

        private void AddNames(ExplorerDocument document, X.Names names, X.Range source, bool localScope, HashSet<string> seen, ref int scanned)
        {
            foreach (X.Name name in names)
            {
                if (++scanned > DefinedNameRules.MaximumNames) break;
                string full = Read(() => name.Name), local = Read(() => name.NameLocal), refers = Read(() => name.RefersTo);
                var matches = DefinedNameRules.Find(full, local, document.Source.Sheet, localScope, document.Formula, document.FormulaLocal);
                if (matches.Count == 0) continue;
                string safety = DefinedNameRules.SafetyKind(refers);
                X.Range target = null;
                if (safety == "safe")
                {
                    try { target = name.RefersToRange; } catch (COMException) { }
                    if (target == null) safety = "unresolved";
                }
                foreach (var match in matches)
                {
                    if (!seen.Add(match.Start + "|" + match.Text)) continue;
                    var component = new ExplorerComponent { Kind = "NamedRange", Display = match.Text, Reference = match.Text, MatchedReference = match.Text,
                        Formula = document.Formula, FormulaLocal = document.FormulaLocal, Start = match.Start, Occurrence = match.Occurrence,
                        SafetyKind = safety, NonNavigable = safety != "safe", Target = target == null ? null : ExcelAuditContext.Location(target),
                        Value = target == null ? (refers.Length == 0 ? "<unresolved>" : refers) : CachedValue(target) };
                    int position = document.Components.FindIndex(c => c.Start < 0 || c.Start > component.Start);
                    if (position < 0) document.Components.Add(component); else document.Components.Insert(position, component);
                }
            }
        }

        private void BuildDependents(ExplorerDocument document, X.Range source)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var location in new DependentTraceReader(context.Application).Read(source)) AddDependent(document, source, context.Resolve(location), seen);
            if (document.Components.Count > 0) return;
            // Match the active VBA flag: no slow workbook-formula scan, even after an empty result.
            try { foreach (X.Range cell in source.DirectDependents.Cells) AddDependent(document, source, cell, seen); }
            catch (COMException) { }
        }

        private void AddDependent(ExplorerDocument document, X.Range source, X.Range cell, HashSet<string> seen)
        {
            if (cell == null) return;
            try
            {
                if (!Convert.ToBoolean(cell.HasFormula)) return;
                var location = ExcelAuditContext.Location(cell);
                if (!seen.Add(location.Key)) return;
                string formula = Convert.ToString(cell.Formula);
                string matched = source.Address[false, false, X.XlReferenceStyle.xlA1];
                foreach (var reference in FormulaReferenceParser.Parse(formula, location.Workbook, false))
                {
                    var target = context.ResolveReference(reference, cell.Worksheet, true);
                    if (target == null || target.Worksheet != source.Worksheet) continue;
                    if (context.Application.Intersect(target, source) == null) continue;
                    matched = reference.Display;
                    break;
                }
                document.Components.Add(new ExplorerComponent { Kind = "Dependent", Display = FormulaReferenceParser.QuoteSheet(location.Sheet) + "!" + location.Address,
                    Formula = formula, FormulaLocal = Convert.ToString(cell.FormulaLocal), MatchedReference = matched,
                    Target = location, Value = CachedValue(cell), Occurrence = 1, Start = -1 });
            }
            catch (COMException) { }
        }

        private string CachedValue(X.Range range)
        {
            if (range == null) return "";
            string key = ExcelAuditContext.Location(range).Key, result;
            if (!values.TryGetValue(key, out result)) values[key] = result = context.Value(range);
            return result;
        }
        private static string Read(Func<object> getter) { try { return Convert.ToString(getter()); } catch (COMException) { return ""; } }

        public string EvaluateFunction(string fragment)
        {
            if (Current == null || string.IsNullOrEmpty(fragment)) return "";
            string result;
            if (!functionValues.TryGetValue(fragment, out result)) functionValues[fragment] = result = context.Evaluate(Current.Source, fragment);
            return result;
        }
        public void Navigate(ExplorerComponent component)
        {
            if (component == null || component.NonNavigable) { ClearSelectionHighlight(); return; }
            var target = context.Resolve(component.Target);
            if (target == null) return;
            IsNavigating = true;
            try
            {
                context.Navigate(target);
                selectedFill.Clear();
                selectedFill.Apply(target, AuditFillSession.Magenta, 10000, false);
            }
            finally { IsNavigating = false; }
        }
        public void ClearSelectionHighlight() { selectedFill.Clear(); }
        public void Clear() { try { selectedFill.Clear(); } finally { auditedFill.Clear(); } }
        public void Dispose() { Clear(); Current = null; values.Clear(); functionValues.Clear(); }
    }
}
