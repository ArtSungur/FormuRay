using System;
using System.Collections.Generic;
using System.Linq;

namespace FormuRay.Core
{
    public sealed class InflowName
    {
        public string FullName { get; set; }
        public string LocalName { get; set; }
        public string ScopeSheet { get; set; }
        // Null for unresolved, external, computed or dynamic names. Retain them for shadowing.
        public string TargetSheet { get; set; }
    }

    public sealed class CrossSheetInflowAnalyzer
    {
        private readonly string sourceSheet;
        private readonly HashSet<string> sheets;
        private readonly List<InflowName> names;
        public CrossSheetInflowAnalyzer(string sourceSheet, IEnumerable<string> worksheets, IEnumerable<InflowName> names = null)
        {
            this.sourceSheet = sourceSheet;
            sheets = new HashSet<string>(worksheets, StringComparer.OrdinalIgnoreCase);
            this.names = (names ?? Enumerable.Empty<InflowName>()).OrderByDescending(n =>
                string.Equals(n.ScopeSheet, sourceSheet, StringComparison.OrdinalIgnoreCase)).ToList();
        }
        private bool OtherSheet(string sheet)
        {
            return !string.IsNullOrEmpty(sheet) && sheets.Contains(sheet)
                && !string.Equals(sheet, sourceSheet, StringComparison.OrdinalIgnoreCase);
        }
        public bool IsInflow(string formula)
        {
            if (string.IsNullOrEmpty(formula) || !formula.StartsWith("=", StringComparison.Ordinal)) return false;
            var references = FormulaReferenceParser.Parse(formula).Where(r => WithinWorksheet(r.Address)).ToList();
            foreach (var reference in references)
                // The legacy parser includes an unquoted implicit-intersection @ in the qualifier.
                if (!reference.IsExternal && OtherSheet(reference.Text.StartsWith("@", StringComparison.Ordinal)
                        ? reference.Sheet.Substring(1) : reference.Sheet)
                    && !(reference.Start > 0 && formula[reference.Start - 1] == ':')
                    && !InsideQualifier(formula, reference.Start)) return true;

            // LET/LAMBDA bindings require a lexical scope engine; do not confuse them with workbook names.
            if (FormulaFunctions.Parse(formula).Any(f => f.Name.Equals("LET", StringComparison.OrdinalIgnoreCase)
                || f.Name.Equals("LAMBDA", StringComparison.OrdinalIgnoreCase)
                || f.Name.EndsWith(".LET", StringComparison.OrdinalIgnoreCase)
                || f.Name.EndsWith(".LAMBDA", StringComparison.OrdinalIgnoreCase))) return false;
            var claimed = new HashSet<int>();
            foreach (var name in names)
            {
                foreach (var match in DefinedNameRules.Find(name.FullName, name.LocalName, sourceSheet, false, formula, formula))
                {
                    int end = match.Start + match.Text.Length;
                    int next = end;
                    while (next < formula.Length && char.IsWhiteSpace(formula[next])) next++;
                    if (next < formula.Length && formula[next] == '(') continue; // Function, not a Name.
                    if (references.Any(r => match.Start >= r.Start && match.Start < r.Start + r.Text.Length)) continue;
                    if (InsideQualifier(formula, match.Start)) continue;
                    if (!claimed.Add(match.Start)) continue; // Worksheet name shadows workbook name, including unresolved names.
                    if (OtherSheet(name.TargetSheet)) return true;
                }
            }
            return false;
        }
        private static bool InsideQualifier(string formula, int position)
        {
            string clean = FormulaReferenceParser.MaskStrings(formula);
            bool quote = false;
            int brackets = 0;
            for (int i = 0; i < position; i++)
            {
                if (clean[i] == '\'')
                {
                    if (quote && i + 1 < position && clean[i + 1] == '\'') i++;
                    else quote = !quote;
                }
                if (clean[i] == '[') brackets++;
                if (clean[i] == ']') brackets--;
            }
            return quote || brackets > 0;
        }
        // RefersToRange is consulted only for a single static, absolute A1 reference.
        public static bool CanResolveName(string refersTo)
        {
            if (DefinedNameRules.SafetyKind(refersTo) != "safe") return false;
            var refs = FormulaReferenceParser.Parse(refersTo);
            if (refs.Count != 1 || refs[0].IsExternal || string.IsNullOrEmpty(refs[0].Sheet) || !WithinWorksheet(refs[0].Address)) return false;
            var reference = refs[0];
            if (refersTo.Trim().TrimStart('=').Trim() != reference.Text) return false;
            string address = reference.Address;
            return address.Split(':').All(part => part.StartsWith("$", StringComparison.Ordinal)
                && (reference.IsWholeRowOrColumn || part.Count(c => c == '$') == 2));
        }
        private static bool WithinWorksheet(string address)
        {
            foreach (string part in address.Replace("$", "").ToUpperInvariant().Split(':'))
            {
                int split = 0, column = 0, row;
                while (split < part.Length && part[split] >= 'A' && part[split] <= 'Z')
                    column = column * 26 + part[split++] - 'A' + 1;
                if (split > 0 && (column < 1 || column > 16384)) return false;
                if (split < part.Length && (!int.TryParse(part.Substring(split), out row) || row < 1 || row > 1048576)) return false;
            }
            return true;
        }
    }
}
