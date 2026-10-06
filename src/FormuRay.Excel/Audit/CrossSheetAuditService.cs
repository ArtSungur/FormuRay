using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using FormuRay.Core;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel.Audit
{
    public sealed class CrossSheetAuditService : IDisposable
    {
        private readonly X.Application application;
        private readonly Func<HighlightStyle> highlightStyle;
        private readonly AuditFillSession fill = new AuditFillSession(true);
        public bool IsActive { get; private set; }
        public CrossSheetAuditService(X.Application application, Func<HighlightStyle> highlightStyle = null) { this.application = application; this.highlightStyle = highlightStyle ?? (() => HighlightSettings.Defaults.Outflow); }
        public void Toggle(bool on) { if (on) Run(); else Clear(); }
        public void Run()
        {
            Clear();
            var style = highlightStyle();
            var book = application.ActiveWorkbook;
            var source = application.ActiveSheet as X.Worksheet;
            if (book == null) throw new InvalidOperationException("Cross Sheet Outflow requires an active workbook.");
            if (source == null) throw new InvalidOperationException("Cross Sheet Outflow requires an active worksheet.");
            string sourceName = source.Name, sourceCodeName = CodeName(source);
            int sheetRows = source.Rows.Count, sheetColumns = source.Columns.Count;
            bool updating = application.ScreenUpdating, events = application.EnableEvents;
            try
            {
                application.ScreenUpdating = false;
                application.EnableEvents = false;
                int highlighted = 0;
                bool limitReached = false;
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                // Include hidden sheets, just as the VBA Cross Sheet audit does.
                foreach (X.Worksheet sheet in book.Worksheets)
                {
                    string codeName = CodeName(sheet);
                    if (sheet.Name == sourceName || (!string.IsNullOrEmpty(codeName) && codeName == sourceCodeName)) continue;
                    ExcelFormulaReader.Visit(sheet, formula =>
                    {
                        if (!CrossSheetAuditRules.CouldReferenceSheet(formula, sourceName)) return true;
                        foreach (var reference in FormulaReferenceParser.Parse(formula))
                        {
                            if (!CrossSheetAuditRules.Accepts(reference, sourceName, formula)) continue;
                            try
                            {
                                var target = source.Range[reference.Address];
                                if (Convert.ToDouble(target.CountLarge) > CrossSheetAuditRules.MaximumRangeCells
                                    || target.Rows.Count == sheetRows || target.Columns.Count == sheetColumns) continue;
                                if (highlighted >= CrossSheetAuditRules.MaximumHighlightedRanges) { limitReached = true; return false; }
                                if (!seen.Add(target.Address[false, false, X.XlReferenceStyle.xlA1])) continue;
                                if (fill.Apply(target, style, CrossSheetAuditRules.MaximumRangeCells, true)) highlighted++;
                            }
                            catch (COMException) { }
                        }
                        return true;
                    });
                    if (limitReached) break;
                }
                IsActive = true;
            }
            catch { Clear(); throw; }
            finally
            {
                try { application.EnableEvents = events; }
                finally { application.ScreenUpdating = updating; }
            }
        }
        public void Clear() { try { fill.Clear(); } finally { IsActive = false; } }
        private static string CodeName(X.Worksheet sheet)
        {
            try { return sheet.CodeName; }
            catch (COMException) { return ""; }
        }
        public void Dispose() { Clear(); }
    }
}
