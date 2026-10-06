using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using FormuRay.Core;
using FormuRay.Excel.Names;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel.Audit
{
    public sealed class CrossSheetInflowService : IDisposable
    {
        private readonly X.Application application;
        private readonly Func<HighlightStyle> highlightStyle;
        private readonly AuditFillSession fill = new AuditFillSession(true);
        public bool IsActive { get; private set; }
        public int Matched { get; private set; }
        public int Highlighted { get; private set; }
        public CrossSheetInflowService(X.Application application, Func<HighlightStyle> highlightStyle = null) { this.application = application; this.highlightStyle = highlightStyle ?? (() => HighlightSettings.Defaults.Inflow); }
        public void Toggle(bool on) { if (on) Run(); else Clear(); }
        public void Run()
        {
            Clear();
            var style = highlightStyle();
            var sheet = application.ActiveSheet as X.Worksheet;
            var book = application.ActiveWorkbook;
            if (sheet == null || book == null) throw new InvalidOperationException("Cross Sheet Inflow requires an active worksheet.");
            if (sheet.ProtectContents) throw new InvalidOperationException("Unprotect the active worksheet before running Cross Sheet Inflow.");
            bool updating = application.ScreenUpdating, events = application.EnableEvents;
            try
            {
                application.ScreenUpdating = false;
                application.EnableEvents = false;
                var sheets = ExcelNameReader.Sheets(book);
                var analyzer = new CrossSheetInflowAnalyzer(sheet.Name, sheets, ExcelNameReader.InflowNames(book, sheets));
                var matches = new List<Tuple<int, int>>();
                ExcelFormulaReader.VisitCells(sheet, (formula, row, column) =>
                {
                    if (analyzer.IsInflow(formula)) matches.Add(Tuple.Create(row, column));
                    return true;
                });
                Matched = matches.Count;
                foreach (var match in matches)
                {
                    try
                    {
                        var cell = (X.Range)sheet.Cells[match.Item1, match.Item2];
                        // AuditFillSession owns its own enumerated cell RCWs until Clear.
                        try { if (fill.Apply(cell, style, 1, true)) Highlighted++; }
                        finally { ComReference.Release(cell); }
                    }
                    catch (COMException) { }
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
        public void Clear()
        {
            try { fill.Clear(); }
            finally { IsActive = false; Matched = Highlighted = 0; }
        }
        public void Dispose() { Clear(); }
    }
}
