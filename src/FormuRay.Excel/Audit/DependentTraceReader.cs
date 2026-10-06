using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using FormuRay.Core;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel.Audit
{
    internal sealed class DependentTraceReader
    {
        private readonly X.Application application;
        internal DependentTraceReader(X.Application application) { this.application = application; }

        internal List<RangeLocation> Read(X.Range auditedRange)
        {
            var results = new List<RangeLocation>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            X.Range source = (X.Range)auditedRange.Cells[1, 1];
            var originalWorkbook = application.ActiveWorkbook;
            object originalSheet = application.ActiveSheet;
            var originalSelection = application.Selection as X.Range;
            bool updating = application.ScreenUpdating, events = application.EnableEvents, alerts = application.DisplayAlerts;
            try
            {
                application.ScreenUpdating = false;
                application.EnableEvents = false;
                application.DisplayAlerts = false;
                Activate(source);
                Try(() => source.Worksheet.ClearArrows());
                source.ShowDependents();
                string sourceKey = ExcelAuditContext.Location(source).Key;
                for (int arrow = 1; arrow <= 1000; arrow++)
                {
                    bool found = false;
                    for (int link = 1; link <= 1000; link++)
                    {
                        X.Range target;
                        try
                        {
                            Activate(source);
                            target = source.NavigateArrow(false, arrow, link) as X.Range;
                            if (target == null) target = application.Selection as X.Range;
                        }
                        catch (COMException) { break; }
                        if (target == null || string.Equals(ExcelAuditContext.Location((X.Range)target.Cells[1, 1]).Key, sourceKey, StringComparison.OrdinalIgnoreCase)) break;
                        found = true;
                        foreach (X.Range cell in target.Cells)
                        {
                            var location = ExcelAuditContext.Location(cell);
                            if (!string.Equals(location.Key, sourceKey, StringComparison.OrdinalIgnoreCase) && seen.Add(location.Key)) results.Add(location);
                        }
                        // LinkNumber is relevant to external-sheet arrows only; for a local arrow
                        // Excel returns the same destination for every link number.
                        if (target.Worksheet == source.Worksheet) break;
                    }
                    if (!found) break;
                }
            }
            catch (COMException ex) { Debug.WriteLine("FormuRay trace: " + ex.Message); }
            finally
            {
                Try(() => source.Worksheet.ClearArrows());
                Try(() => { if (originalWorkbook != null) originalWorkbook.Activate(); });
                Try(() =>
                {
                    if (originalSelection != null) { originalSelection.Worksheet.Activate(); originalSelection.Select(); }
                    else if (originalSheet is X.Worksheet) ((X.Worksheet)originalSheet).Activate();
                    else if (originalSheet is X.Chart) ((X.Chart)originalSheet).Activate();
                });
                Try(() => application.DisplayAlerts = alerts);
                Try(() => application.EnableEvents = events);
                Try(() => application.ScreenUpdating = updating);
            }
            return results;
        }

        private static void Activate(X.Range source)
        {
            ((X.Workbook)source.Worksheet.Parent).Activate();
            source.Worksheet.Activate();
            source.Select();
        }
        private static void Try(Action action)
        {
            try { action(); } catch (COMException ex) { Debug.WriteLine(ex.Message); }
        }
    }
}
