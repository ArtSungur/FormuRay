using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using FormuRay.Core;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel
{
    public sealed class UsedRangeInspection
    {
        public UsedRangeCleanupPlan Plan { get; private set; }
        public ContentRegionAnalysis Regions { get; private set; }
        internal UsedRangeInspection(UsedRangeCleanupPlan plan, ContentRegionAnalysis regions)
        {
            Plan = plan; Regions = regions;
        }
    }

    public sealed class UsedRangeOptimizationPreview
    {
        public UsedRangeCleanupPlan Plan { get; private set; }
        public ContentRemovalPreview Removal { get; private set; }
        public bool Manual { get; private set; }
        internal UsedRangeOptimizationPreview(UsedRangeCleanupPlan plan, ContentRemovalPreview removal, bool manual)
        {
            Plan = plan; Removal = removal; Manual = manual;
        }
    }

    public sealed class UsedRangeOptimizerService : IDisposable
    {
        private readonly X.Application application;
        private X.Worksheet sheet;
        public string SheetName { get; private set; }

        public UsedRangeOptimizerService(X.Application application)
        {
            this.application = application;
            sheet = application.ActiveSheet as X.Worksheet;
            if (sheet == null) throw new InvalidOperationException("Select a worksheet before opening Used Range Optimizer.");
            SheetName = sheet.Name;
        }

        public UsedRangeInspection Inspect()
        {
            X.Range used = null, rows = null, columns = null;
            try
            {
                used = sheet.UsedRange;
                rows = used.Rows; columns = used.Columns;
                var rectangle = new CellRectangle(used.Row, used.Column,
                    used.Row + rows.Count - 1, used.Column + columns.Count - 1);
                int row, column;
                FindContentBoundary(out row, out column);
                var plan = UsedRangeCleanupPlan.Detect(rectangle, row, column);
                var areas = new List<ContentArea>();
                if (row != 0)
                {
                    ReadAreas(X.XlCellType.xlCellTypeConstants, CellContentKind.Constant, areas);
                    ReadAreas(X.XlCellType.xlCellTypeFormulas, CellContentKind.Formula, areas);
                }
                return new UsedRangeInspection(plan, new ContentRegionAnalysis(areas));
            }
            finally { ComReference.Release(columns); ComReference.Release(rows); ComReference.Release(used); }
        }

        private void ReadAreas(X.XlCellType type, CellContentKind kind, List<ContentArea> destination)
        {
            X.Range cells = null, populated = null;
            X.Areas areas = null;
            try
            {
                cells = sheet.Cells;
                try { populated = cells.SpecialCells(type, Type.Missing); }
                catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x800A03EC)) { return; }
                areas = populated.Areas;
                for (int i = 1; i <= areas.Count; i++)
                {
                    X.Range area = null, rows = null, columns = null;
                    try
                    {
                        area = areas[i];
                        rows = area.Rows; columns = area.Columns;
                        destination.Add(new ContentArea(new CellRectangle(area.Row, area.Column,
                            area.Row + rows.Count - 1, area.Column + columns.Count - 1), kind));
                    }
                    finally { ComReference.Release(columns); ComReference.Release(rows); ComReference.Release(area); }
                }
            }
            finally { ComReference.Release(areas); ComReference.Release(populated); ComReference.Release(cells); }
        }

        public UsedRangeOptimizationPreview Prepare(int retainedRow, int retainedColumn, bool manual)
        {
            UsedRangeInspection inspection = Inspect();
            int row = manual ? retainedRow : Math.Max(retainedRow, inspection.Plan.ContentRow);
            int column = manual ? retainedColumn : Math.Max(retainedColumn, inspection.Plan.ContentColumn);
            var plan = new UsedRangeCleanupPlan(inspection.Plan.Used,
                inspection.Plan.ContentRow, inspection.Plan.ContentColumn, row, column);
            return new UsedRangeOptimizationPreview(plan, inspection.Regions.RemovalFor(plan), manual);
        }

        private void FindContentBoundary(out int row, out int column)
        {
            X.Range cells = null, origin = null, lastRow = null, lastColumn = null;
            try
            {
                cells = sheet.Cells;
                origin = (X.Range)cells[1, 1];
                lastRow = cells.Find("*", origin, X.XlFindLookIn.xlFormulas, X.XlLookAt.xlPart,
                    X.XlSearchOrder.xlByRows, X.XlSearchDirection.xlPrevious, false, false, false);
                if (lastRow == null) { row = 0; column = 0; return; }
                lastColumn = cells.Find("*", origin, X.XlFindLookIn.xlFormulas, X.XlLookAt.xlPart,
                    X.XlSearchOrder.xlByColumns, X.XlSearchDirection.xlPrevious, false, false, false);
                row = lastRow.Row;
                column = lastColumn.Column;
            }
            finally
            {
                ComReference.Release(lastColumn); ComReference.Release(lastRow);
                ComReference.Release(origin); ComReference.Release(cells);
            }
        }

        public string Optimize(UsedRangeOptimizationPreview preview, bool contentRemovalConfirmed)
        {
            if (preview == null) throw new ArgumentNullException("preview");
            if (!preview.Removal.IsAuthorized(contentRemovalConfirmed))
                throw new InvalidOperationException("Content removal was cancelled. Nothing was cleared.");
            if (sheet.ProtectContents || sheet.ProtectDrawingObjects || sheet.ProtectScenarios)
                throw new InvalidOperationException("This worksheet is protected. Unprotect it before optimizing its used range.");
            // Recheck the cells to be removed after confirmation. Never apply a stale content warning.
            UsedRangeInspection fresh = Inspect();
            int safeRow = preview.Manual ? preview.Plan.RetainedRow
                : Math.Max(preview.Plan.RetainedRow, fresh.Plan.ContentRow);
            int safeColumn = preview.Manual ? preview.Plan.RetainedColumn
                : Math.Max(preview.Plan.RetainedColumn, fresh.Plan.ContentColumn);
            var plan = new UsedRangeCleanupPlan(fresh.Plan.Used, fresh.Plan.ContentRow,
                fresh.Plan.ContentColumn, safeRow, safeColumn);
            ContentRemovalPreview currentRemoval = fresh.Regions.RemovalFor(plan);
            if (!preview.Removal.SameContent(currentRemoval))
                throw new InvalidOperationException("Sheet content changed after the warning. Nothing was cleared. Reopen Used Range Optimizer and review the new cutoff.");
            if (!preview.Manual && currentRemoval.RequiresConfirmation)
                throw new InvalidOperationException("Auto mode found content outside the retained boundary. Nothing was cleared.");
            if (!plan.HasExcess)
                return "No excess area remains. Nothing was cleared.";

            bool updating = application.ScreenUpdating, events = application.EnableEvents;
            object status = application.StatusBar;
            try
            {
                application.ScreenUpdating = false;
                application.EnableEvents = false;
                application.StatusBar = "FormuRay: optimizing used range...";
                Clear(plan.RowTail);
                Clear(plan.ColumnTail);
                // Accessing UsedRange prompts Excel to refresh its internal boundary when possible.
                UsedRangeCleanupPlan after = Inspect().Plan;
                return (after.Used.LastRow <= safeRow && after.Used.LastColumn <= safeColumn
                    ? "Cleanup completed. Current Used Range: " + Address(after.Used.LastRow, after.Used.LastColumn) + "."
                    : "Cleanup completed, but Excel still reports the old Used Range. Save, close, and reopen the workbook, then check Ctrl+End again.");
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("Excel could not clear the excess area. Check worksheet protection, merged cells, and workbook permissions. " + ex.Message, ex);
            }
            finally
            {
                try { application.StatusBar = status; }
                finally { try { application.EnableEvents = events; } finally { application.ScreenUpdating = updating; } }
            }
        }

        private void Clear(CellRectangle rectangle)
        {
            if (rectangle == null) return;
            X.Range cells = null, first = null, last = null, area = null;
            try
            {
                cells = sheet.Cells;
                first = (X.Range)cells[rectangle.FirstRow, rectangle.FirstColumn];
                last = (X.Range)cells[rectangle.LastRow, rectangle.LastColumn];
                area = sheet.Range[first, last];
                area.Clear();
            }
            finally { ComReference.Release(area); ComReference.Release(last); ComReference.Release(first); ComReference.Release(cells); }
        }

        public static string Address(int row, int column)
        {
            string letters = "";
            while (column > 0) { column--; letters = (char)('A' + column % 26) + letters; column /= 26; }
            return letters + row;
        }

        public void Dispose() { ComReference.Release(sheet); sheet = null; }
    }
}
