using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using FormuRay.Core;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel
{
    public sealed class FormulaConsistencyService : IDisposable
    {
        private const int ReadBlockSize = 4096;
        private struct FormulaLocation
        {
            internal int Row, Column;
            internal FormulaLocation(int row, int column) { Row = row; Column = column; }
        }
        private readonly X.Application application;
        private readonly Func<HighlightStyle> highlightStyle;
        private readonly List<CellFillState> highlights = new List<CellFillState>();
        public bool IsActive { get; private set; }

        public FormulaConsistencyService(X.Application application, Func<HighlightStyle> highlightStyle = null)
        {
            if (application == null) throw new ArgumentNullException("application");
            this.application = application;
            this.highlightStyle = highlightStyle ?? (() => HighlightSettings.Defaults.FormulaConsistency);
        }

        public int Toggle(bool turnOn)
        {
            if (turnOn) return Run();
            Clear();
            return 0;
        }

        public int Run()
        {
            Clear();
            X.Workbook workbook = null;
            object activeSheet = null;
            try
            {
                workbook = application.ActiveWorkbook;
                if (workbook == null)
                    throw new InvalidOperationException("Formula Consistency requires an active workbook.");
                activeSheet = application.ActiveSheet;
                var worksheet = activeSheet as X.Worksheet;
                if (worksheet == null)
                    throw new InvalidOperationException("Formula Consistency requires an active worksheet.");

                bool updating = application.ScreenUpdating;
                bool events = application.EnableEvents;
                object status = application.StatusBar;
                try
                {
                    application.ScreenUpdating = false;
                    application.EnableEvents = false;
                    application.StatusBar = "FormuRay: checking formula consistency...";
                    var changes = FormulaConsistencyAnalyzer.FindChanges(ReadFormulaPatterns(worksheet));
                    Highlight(worksheet, changes, highlightStyle());
                    IsActive = true;
                    return changes.Count;
                }
                catch
                {
                    Clear();
                    throw;
                }
                finally
                {
                    try
                    {
                        // This PIA marshals False as the literal text "FALSE" on assignment.
                        // A null value returns control of the status bar to Excel.
                        if (status is bool && !(bool)status) application.StatusBar = null;
                        else application.StatusBar = status;
                    }
                    finally { try { application.EnableEvents = events; } finally { application.ScreenUpdating = updating; } }
                }
            }
            finally
            {
                ComReference.Release(activeSheet);
                ComReference.Release(workbook);
            }
        }

        private static IList<FormulaPatternCell> ReadFormulaPatterns(X.Worksheet worksheet)
        {
            var locations = ReadFormulaLocations(worksheet);
            locations.Sort((left, right) => left.Row == right.Row
                ? left.Column.CompareTo(right.Column) : left.Row.CompareTo(right.Row));
            var result = new List<FormulaPatternCell>();
            X.Range cells = null;
            try
            {
                cells = worksheet.Cells;
                int start = 0;
                while (start < locations.Count)
                {
                    int end = start + 1;
                    while (end < locations.Count && locations[end].Row == locations[end - 1].Row
                        && locations[end].Column == locations[end - 1].Column + 1)
                        end++;
                    if (end - start >= 2) ReadRun(cells, locations, start, end, result);
                    start = end;
                }
                return result;
            }
            finally { ComReference.Release(cells); }
        }

        private static List<FormulaLocation> ReadFormulaLocations(X.Worksheet worksheet)
        {
            var result = new List<FormulaLocation>();
            X.Range used = null, formulas = null;
            X.Areas areas = null;
            try
            {
                used = worksheet.UsedRange;
                try { formulas = used.SpecialCells(X.XlCellType.xlCellTypeFormulas); }
                catch (COMException ex)
                {
                    if (ex.ErrorCode != unchecked((int)0x800A03EC)) throw;
                    return result;
                }
                areas = formulas.Areas;
                for (int index = 1; index <= areas.Count; index++)
                {
                    X.Range area = null;
                    X.Range rows = null, columns = null;
                    try
                    {
                        area = areas[index];
                        rows = area.Rows; columns = area.Columns;
                        int firstRow = area.Row, firstColumn = area.Column;
                        int rowCount = rows.Count, columnCount = columns.Count;
                        for (int row = 0; row < rowCount; row++)
                            for (int column = 0; column < columnCount; column++)
                                result.Add(new FormulaLocation(firstRow + row, firstColumn + column));
                    }
                    finally { ComReference.Release(columns); ComReference.Release(rows); ComReference.Release(area); }
                }
                return result;
            }
            finally
            {
                ComReference.Release(areas);
                ComReference.Release(formulas);
                ComReference.Release(used);
            }
        }

        private static void ReadRun(X.Range cells, IList<FormulaLocation> locations, int start, int end,
            IList<FormulaPatternCell> result)
        {
            for (int index = start; index < end; index += ReadBlockSize)
            {
                X.Range first = null, block = null;
                try
                {
                    int row = locations[index].Row, column = locations[index].Column;
                    int width = Math.Min(ReadBlockSize, end - index);
                    first = (X.Range)cells[row, column];
                    block = first.Resize[1, width];
                    object values;
                    try { values = block.FormulaR1C1; }
                    catch (COMException)
                    {
                        ReadCells(block, result);
                        continue;
                    }
                    var matrix = values as object[,];
                    if (matrix == null) { ReadCells(block, result); continue; }
                    for (int offset = 0; offset < width; offset++)
                        result.Add(new FormulaPatternCell(row, column + offset,
                            Pattern(matrix[matrix.GetLowerBound(0), offset + matrix.GetLowerBound(1)])));
                }
                finally { ComReference.Release(block); ComReference.Release(first); }
            }
        }

        private static void ReadCells(X.Range block, IList<FormulaPatternCell> result)
        {
            X.Range cells = null;
            try
            {
                cells = block.Cells;
                int count = cells.Count;
                for (int index = 1; index <= count; index++)
                {
                    X.Range cell = null;
                    try
                    {
                        cell = (X.Range)cells[index];
                        result.Add(new FormulaPatternCell(cell.Row, cell.Column, Pattern(cell.FormulaR1C1)));
                    }
                    finally { ComReference.Release(cell); }
                }
            }
            finally { ComReference.Release(cells); }
        }

        private static string Pattern(object value)
        {
            string pattern = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (string.IsNullOrEmpty(pattern))
                throw new InvalidOperationException("Excel returned an empty FormulaR1C1 for a formula cell.");
            return pattern;
        }

        private void Highlight(X.Worksheet worksheet, IList<FormulaPatternCell> changes, HighlightStyle style)
        {
            X.Range cells = null;
            try
            {
                cells = worksheet.Cells;
                foreach (var change in changes)
                {
                    X.Range cell = null;
                    try
                    {
                        cell = (X.Range)cells[change.Row, change.Column];
                        var state = new CellFillState(cell);
                        highlights.Add(state);
                        cell = null;
                        state.Highlight(style);
                    }
                    catch (COMException ex)
                    {
                        Debug.WriteLine("FormuRay Formula Consistency: could not highlight cell: " + ex.Message);
                    }
                    finally { ComReference.Release(cell); }
                }
            }
            finally { ComReference.Release(cells); }
        }

        public void Clear()
        {
            try
            {
                foreach (var state in highlights)
                {
                    try { state.Restore(); }
                    finally { state.Dispose(); }
                }
            }
            finally { highlights.Clear(); IsActive = false; }
        }

        public void Dispose() { Clear(); }
    }
}
