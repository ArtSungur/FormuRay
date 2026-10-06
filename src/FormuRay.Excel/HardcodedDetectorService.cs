using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using FormuRay.Core;
using Interop = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel
{
    /// <summary>One toggle session across Excel, just like modHardcodedDetector's module state.</summary>
    public sealed class HardcodedDetectorService : IDisposable
    {
        private const int ReadBlockSize = 4096;
        private readonly Interop.Application application;
        private readonly Func<HighlightStyle> highlightStyle;
        private HighlightStyle activeStyle;
        private readonly List<CellFillState> states = new List<CellFillState>();
        private readonly HashSet<string> cellKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public HardcodedDetectorService(Interop.Application application, Func<HighlightStyle> highlightStyle = null)
        {
            if (application == null) throw new ArgumentNullException("application");
            this.application = application;
            this.highlightStyle = highlightStyle ?? (() => HighlightSettings.Defaults.HardCode);
        }

        public bool IsActive { get; private set; }

        public void Toggle(bool turnOn)
        {
            if (turnOn) Run();
            else Clear();
        }

        public void Run()
        {
            // A direct rerun clears the previous sheet before validating the new active sheet.
            Clear();
            activeStyle = highlightStyle();
            Interop.Workbook workbook = null;
            object activeSheet = null;
            Interop.Range usedRange = null;
            Interop.Range constants = null;
            var timer = Stopwatch.StartNew();
            try
            {
                workbook = application.ActiveWorkbook;
                if (workbook == null)
                    throw new InvalidOperationException("HardCode Detector requires an active workbook.");
                activeSheet = application.ActiveSheet;
                var worksheet = activeSheet as Interop.Worksheet;
                if (worksheet == null)
                    throw new InvalidOperationException("HardCode Detector requires an active worksheet.");

                usedRange = worksheet.UsedRange;
                try { constants = usedRange.SpecialCells(Interop.XlCellType.xlCellTypeConstants); }
                catch (COMException ex)
                {
                    // Excel also raises 1004 when SpecialCells finds no constants; VBA treats it as empty.
                    if (ex.ErrorCode != unchecked((int)0x800A03EC)) throw;
                }

                if (constants != null)
                {
                    var previousScreenUpdating = application.ScreenUpdating;
                    var previousEnableEvents = application.EnableEvents;
                    try
                    {
                        application.ScreenUpdating = false;
                        application.EnableEvents = false;
                        ScanConstants(constants);
                    }
                    finally
                    {
                        try { application.EnableEvents = previousEnableEvents; }
                        finally { application.ScreenUpdating = previousScreenUpdating; }
                    }
                }
                // No matches still means the toggle is on, as in VBA.
                IsActive = true;
                Debug.WriteLine("FormuRay HardcodedDetector: sheet=" + worksheet.Name
                    + " snapshots=" + states.Count + " seconds=" + timer.Elapsed.TotalSeconds.ToString("F2"));
            }
            catch
            {
                Clear();
                throw;
            }
            finally
            {
                ComReference.Release(constants);
                ComReference.Release(usedRange);
                ComReference.Release(activeSheet);
                ComReference.Release(workbook);
            }
        }

        private void ScanConstants(Interop.Range constants)
        {
            Interop.Areas areas = null;
            int highlighted = 0;
            try
            {
                areas = constants.Areas;
                int areaCount = areas.Count;
                for (int index = 1; index <= areaCount; index++)
                {
                    Interop.Range area = null;
                    try
                    {
                        area = areas[index];
                        if (ScanArea(area, ref highlighted)) break;
                    }
                    finally { ComReference.Release(area); }
                }
                Debug.WriteLine("FormuRay HardcodedDetector: highlighted=" + highlighted);
            }
            finally { ComReference.Release(areas); }
        }

        // Preserve SpecialCells(all constants) area order and row-major cell order, including at
        // the 20,000-cell boundary. Filtering SpecialCells to numbers can regroup its areas.
        private bool ScanArea(Interop.Range area, ref int highlighted)
        {
            Interop.Range rows = null;
            Interop.Range columns = null;
            Interop.Range cells = null;
            try
            {
                rows = area.Rows;
                columns = area.Columns;
                cells = area.Cells;
                int rowCount = rows.Count;
                int columnCount = columns.Count;
                int blockRows = Math.Max(1, ReadBlockSize / columnCount);
                for (int row = 1; row <= rowCount; row += blockRows)
                {
                    Interop.Range first = null;
                    Interop.Range block = null;
                    try
                    {
                        int height = Math.Min(blockRows, rowCount - row + 1);
                        first = (Interop.Range)cells[row, 1];
                        block = first.Resize[height, columnCount];
                        object values = block.Value2;
                        var matrix = values as object[,];
                        for (int r = 0; r < height; r++)
                        {
                            for (int c = 0; c < columnCount; c++)
                            {
                                object value = matrix == null ? values
                                    : matrix[r + matrix.GetLowerBound(0), c + matrix.GetLowerBound(1)];
                                // Value2 marshals Excel numbers (including dates/currency) as Double;
                                // VT_ERROR values arrive as Int32 HRESULTs (or ErrorWrapper).
                                bool isError = value is int || value is ErrorWrapper;
                                if (!HardcodedInputRules.IsHardcodedInput(value, false, isError)) continue;
                                if (highlighted >= HardcodedInputRules.MaximumHighlightedCells)
                                {
                                    Debug.WriteLine("FormuRay HardcodedDetector: highlight limit reached");
                                    return true;
                                }
                                HighlightCell(cells, row + r, c + 1, ref highlighted);
                            }
                        }
                    }
                    finally
                    {
                        ComReference.Release(block);
                        ComReference.Release(first);
                    }
                }
                return false;
            }
            finally
            {
                ComReference.Release(cells);
                ComReference.Release(columns);
                ComReference.Release(rows);
            }
        }

        private void HighlightCell(Interop.Range cells, int row, int column, ref int highlighted)
        {
            Interop.Range cell = null;
            try
            {
                cell = (Interop.Range)cells[row, column];
                // All cells in a run belong to one sheet; its local address is an equivalent key.
                string key = cell.Address[false, false, Interop.XlReferenceStyle.xlA1];
                if (!cellKeys.Add(key)) return;
                var state = new CellFillState(cell);
                states.Add(state);
                cell = null; // The session now owns the reference, including if Highlight fails.
                state.Highlight(activeStyle);
                highlighted++;
            }
            catch (COMException ex)
            {
                // VBA skips cells that cannot be formatted, e.g. on a protected sheet.
                Debug.WriteLine("FormuRay: could not highlight cell: " + ex.Message);
            }
            finally { ComReference.Release(cell); }
        }

        public void Clear()
        {
            try
            {
                foreach (var state in states)
                {
                    try { state.Restore(); }
                    finally { state.Dispose(); }
                }
            }
            finally
            {
                states.Clear();
                cellKeys.Clear();
                IsActive = false;
            }
        }

        public void Dispose() { Clear(); }
    }
}
