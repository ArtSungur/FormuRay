using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using FormuRay.Core;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel
{
    public sealed class PasteFormulasService : IDisposable
    {
        // Bound synchronous copy-time work on Excel's UI thread. Native Copy is never cancelled.
        public const int MaximumSnapshotCells = 250000;
        private readonly X.Application application;
        private FormulaClipboardSnapshot snapshot;
        private string captureError;
        public PasteFormulasService(X.Application application) { this.application = application; }

        public void CaptureCopy()
        {
            Clear();
            X.Range source = null;
            try
            {
                source = application.Selection as X.Range;
                if (source == null) return;
                snapshot = ReadSnapshot(source);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                captureError = ex is InvalidOperationException ? ex.Message : "This selection could not be captured. Copy a rectangular cell range and try again.";
            }
            finally { ComReference.Release(source); }
        }

        // Called after Excel processes the unmodified Ctrl+C message.
        public void ConfirmCopy()
        {
            try { if (application.CutCopyMode != X.XlCutCopyMode.xlCopy) Clear(); }
            catch (COMException) { Clear(); }
        }

        public void Clear() { snapshot = null; captureError = null; }

        private FormulaClipboardSnapshot ReadSnapshot(X.Range source)
        {
            X.Areas areas = null;
            X.Range rows = null, columns = null, formulaCells = null;
            X.Worksheet sheet = null;
            X.Workbook book = null;
            try
            {
                areas = source.Areas;
                if (areas.Count != 1) throw new InvalidOperationException("Paste Formulas supports one rectangular source range. Copy a single area.");
                if (Convert.ToDouble(source.CountLarge) > MaximumSnapshotCells)
                    throw new InvalidOperationException("Paste Formulas supports up to 250,000 copied cells. Copy a smaller range.");
                if (!Equals(source.MergeCells, false)) throw new InvalidOperationException("Copy a range without merged cells for Paste Formulas.");
                rows = source.Rows; columns = source.Columns;
                int rowCount = rows.Count, columnCount = columns.Count;
                var mask = new bool[rowCount, columnCount];
                object hasFormula = source.HasFormula;
                if (Equals(hasFormula, true))
                {
                    for (int r = 0; r < rowCount; r++) for (int c = 0; c < columnCount; c++) mask[r, c] = true;
                }
                else if (!Equals(hasFormula, false))
                {
                    // HasFormula distinguishes text beginning '=' from real formulas. Read
                    // formula areas (not individual cells); Formula/Value2 are each bulk reads.
                    formulaCells = source.SpecialCells(X.XlCellType.xlCellTypeFormulas);
                    X.Areas formulaAreas = formulaCells.Areas;
                    try
                    {
                        int firstRow = source.Row, firstColumn = source.Column;
                        for (int i = 1; i <= formulaAreas.Count; i++)
                        {
                            X.Range area = formulaAreas[i], areaRows = null, areaColumns = null;
                            try
                            {
                                areaRows = area.Rows; areaColumns = area.Columns;
                                int r0 = area.Row - firstRow, c0 = area.Column - firstColumn;
                                int rEnd = Math.Min(rowCount, r0 + areaRows.Count), cEnd = Math.Min(columnCount, c0 + areaColumns.Count);
                                for (int r = Math.Max(0, r0); r < rEnd; r++)
                                    for (int c = Math.Max(0, c0); c < cEnd; c++) mask[r, c] = true;
                            }
                            finally { ComReference.Release(areaColumns); ComReference.Release(areaRows); ComReference.Release(area); }
                        }
                    }
                    finally { ComReference.Release(formulaAreas); }
                }
                sheet = source.Worksheet; book = (X.Workbook)sheet.Parent;
                object values = source.Value2;
                PreserveErrorConstants(source, rowCount, columnCount, ref values);
                return new FormulaClipboardSnapshot(book.FullName, sheet.Name, source.Address[true, true, X.XlReferenceStyle.xlA1],
                    rowCount, columnCount, source.Formula, values, mask);
            }
            finally
            {
                ComReference.Release(book); ComReference.Release(sheet); ComReference.Release(formulaCells);
                ComReference.Release(columns); ComReference.Release(rows); ComReference.Release(areas);
            }
        }

        private void PreserveErrorConstants(X.Range source, int height, int width, ref object values)
        {
            // VT_ERROR can marshal as Int32. Identify error constants through Excel,
            // never by their numeric code (which could also be a legitimate value).
            X.Range errors = null;
            X.Areas areas = null;
            try
            {
                try { errors = source.SpecialCells(X.XlCellType.xlCellTypeConstants, X.XlSpecialCellsValue.xlErrors); }
                catch (COMException) { return; } // Excel raises 1004 when none exist.
                // SpecialCells on a single cell can expand to UsedRange. Restrict it
                // natively before enumerating areas so unrelated cells are never visited.
                var intersection = application.Intersect(source, errors);
                ComReference.Release(errors); errors = intersection;
                if (errors == null) return;
                areas = errors.Areas;
                int top = source.Row, left = source.Column;
                var matrix = values as Array;
                for (int i = 1; i <= areas.Count; i++)
                {
                    X.Range area = areas[i], rows = null, columns = null;
                    try
                    {
                        rows = area.Rows; columns = area.Columns;
                        int r0 = area.Row - top, c0 = area.Column - left;
                        int endRow = Math.Min(height, r0 + rows.Count), endColumn = Math.Min(width, c0 + columns.Count);
                        for (int r = Math.Max(0, r0); r < endRow; r++)
                            for (int c = Math.Max(0, c0); c < endColumn; c++)
                            {
                                object value = matrix == null ? values : matrix.GetValue(r + matrix.GetLowerBound(0), c + matrix.GetLowerBound(1));
                                var error = value as ErrorWrapper ?? new ErrorWrapper(Convert.ToInt32(value));
                                if (matrix == null) values = error;
                                else matrix.SetValue(error, r + matrix.GetLowerBound(0), c + matrix.GetLowerBound(1));
                            }
                    }
                    finally { ComReference.Release(columns); ComReference.Release(rows); ComReference.Release(area); }
                }
            }
            finally { ComReference.Release(areas); ComReference.Release(errors); }
        }

        public void Paste()
        {
            var copied = snapshot;
            if (copied == null) throw new InvalidOperationException(captureError ?? "Nothing has been copied for Paste Formulas.");
            X.Range selection = null, cells = null, first = null, destination = null, rows = null, columns = null;
            X.Areas areas = null;
            X.Worksheet sheet = null;
            try
            {
                selection = application.Selection as X.Range;
                if (selection == null) throw new InvalidOperationException("Select a worksheet cell for Paste Formulas.");
                areas = selection.Areas;
                if (areas.Count != 1) throw new InvalidOperationException("Select a single destination area for Paste Formulas.");
                cells = selection.Cells; first = (X.Range)cells[1, 1]; sheet = first.Worksheet;
                rows = sheet.Rows; columns = sheet.Columns;
                copied.ValidateDestination(first.Row, first.Column, rows.Count, columns.Count);
                if (sheet.ProtectContents) throw new InvalidOperationException("Unprotect the destination worksheet before using Paste Formulas.");
                destination = first.Resize[copied.Rows, copied.Columns];
                if (!Equals(destination.MergeCells, false)) throw new InvalidOperationException("The destination must not contain merged cells.");
                if (!Equals(destination.HasArray, false)) throw new InvalidOperationException("The destination must not contain array formulas.");
                // Read all live appearance before touching destination: overlapping ranges
                // must not feed already-pasted formatting back into the source read.
                var appearance = TryReadSourceAppearance(copied);
                WriteContents(destination, copied.CreateFormulaAssignment());
                if (appearance != null) ApplyAppearance(destination, appearance);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("Excel could not insert the copied contents. Check that the destination is editable and does not intersect an array, spill or table restriction.", ex);
            }
            finally
            {
                ComReference.Release(destination); ComReference.Release(columns); ComReference.Release(rows);
                ComReference.Release(sheet); ComReference.Release(first); ComReference.Release(cells);
                ComReference.Release(areas); ComReference.Release(selection);
            }
        }

        private CellAppearanceSnapshot TryReadSourceAppearance(FormulaClipboardSnapshot copied)
        {
            X.Workbooks books = null;
            try
            {
                books = application.Workbooks;
                for (int i = 1; i <= books.Count; i++)
                {
                    X.Workbook book = books[i];
                    X.Sheets sheets = null; X.Worksheet sheet = null; X.Range source = null;
                    try
                    {
                        if (!string.Equals(book.FullName, copied.Workbook, StringComparison.OrdinalIgnoreCase)) continue;
                        sheets = book.Worksheets; sheet = sheets[copied.Worksheet] as X.Worksheet;
                        if (sheet == null) return null;
                        source = sheet.Range[copied.Address];
                        return CellAppearanceSnapshot.Read(source, copied.Rows, copied.Columns);
                    }
                    finally { ComReference.Release(source); ComReference.Release(sheet); ComReference.Release(sheets); ComReference.Release(book); }
                }
            }
            catch (Exception ex) { Debug.WriteLine("FormuRay: source appearance unavailable: " + ex.Message); }
            finally { ComReference.Release(books); }
            // Closed/renamed/deleted source: literal contents are still fully available.
            return null;
        }

        private void ApplyAppearance(X.Range destination, CellAppearanceSnapshot appearance)
        {
            bool events = application.EnableEvents, updating = application.ScreenUpdating;
            try
            {
                application.EnableEvents = false; application.ScreenUpdating = false;
                appearance.Apply(destination);
            }
            catch (Exception ex) { Debug.WriteLine("FormuRay: contents inserted; some formatting could not be applied: " + ex.Message); }
            finally
            {
                try { application.EnableEvents = events; }
                finally { application.ScreenUpdating = updating; }
            }
        }

        private sealed class NumberFormatBlock
        {
            internal X.Range Range;
            internal object Format;
        }

        private void WriteContents(X.Range destination, object[,] contents)
        {
            var formats = new List<NumberFormatBlock>();
            bool events = application.EnableEvents, updating = application.ScreenUpdating;
            try
            {
                ReadNumberFormats(destination, formats);
                application.EnableEvents = false; application.ScreenUpdating = false;
                try
                {
                    // Text-formatted destinations must still receive formulas. Restore every
                    // destination number format, including formats Excel auto-changes for dates.
                    destination.NumberFormat = "General";
                    destination.Formula = contents; // A1 strings unchanged; never PasteSpecial/R1C1.
                }
                finally
                {
                    foreach (var block in formats) block.Range.NumberFormat = block.Format;
                }
            }
            finally
            {
                foreach (var block in formats) ComReference.Release(block.Range);
                try { application.EnableEvents = events; }
                finally { application.ScreenUpdating = updating; }
            }
        }

        private static void ReadNumberFormats(X.Range range, List<NumberFormatBlock> result)
        {
            object format = range.NumberFormat;
            X.Range cells = null, first = null, block = null, rows = null, columns = null;
            try
            {
                cells = range.Cells; first = (X.Range)cells[1, 1];
                rows = range.Rows; columns = range.Columns;
                int height = rows.Count, width = columns.Count;
                if (format != null && format != DBNull.Value)
                {
                    block = first.Resize[height, width];
                    result.Add(new NumberFormatBlock { Range = block, Format = format }); block = null;
                    return;
                }
                // Most ranges need one format read; split only genuinely mixed formatting.
                bool vertical = height >= width;
                int half = (vertical ? height : width) / 2;
                if (half == 0) throw new InvalidOperationException("The destination number format could not be read.");
                block = first.Resize[vertical ? half : height, vertical ? width : half];
                ReadNumberFormats(block, result); ComReference.Release(block); block = null;
                ComReference.Release(first); first = (X.Range)cells[vertical ? half + 1 : 1, vertical ? 1 : half + 1];
                block = first.Resize[vertical ? height - half : height, vertical ? width : width - half];
                ReadNumberFormats(block, result);
            }
            finally
            {
                ComReference.Release(block); ComReference.Release(columns); ComReference.Release(rows);
                ComReference.Release(first); ComReference.Release(cells);
            }
        }
        public void Dispose() { Clear(); }
    }
}
