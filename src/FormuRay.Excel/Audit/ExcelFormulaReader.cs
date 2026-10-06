using System;
using System.Runtime.InteropServices;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel.Audit
{
    internal static class ExcelFormulaReader
    {
        internal static void Visit(X.Worksheet sheet, Func<string, bool> visit)
        {
            VisitCells(sheet, (formula, row, column) => visit(formula), false);
        }

        internal static void VisitCells(X.Worksheet sheet, Func<string, int, int, bool> visit, bool coordinates = true)
        {
            X.Range formulas;
            try { formulas = sheet.UsedRange.SpecialCells(X.XlCellType.xlCellTypeFormulas); }
            catch (COMException) { return; }
            foreach (X.Range area in formulas.Areas)
            {
                int rows = area.Rows.Count, columns = area.Columns.Count;
                int firstRow = coordinates ? area.Row : 0, firstColumn = coordinates ? area.Column : 0;
                int blockRows = Math.Max(1, 4096 / columns);
                for (int row = 1; row <= rows; row += blockRows)
                {
                    X.Range block = ((X.Range)area.Cells[row, 1]).Resize[Math.Min(blockRows, rows - row + 1), columns];
                    object values;
                    try { values = block.Formula; }
                    catch (COMException)
                    {
                        foreach (X.Range cell in block.Cells)
                        {
                            string text;
                            try { text = Convert.ToString(cell.Formula); } catch (COMException) { continue; }
                            if (!visit(text, coordinates ? cell.Row : 0, coordinates ? cell.Column : 0)) return;
                        }
                        continue;
                    }
                    var matrix = values as object[,];
                    if (matrix == null) { if (!visit(Convert.ToString(values), firstRow + row - 1, firstColumn)) return; }
                    else
                    {
                        for (int r = matrix.GetLowerBound(0); r <= matrix.GetUpperBound(0); r++)
                            for (int c = matrix.GetLowerBound(1); c <= matrix.GetUpperBound(1); c++)
                                if (!visit(Convert.ToString(matrix[r, c]), firstRow + row - 1 + r - matrix.GetLowerBound(0),
                                    firstColumn + c - matrix.GetLowerBound(1))) return;
                    }
                }
            }
        }
    }
}
