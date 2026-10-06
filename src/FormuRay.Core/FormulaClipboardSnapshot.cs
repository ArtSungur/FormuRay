using System;

namespace FormuRay.Core
{
    // Entirely managed, immutable copy-time contents; no workbook/Range lifetime dependency.
    public sealed class FormulaClipboardSnapshot
    {
        private readonly object[,] contents;
        private readonly bool[,] formulas;
        public string Workbook { get; private set; }
        public string Worksheet { get; private set; }
        public string Address { get; private set; }
        public int Rows { get { return contents.GetLength(0); } }
        public int Columns { get { return contents.GetLength(1); } }

        public FormulaClipboardSnapshot(string workbook, string worksheet, string address,
            int rows, int columns, object formulaData, object valueData, bool[,] formulaMask)
        {
            if (rows < 1 || columns < 1 || formulaMask == null
                || formulaMask.GetLength(0) != rows || formulaMask.GetLength(1) != columns)
                throw new ArgumentException("Invalid copied range dimensions.");
            Workbook = workbook; Worksheet = worksheet; Address = address;
            contents = new object[rows, columns];
            formulas = (bool[,])formulaMask.Clone();
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < columns; c++)
                {
                    object value = Read(formulas[r, c] ? formulaData : valueData, r, c, rows, columns);
                    if (formulas[r, c] && (!(value is string) || !((string)value).StartsWith("=", StringComparison.Ordinal)))
                        throw new ArgumentException("A formula cell must contain an A1 formula string.");
                    contents[r, c] = value == DBNull.Value ? null : value;
                }
        }

        private static object Read(object data, int row, int column, int rows, int columns)
        {
            var matrix = data as Array;
            if (matrix == null)
            {
                if (rows != 1 || columns != 1) throw new ArgumentException("Expected a two-dimensional range array.");
                return data;
            }
            if (matrix.Rank != 2 || matrix.GetLength(0) != rows || matrix.GetLength(1) != columns)
                throw new ArgumentException("Copied range array dimensions do not match.");
            return matrix.GetValue(row + matrix.GetLowerBound(0), column + matrix.GetLowerBound(1));
        }

        public object ContentAt(int row, int column) { return contents[row, column]; }
        public bool IsFormula(int row, int column) { return formulas[row, column]; }

        public object[,] CreateFormulaAssignment()
        {
            var result = new object[Rows, Columns];
            for (int r = 0; r < Rows; r++)
                for (int c = 0; c < Columns; c++)
                {
                    object value = contents[r, c];
                    // Formula accepts constants too. Quote ALL text so =, +, numeric text,
                    // dates, apostrophes and empty strings cannot be reinterpreted by Excel.
                    result[r, c] = !formulas[r, c] && value is string ? "'" + (string)value : value;
                }
            return result;
        }

        public void ValidateDestination(int row, int column, int sheetRows, int sheetColumns)
        {
            if (row < 1 || column < 1 || Rows > sheetRows - row + 1 || Columns > sheetColumns - column + 1)
                throw new InvalidOperationException("The copied block would extend beyond the worksheet boundary.");
        }
    }
}
