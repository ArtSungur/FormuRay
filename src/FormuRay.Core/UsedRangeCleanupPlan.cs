using System;

namespace FormuRay.Core
{
    public sealed class CellRectangle
    {
        public int FirstRow { get; private set; }
        public int FirstColumn { get; private set; }
        public int LastRow { get; private set; }
        public int LastColumn { get; private set; }
        public CellRectangle(int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            FirstRow = firstRow; FirstColumn = firstColumn; LastRow = lastRow; LastColumn = lastColumn;
        }
    }

    public sealed class UsedRangeCleanupPlan
    {
        public const int MaxRow = 1048576;
        public const int MaxColumn = 16384;
        public CellRectangle Used { get; private set; }
        public int ContentRow { get; private set; }
        public int ContentColumn { get; private set; }
        public bool HasContent { get { return ContentRow != 0; } }
        public int RetainedRow { get; private set; }
        public int RetainedColumn { get; private set; }
        public int ExcessRows { get { return Math.Max(0, Used.LastRow - RetainedRow); } }
        public int ExcessColumns { get { return Math.Max(0, Used.LastColumn - RetainedColumn); } }
        public bool HasExcess { get { return ExcessRows > 0 || ExcessColumns > 0; } }
        public CellRectangle RowTail { get; private set; }
        public CellRectangle ColumnTail { get; private set; }

        public UsedRangeCleanupPlan(CellRectangle used, int contentRow, int contentColumn,
            int retainedRow, int retainedColumn)
        {
            if (used == null || used.FirstRow < 1 || used.FirstColumn < 1 ||
                used.LastRow < used.FirstRow || used.LastColumn < used.FirstColumn ||
                used.LastRow > MaxRow || used.LastColumn > MaxColumn)
                throw new ArgumentOutOfRangeException("used");
            if ((contentRow == 0) != (contentColumn == 0) || contentRow < 0 || contentRow > MaxRow ||
                contentColumn < 0 || contentColumn > MaxColumn)
                throw new ArgumentOutOfRangeException("contentRow");
            if (retainedRow < 1 || retainedRow > MaxRow ||
                retainedColumn < 1 || retainedColumn > MaxColumn)
                throw new ArgumentOutOfRangeException("retainedRow", "Retained boundary must remain within the worksheet.");
            Used = used; ContentRow = contentRow; ContentColumn = contentColumn;
            RetainedRow = retainedRow; RetainedColumn = retainedColumn;
            int rowStart = Math.Max(used.FirstRow, retainedRow + 1);
            if (rowStart <= used.LastRow)
                RowTail = new CellRectangle(rowStart, used.FirstColumn, used.LastRow, used.LastColumn);
            int columnStart = Math.Max(used.FirstColumn, retainedColumn + 1);
            int columnEndRow = Math.Min(used.LastRow, retainedRow);
            if (columnStart <= used.LastColumn && used.FirstRow <= columnEndRow)
                ColumnTail = new CellRectangle(used.FirstRow, columnStart, columnEndRow, used.LastColumn);
        }

        public static UsedRangeCleanupPlan Detect(CellRectangle used, int contentRow, int contentColumn)
        {
            return new UsedRangeCleanupPlan(used, contentRow, contentColumn,
                Math.Max(1, contentRow), Math.Max(1, contentColumn));
        }
    }
}
