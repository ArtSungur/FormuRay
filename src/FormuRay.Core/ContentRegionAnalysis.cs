using System;
using System.Collections.Generic;
using System.Linq;

namespace FormuRay.Core
{
    public enum CellContentKind { Constant, Formula }

    // Each rectangle comes from an Excel SpecialCells area, so every cell in it has content.
    public sealed class ContentArea
    {
        public CellRectangle Bounds { get; private set; }
        public CellContentKind Kind { get; private set; }
        public long Cells { get { return (long)(Bounds.LastRow - Bounds.FirstRow + 1) * (Bounds.LastColumn - Bounds.FirstColumn + 1); } }
        public ContentArea(CellRectangle bounds, CellContentKind kind)
        {
            if (bounds == null) throw new ArgumentNullException("bounds");
            Bounds = bounds; Kind = kind;
        }
    }

    public sealed class ContentBlock
    {
        internal readonly List<ContentArea> Areas = new List<ContentArea>();
        public CellRectangle Bounds { get; private set; }
        public long Cells { get; private set; }
        public long FormulaCells { get; private set; }
        public long ConstantCells { get { return Cells - FormulaCells; } }
        public string ContentDescription
        {
            get { return FormulaCells == 0 ? "Constants" : ConstantCells == 0 ? "Formulas" : "Constants/Formulas"; }
        }
        internal void Add(ContentArea area)
        {
            Areas.Add(area);
            Cells += area.Cells;
            if (area.Kind == CellContentKind.Formula) FormulaCells += area.Cells;
            if (Bounds == null) Bounds = area.Bounds;
            else Bounds = new CellRectangle(Math.Min(Bounds.FirstRow, area.Bounds.FirstRow),
                Math.Min(Bounds.FirstColumn, area.Bounds.FirstColumn),
                Math.Max(Bounds.LastRow, area.Bounds.LastRow), Math.Max(Bounds.LastColumn, area.Bounds.LastColumn));
        }
        internal void Merge(ContentBlock other)
        {
            foreach (ContentArea area in other.Areas) Add(area);
        }
    }

    public sealed class ContentRemovalPreview
    {
        public IList<ContentBlock> Blocks { get; private set; }
        public long Cells { get { return Blocks.Sum(b => b.Cells); } }
        public long FormulaCells { get { return Blocks.Sum(b => b.FormulaCells); } }
        public bool RequiresConfirmation { get { return Cells > 0; } }
        internal string Fingerprint { get; private set; }
        internal ContentRemovalPreview(IList<ContentBlock> blocks, string fingerprint)
        {
            Blocks = blocks; Fingerprint = fingerprint;
        }
        public bool IsAuthorized(bool confirmed) { return !RequiresConfirmation || confirmed; }
        public bool SameContent(ContentRemovalPreview other)
        {
            return other != null && Fingerprint == other.Fingerprint;
        }
    }

    public sealed class ContentRegionAnalysis
    {
        public const int MainRowBuffer = 10;
        public const int MainColumnBuffer = 3;
        private readonly List<ContentArea> source;
        public ContentBlock Main { get; private set; }
        public IList<ContentBlock> Distant { get; private set; }
        public IList<ContentBlock> All { get; private set; }

        public ContentRegionAnalysis(IEnumerable<ContentArea> areas)
        {
            source = areas.ToList();
            var blocks = Group(source);
            All = blocks;
            Main = blocks.OrderByDescending(Score).ThenBy(b => b.Bounds.FirstRow)
                .ThenBy(b => b.Bounds.FirstColumn).FirstOrDefault();
            Distant = blocks.Where(b => b != Main && !WithinMainBuffer(b.Bounds, Main.Bounds))
                .OrderBy(b => b.Bounds.FirstRow)
                .ThenBy(b => b.Bounds.FirstColumn).ToList();
        }

        private static bool WithinMainBuffer(CellRectangle candidate, CellRectangle main)
        {
            return candidate.FirstRow >= Math.Max(1, main.FirstRow - MainRowBuffer)
                && candidate.LastRow <= Math.Min(UsedRangeCleanupPlan.MaxRow, main.LastRow + MainRowBuffer)
                && candidate.FirstColumn >= Math.Max(1, main.FirstColumn - MainColumnBuffer)
                && candidate.LastColumn <= Math.Min(UsedRangeCleanupPlan.MaxColumn, main.LastColumn + MainColumnBuffer);
        }

        private static double Score(ContentBlock block)
        {
            // A substantial early model wins over a later, larger table; a stray early cell does not.
            return block.Cells / (1.0 + (block.Bounds.FirstRow - 1) / 100.0
                + (block.Bounds.FirstColumn - 1) / 20.0);
        }

        private static bool Near(CellRectangle a, CellRectangle b)
        {
            bool rowsOverlap = a.FirstRow <= b.LastRow && b.FirstRow <= a.LastRow;
            bool columnsOverlap = a.FirstColumn <= b.LastColumn && b.FirstColumn <= a.LastColumn;
            return (rowsOverlap && (long)a.FirstColumn <= (long)b.LastColumn + 2
                    && (long)b.FirstColumn <= (long)a.LastColumn + 2)
                || (columnsOverlap && (long)a.FirstRow <= (long)b.LastRow + 2
                    && (long)b.FirstRow <= (long)a.LastRow + 2);
        }

        private static List<ContentBlock> Group(IEnumerable<ContentArea> areas)
        {
            var blocks = new List<ContentBlock>();
            foreach (ContentArea area in areas)
            {
                var block = new ContentBlock(); block.Add(area);
                bool merged;
                do
                {
                    merged = false;
                    for (int i = blocks.Count - 1; i >= 0; i--)
                        if (Near(block.Bounds, blocks[i].Bounds))
                        {
                            block.Merge(blocks[i]); blocks.RemoveAt(i); merged = true;
                        }
                } while (merged);
                blocks.Add(block);
            }
            return blocks;
        }

        public ContentRemovalPreview RemovalFor(UsedRangeCleanupPlan plan)
        {
            var removedRows = new List<ContentArea>();
            var removedColumns = new List<ContentArea>();
            foreach (ContentArea area in source)
            {
                AddIntersection(removedRows, area, plan.RowTail);
                AddIntersection(removedColumns, area, plan.ColumnTail);
            }
            var removed = removedRows.Concat(removedColumns).ToList();
            string fingerprint = string.Join(";", removed.Select(a => string.Format("{0}:{1},{2},{3},{4}",
                (int)a.Kind, a.Bounds.FirstRow, a.Bounds.FirstColumn, a.Bounds.LastRow, a.Bounds.LastColumn))
                .OrderBy(s => s, StringComparer.Ordinal));
            return new ContentRemovalPreview(Group(removedRows).Concat(Group(removedColumns)).OrderBy(b => b.Bounds.FirstRow)
                .ThenBy(b => b.Bounds.FirstColumn).ToList(), fingerprint);
        }

        private static void AddIntersection(List<ContentArea> result, ContentArea area, CellRectangle tail)
        {
            if (tail == null) return;
            int firstRow = Math.Max(area.Bounds.FirstRow, tail.FirstRow);
            int firstColumn = Math.Max(area.Bounds.FirstColumn, tail.FirstColumn);
            int lastRow = Math.Min(area.Bounds.LastRow, tail.LastRow);
            int lastColumn = Math.Min(area.Bounds.LastColumn, tail.LastColumn);
            if (firstRow <= lastRow && firstColumn <= lastColumn)
                result.Add(new ContentArea(new CellRectangle(firstRow, firstColumn, lastRow, lastColumn), area.Kind));
        }
    }
}
