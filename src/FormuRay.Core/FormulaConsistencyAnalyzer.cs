using System;
using System.Collections.Generic;
using System.Linq;

namespace FormuRay.Core
{
    public sealed class FormulaPatternCell
    {
        public int Row { get; private set; }
        public int Column { get; private set; }
        public string Pattern { get; private set; }

        public FormulaPatternCell(int row, int column, string pattern)
        {
            if (row < 1) throw new ArgumentOutOfRangeException("row");
            if (column < 1) throw new ArgumentOutOfRangeException("column");
            if (string.IsNullOrEmpty(pattern)) throw new ArgumentException("Formula pattern is required.", "pattern");
            Row = row;
            Column = column;
            Pattern = pattern;
        }
    }

    public static class FormulaConsistencyAnalyzer
    {
        private sealed class PatternRun
        {
            internal int Start;
            internal int Length;
            internal string Pattern;
        }

        public static IList<FormulaPatternCell> FindChanges(IEnumerable<FormulaPatternCell> formulaCells)
        {
            if (formulaCells == null) throw new ArgumentNullException("formulaCells");
            var cells = formulaCells.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToList();
            var changes = new List<FormulaPatternCell>();
            int start = 0;
            while (start < cells.Count)
            {
                int end = start + 1;
                while (end < cells.Count && cells[end].Row == cells[end - 1].Row
                    && cells[end].Column == cells[end - 1].Column + 1)
                    end++;
                if (end - start >= 2) DetectSegment(cells, start, end, changes);
                start = end;
            }
            return changes;
        }

        private static void DetectSegment(IList<FormulaPatternCell> cells, int start, int end,
            IList<FormulaPatternCell> changes)
        {
            var runs = new List<PatternRun>();
            for (int index = start; index < end; index++)
            {
                string pattern = cells[index].Pattern;
                if (runs.Count == 0 || !StringComparer.Ordinal.Equals(runs[runs.Count - 1].Pattern, pattern))
                    runs.Add(new PatternRun { Start = index, Length = 1, Pattern = pattern });
                else
                    runs[runs.Count - 1].Length++;
            }
            // The first run is the segment baseline. A -> X -> A marks X only.
            for (int index = 1; index < runs.Count; index++)
            {
                if (index >= 2 && runs[index - 1].Length == 1
                    && StringComparer.Ordinal.Equals(runs[index].Pattern, runs[index - 2].Pattern))
                    continue;
                int change = runs[index].Start;
                if (cells[change].Row == cells[change - 1].Row
                    && cells[change].Column == cells[change - 1].Column + 1)
                    changes.Add(cells[change]);
            }
        }
    }
}
