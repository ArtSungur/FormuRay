using System;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel
{
    // Short-lived managed appearance snapshot read at Paste, before ANY destination
    // writes (including overlapping source/destination). No clipboard, Range.Copy,
    // PasteSpecial, styles, conditional rules or long-lived COM references.
    internal sealed class CellAppearanceSnapshot
    {
        private static readonly X.XlBordersIndex[] Edges = { X.XlBordersIndex.xlEdgeLeft,
            X.XlBordersIndex.xlEdgeTop, X.XlBordersIndex.xlEdgeRight, X.XlBordersIndex.xlEdgeBottom,
            X.XlBordersIndex.xlDiagonalDown, X.XlBordersIndex.xlDiagonalUp };
        private readonly Cell[,] cells;
        private CellAppearanceSnapshot(int rows, int columns) { cells = new Cell[rows, columns]; }

        internal static CellAppearanceSnapshot Read(X.Range source, int rows, int columns)
        {
            var result = new CellAppearanceSnapshot(rows, columns);
            X.Range sourceCells = source.Cells;
            try
            {
                // Excel has no bulk array getter for Font/Interior/Borders. Keep these
                // per-cell reads out of the Ctrl+C hook; contents still use bulk reads.
                for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++)
                {
                    X.Range cell = (X.Range)sourceCells[r + 1, c + 1];
                    try { result.cells[r, c] = Cell.Read(cell); }
                    finally { ComReference.Release(cell); }
                }
                return result;
            }
            finally { ComReference.Release(sourceCells); }
        }

        internal void Apply(X.Range destination)
        {
            X.Range targetCells = destination.Cells;
            try
            {
                for (int r = 0; r < cells.GetLength(0); r++) for (int c = 0; c < cells.GetLength(1); c++)
                {
                    X.Range cell = (X.Range)targetCells[r + 1, c + 1];
                    try { cells[r, c].Apply(cell); }
                    finally { ComReference.Release(cell); }
                }
            }
            finally { ComReference.Release(targetCells); }
        }

        private sealed class Cell
        {
            private object numberFormat, horizontal, vertical, wrap, orientation, indent, shrink, locked, hidden, readingOrder;
            private object fontName, fontSize, bold, italic, underline, strike, subscript, superscript, fontColor, fontColorIndex;
            private object pattern, fillColor, patternColor;
            private Border[] borders;
            private Gradient gradient;

            internal static Cell Read(X.Range cell)
            {
                var result = new Cell {
                    numberFormat = cell.NumberFormat, horizontal = cell.HorizontalAlignment, vertical = cell.VerticalAlignment,
                    wrap = cell.WrapText, orientation = cell.Orientation, indent = cell.IndentLevel, shrink = cell.ShrinkToFit,
                    locked = cell.Locked, hidden = cell.FormulaHidden, readingOrder = cell.ReadingOrder
                };
                X.Font font = null; X.Interior fill = null; X.Borders edges = null;
                try
                {
                    font = cell.Font;
                    result.fontName = font.Name; result.fontSize = font.Size; result.bold = font.Bold;
                    result.italic = font.Italic; result.underline = font.Underline; result.strike = font.Strikethrough;
                    result.subscript = font.Subscript; result.superscript = font.Superscript;
                    result.fontColor = font.Color; result.fontColorIndex = font.ColorIndex;
                    fill = cell.Interior;
                    result.pattern = fill.Pattern;
                    int patternValue = Convert.ToInt32(result.pattern);
                    if (patternValue == (int)X.XlPattern.xlPatternLinearGradient || patternValue == (int)X.XlPattern.xlPatternRectangularGradient)
                        result.gradient = Gradient.Read(fill, patternValue == (int)X.XlPattern.xlPatternLinearGradient);
                    else { result.fillColor = fill.Color; result.patternColor = fill.PatternColor; }
                    edges = cell.Borders; result.borders = new Border[Edges.Length];
                    for (int i = 0; i < Edges.Length; i++)
                    {
                        X.Border edge = edges[Edges[i]];
                        try { result.borders[i] = new Border { Line = edge.LineStyle, Weight = edge.Weight, Color = edge.Color, ColorIndex = edge.ColorIndex }; }
                        finally { ComReference.Release(edge); }
                    }
                    return result;
                }
                finally { ComReference.Release(edges); ComReference.Release(fill); ComReference.Release(font); }
            }

            internal void Apply(X.Range cell)
            {
                cell.NumberFormat = numberFormat;
                cell.HorizontalAlignment = horizontal; cell.VerticalAlignment = vertical;
                cell.WrapText = wrap; cell.Orientation = orientation; cell.IndentLevel = indent; cell.ShrinkToFit = shrink;
                cell.Locked = locked; cell.FormulaHidden = hidden; cell.ReadingOrder = Convert.ToInt32(readingOrder);
                X.Font font = null; X.Interior fill = null; X.Borders edges = null;
                try
                {
                    font = cell.Font;
                    // Mixed rich-text attributes may be DBNull. Do not assign DBNull to Excel.
                    if (Present(fontName)) font.Name = fontName;
                    if (Present(fontSize)) font.Size = fontSize;
                    if (Present(bold)) font.Bold = bold;
                    if (Present(italic)) font.Italic = italic;
                    if (Present(underline)) font.Underline = underline;
                    if (Present(strike)) font.Strikethrough = strike;
                    if (Present(subscript)) font.Subscript = subscript;
                    if (Present(superscript)) font.Superscript = superscript;
                    // Resolved RGB preserves appearance across workbooks with different themes.
                    // Color already includes tint, so applying the source tint again is wrong.
                    if (Present(fontColor)) { font.Color = fontColor; font.TintAndShade = 0; }
                    if (Equals(fontColorIndex, (int)X.XlColorIndex.xlColorIndexAutomatic)) font.ColorIndex = fontColorIndex;
                    fill = cell.Interior;
                    if (gradient != null) { fill.Pattern = pattern; gradient.Apply(fill); }
                    else
                    {
                        fill.Color = fillColor; fill.TintAndShade = 0;
                        fill.PatternColor = patternColor; fill.PatternTintAndShade = 0;
                        fill.Pattern = pattern; // xlPatternNone must remain truly unfilled.
                    }
                    edges = cell.Borders;
                    for (int i = 0; i < Edges.Length; i++)
                    {
                        X.Border edge = edges[Edges[i]];
                        try
                        {
                            var data = borders[i];
                            if (Convert.ToInt32(data.Line) != (int)X.XlLineStyle.xlLineStyleNone)
                            {
                                edge.Weight = data.Weight; edge.Color = data.Color; edge.TintAndShade = 0;
                                if (Equals(data.ColorIndex, (int)X.XlColorIndex.xlColorIndexAutomatic)) edge.ColorIndex = data.ColorIndex;
                            }
                            edge.LineStyle = data.Line;
                        }
                        finally { ComReference.Release(edge); }
                    }
                }
                finally { ComReference.Release(edges); ComReference.Release(fill); ComReference.Release(font); }
            }
            private static bool Present(object value) { return value != null && value != DBNull.Value; }
        }
        private sealed class Border { internal object Line, Weight, Color, ColorIndex; }

        private sealed class Gradient
        {
            private bool linear;
            private double degree, left, right, top, bottom;
            private double[] positions, tints;
            private int[] colors;
            internal static Gradient Read(X.Interior fill, bool linear)
            {
                var result = new Gradient { linear = linear };
                dynamic gradient = fill.Gradient;
                X.ColorStops stops = null;
                try
                {
                    if (linear) result.degree = gradient.Degree;
                    else { result.left = gradient.RectangleLeft; result.right = gradient.RectangleRight; result.top = gradient.RectangleTop; result.bottom = gradient.RectangleBottom; }
                    stops = gradient.ColorStops;
                    int count = Convert.ToInt32(stops.Count);
                    result.positions = new double[count]; result.colors = new int[count]; result.tints = new double[count];
                    for (int i = 0; i < count; i++)
                    {
                        X.ColorStop stop = stops[i + 1];
                        try { result.positions[i] = Convert.ToDouble(stop.Position); result.colors[i] = Convert.ToInt32(stop.Color); result.tints[i] = Convert.ToDouble(stop.TintAndShade); }
                        finally { ComReference.Release(stop); }
                    }
                    return result;
                }
                finally { ComReference.Release(stops); ComReference.Release((object)gradient); }
            }
            internal void Apply(X.Interior fill)
            {
                dynamic gradient = fill.Gradient;
                X.ColorStops stops = null;
                try
                {
                    if (linear) gradient.Degree = degree;
                    else { gradient.RectangleLeft = left; gradient.RectangleRight = right; gradient.RectangleTop = top; gradient.RectangleBottom = bottom; }
                    stops = gradient.ColorStops; stops.Clear();
                    for (int i = 0; i < positions.Length; i++)
                    {
                        X.ColorStop stop = stops.Add(positions[i]);
                        try { stop.Color = colors[i]; stop.TintAndShade = tints[i]; }
                        finally { ComReference.Release(stop); }
                    }
                }
                finally { ComReference.Release(stops); ComReference.Release((object)gradient); }
            }
        }
    }
}
