using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using FormuRay.Core;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel.Audit
{
    internal sealed class AuditFillSession : IDisposable
    {
        internal const int Magenta = 255 + (255 << 16);
        internal const int Green = 198 + (239 << 8) + (206 << 16);
        private sealed class Snapshot
        {
            internal X.Range Cell;
            internal object Pattern, Color, ColorIndex, Tint;
            internal object Theme, PatternColor, PatternColorIndex, PatternTheme, PatternTint;
            internal int Background;
        }
        private readonly List<Snapshot> states = new List<Snapshot>();
        private readonly Dictionary<string, Snapshot> keys = new Dictionary<string, Snapshot>(StringComparer.OrdinalIgnoreCase);
        private readonly bool preserveDetailedFill;
        internal AuditFillSession(bool preserveDetailedFill = false) { this.preserveDetailedFill = preserveDetailedFill; }

        internal bool Apply(X.Range range, int color, int maximumCells, bool includeTint)
        { return Apply(range, color, maximumCells, includeTint, null); }
        internal bool Apply(X.Range range, HighlightStyle style, int maximumCells, bool includeTint)
        { return Apply(range, 0, maximumCells, includeTint, style); }
        private bool Apply(X.Range range, int color, int maximumCells, bool includeTint, HighlightStyle style)
        {
            if (range == null || Convert.ToDouble(range.CountLarge) > maximumCells) return false;
            if (style != null && style.OpacityPercent == 0) return true;
            var sheet = range.Worksheet;
            if (range.Rows.Count == sheet.Rows.Count || range.Columns.Count == sheet.Columns.Count) return false;
            string prefix = ExcelAuditContext.Location(range).Key;
            prefix = prefix.Substring(0, prefix.LastIndexOf('|') + 1);
            try
            {
                // Only these cell RCWs are retained, specifically to restore formatting across renames.
                // They are released once when the highlight session ends, not force-released globally.
                foreach (X.Range cell in range.Cells)
                {
                    bool retained = false;
                    X.Interior fill = null;
                    try
                    {
                        fill = cell.Interior;
                        string key = prefix + cell.Address[false, false, X.XlReferenceStyle.xlA1];
                        if (!keys.ContainsKey(key))
                        {
                            var snapshot = new Snapshot { Cell = cell, Pattern = fill.Pattern, Color = fill.Color, ColorIndex = fill.ColorIndex };
                            if (style != null) snapshot.Background = HighlightFill.Background(cell);
                            if (preserveDetailedFill)
                            {
                                // Gradient stops are not represented by Interior.Color. Leave such cells untouched.
                                int pattern = Convert.ToInt32(snapshot.Pattern);
                                if (pattern == (int)X.XlPattern.xlPatternLinearGradient || pattern == (int)X.XlPattern.xlPatternRectangularGradient) return false;
                                snapshot.PatternColor = fill.PatternColor;
                                snapshot.PatternColorIndex = fill.PatternColorIndex;
                                snapshot.PatternTint = fill.PatternTintAndShade;
                                try { snapshot.Theme = fill.ThemeColor; } catch (COMException) { }
                                try { snapshot.PatternTheme = fill.PatternThemeColor; } catch (COMException) { }
                            }
                            if (includeTint) { try { snapshot.Tint = fill.TintAndShade; } catch (COMException) { } }
                            if (preserveDetailedFill)
                            {
                                // Color getters include tint; read the base RGB to avoid applying tint twice on restore.
                                if ((snapshot.Theme == null || Convert.ToInt32(snapshot.Theme) <= 0)
                                    && snapshot.Tint != null && Convert.ToDouble(snapshot.Tint) != 0)
                                {
                                    try { fill.TintAndShade = 0; snapshot.Color = fill.Color; }
                                    finally { fill.TintAndShade = snapshot.Tint; }
                                }
                                if ((snapshot.PatternTheme == null || Convert.ToInt32(snapshot.PatternTheme) <= 0)
                                    && snapshot.PatternTint != null && Convert.ToDouble(snapshot.PatternTint) != 0)
                                {
                                    try { fill.PatternTintAndShade = 0; snapshot.PatternColor = fill.PatternColor; }
                                    finally { fill.PatternTintAndShade = snapshot.PatternTint; }
                                }
                            }
                            states.Add(snapshot);
                            keys.Add(key, snapshot);
                            retained = true;
                        }
                        if (style != null) HighlightFill.Apply(fill, keys[key].Background, style);
                        else
                        {
                            fill.Pattern = X.XlPattern.xlPatternSolid;
                            fill.Color = color;
                            if (includeTint) { try { fill.TintAndShade = 0; } catch (COMException) { } }
                        }
                    }
                    finally { ComReference.Release(fill); if (!retained) ComReference.Release(cell); }
                }
                return true;
            }
            catch (COMException ex) { Debug.WriteLine("FormuRay fill: " + ex.Message); return false; }
        }

        internal void Clear()
        {
            foreach (var snapshot in states)
            {
                X.Interior fill = null;
                try
                {
                    fill = snapshot.Cell.Interior;
                    if (preserveDetailedFill)
                    {
                        // Inflow restores RGB/theme/pattern metadata. Existing audit callers retain their exact legacy path below.
                        if (snapshot.Theme != null && Convert.ToInt32(snapshot.Theme) > 0) fill.ThemeColor = snapshot.Theme;
                        else fill.Color = snapshot.Color;
                        if (Convert.ToInt32(snapshot.ColorIndex) < 0) fill.ColorIndex = snapshot.ColorIndex;
                        if (snapshot.Tint != null) fill.TintAndShade = snapshot.Tint;
                        if (snapshot.PatternTheme != null && Convert.ToInt32(snapshot.PatternTheme) > 0) fill.PatternThemeColor = snapshot.PatternTheme;
                        else fill.PatternColor = snapshot.PatternColor;
                        if (Convert.ToInt32(snapshot.PatternColorIndex) < 0) fill.PatternColorIndex = snapshot.PatternColorIndex;
                        if (snapshot.PatternTint != null) fill.PatternTintAndShade = snapshot.PatternTint;
                        fill.Pattern = snapshot.Pattern;
                        continue;
                    }
                    fill.Pattern = snapshot.Pattern;
                    fill.Color = snapshot.Color;
                    fill.ColorIndex = snapshot.ColorIndex;
                    if (snapshot.Tint != null) fill.TintAndShade = snapshot.Tint;
                }
                catch (COMException ex) { Debug.WriteLine("FormuRay restore: " + ex.Message); }
                catch (InvalidComObjectException) { }
                finally { ComReference.Release(fill); ComReference.Release(snapshot.Cell); }
            }
            states.Clear();
            keys.Clear();
        }
        public void Dispose() { Clear(); }
    }
}
