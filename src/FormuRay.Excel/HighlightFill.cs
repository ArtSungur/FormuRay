using System;
using System.Runtime.InteropServices;
using FormuRay.Core;
using X = Microsoft.Office.Interop.Excel;
namespace FormuRay.Excel
{
    internal static class HighlightFill
    {
        internal static int Background(X.Range cell)
        {
            X.DisplayFormat display = null; X.Interior fill = null;
            try
            {
                display = cell.DisplayFormat; fill = display.Interior;
                return Convert.ToInt32(fill.Pattern) == (int)X.XlPattern.xlPatternNone ? 0xFFFFFF : Convert.ToInt32(fill.Color);
            }
            catch (COMException)
            {
                ComReference.Release(fill); fill = cell.Interior;
                return Convert.ToInt32(fill.Pattern) == (int)X.XlPattern.xlPatternNone ? 0xFFFFFF : Convert.ToInt32(fill.Color);
            }
            finally { ComReference.Release(fill); ComReference.Release(display); }
        }
        internal static void Apply(X.Interior fill, int originalBackground, HighlightStyle style)
        {
            if (style.OpacityPercent == 0) return;
            fill.Pattern = X.XlPattern.xlPatternSolid;
            fill.Color = style.BlendOle(originalBackground);
            fill.TintAndShade = 0;
        }
    }
}
