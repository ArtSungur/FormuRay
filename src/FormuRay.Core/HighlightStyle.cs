using System;
namespace FormuRay.Core
{
    public sealed class HighlightStyle
    {
        public int Rgb { get; private set; }
        public int OpacityPercent { get; private set; }
        public HighlightStyle(int rgb, int opacityPercent)
        {
            if (rgb < 0 || rgb > 0xFFFFFF) throw new ArgumentOutOfRangeException("rgb");
            if (opacityPercent < 0 || opacityPercent > 100) throw new ArgumentOutOfRangeException("opacityPercent");
            Rgb = rgb; OpacityPercent = opacityPercent;
        }
        // RGB is stored as #RRGGBB. Excel/OLE uses 0x00BBGGRR, never ARGB.
        public int BlendOle(int background)
        {
            int result = 0;
            for (int shift = 0; shift <= 16; shift += 8)
            {
                int original = (background >> shift) & 255;
                int highlight = (Rgb >> (16 - shift)) & 255;
                result |= ((original * (100 - OpacityPercent) + highlight * OpacityPercent + 50) / 100) << shift;
            }
            return result;
        }
    }
    public sealed class HighlightSettings
    {
        public HighlightStyle HardCode { get; private set; }
        public HighlightStyle FormulaConsistency { get; private set; }
        public HighlightStyle Inflow { get; private set; }
        public HighlightStyle Outflow { get; private set; }
        public HighlightSettings(HighlightStyle hardCode, HighlightStyle inflow, HighlightStyle outflow,
            HighlightStyle formulaConsistency = null)
        {
            HardCode = hardCode ?? Defaults.HardCode; Inflow = inflow ?? Defaults.Inflow; Outflow = outflow ?? Defaults.Outflow;
            FormulaConsistency = formulaConsistency ?? Defaults.FormulaConsistency;
        }
        // Keep the established appearance: prior fills were fully opaque.
        public static HighlightSettings Defaults { get { return new HighlightSettings(new HighlightStyle(0xFFC7CE, 100),
            new HighlightStyle(0xD1EBDC, 100), new HighlightStyle(0xCFE2F3, 100), new HighlightStyle(0xF2D49B, 85)); } }
    }
}
