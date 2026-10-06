using System;
using System.Collections.Generic;

namespace FormuRay.Core
{
    public enum AutoFormatCellType { InputPercent, InputValue, InputDate, FormulaPercent, FormulaValue, FormulaDate }
    public enum AutoFormatHorizontal { Unchanged, General, Left, Center, Right }
    public enum AutoFormatVertical { Unchanged, Top, Center, Bottom }
    public enum AutoFormatBorder { Unchanged, None, Outline, Bottom, Top, TopAndBottom, All }
    public enum AutoFormatLineStyle { Continuous, Dash, Dot, Double }
    public enum AutoFormatBorderWeight { Thin, Medium, Thick }

    public sealed class AutoFormatStyle
    {
        public int FillRgb { get; private set; }
        public int FontRgb { get; private set; }
        public bool Bold { get; private set; }
        public bool Italic { get; private set; }
        public string NumberFormat { get; private set; }
        public AutoFormatHorizontal Horizontal { get; private set; }
        public AutoFormatVertical Vertical { get; private set; }
        public AutoFormatBorder Border { get; private set; }
        public int BorderRgb { get; private set; }
        public AutoFormatLineStyle BorderLineStyle { get; private set; }
        public AutoFormatBorderWeight BorderWeight { get; private set; }

        public AutoFormatStyle(int fillRgb, int fontRgb, bool bold, bool italic, string numberFormat,
            AutoFormatHorizontal horizontal, AutoFormatVertical vertical, AutoFormatBorder border,
            int borderRgb, AutoFormatLineStyle borderLineStyle, AutoFormatBorderWeight borderWeight)
        {
            if (fillRgb < 0 || fillRgb > 0xFFFFFF || fontRgb < 0 || fontRgb > 0xFFFFFF ||
                borderRgb < 0 || borderRgb > 0xFFFFFF) throw new ArgumentOutOfRangeException("color");
            if (string.IsNullOrWhiteSpace(numberFormat)) throw new ArgumentException("Number format is required.", "numberFormat");
            if (!Enum.IsDefined(typeof(AutoFormatHorizontal), horizontal) || !Enum.IsDefined(typeof(AutoFormatVertical), vertical)
                || !Enum.IsDefined(typeof(AutoFormatBorder), border) || !Enum.IsDefined(typeof(AutoFormatLineStyle), borderLineStyle)
                || !Enum.IsDefined(typeof(AutoFormatBorderWeight), borderWeight)) throw new ArgumentOutOfRangeException("style");
            FillRgb = fillRgb; FontRgb = fontRgb; Bold = bold; Italic = italic; NumberFormat = numberFormat;
            Horizontal = horizontal; Vertical = vertical; Border = border; BorderRgb = borderRgb;
            BorderLineStyle = borderLineStyle; BorderWeight = borderWeight;
        }
    }

    public sealed class AutoFormatSettings
    {
        private static readonly AutoFormatCellType[] kinds = (AutoFormatCellType[])Enum.GetValues(typeof(AutoFormatCellType));
        private readonly AutoFormatStyle[] styles;
        public static IEnumerable<AutoFormatCellType> Types { get { return kinds; } }
        public AutoFormatSettings(IEnumerable<AutoFormatStyle> styles)
        {
            this.styles = new List<AutoFormatStyle>(styles).ToArray();
            if (this.styles.Length != kinds.Length) throw new ArgumentException("Six styles are required.", "styles");
            foreach (var style in this.styles) if (style == null) throw new ArgumentException("A style is missing.", "styles");
        }
        public AutoFormatStyle Get(AutoFormatCellType type)
        {
            int index = (int)type;
            if (index < 0 || index >= styles.Length) throw new ArgumentOutOfRangeException("type");
            return styles[index];
        }
        public AutoFormatSettings With(AutoFormatCellType type, AutoFormatStyle style)
        {
            if (style == null) throw new ArgumentNullException("style");
            if ((int)type < 0 || (int)type >= styles.Length) throw new ArgumentOutOfRangeException("type");
            var copy = (AutoFormatStyle[])styles.Clone();
            copy[(int)type] = style;
            return new AutoFormatSettings(copy);
        }
        public static AutoFormatSettings Defaults
        {
            get
            {
                return new AutoFormatSettings(new[] {
                    Default(0xE7F0F7, "0.0%"), Default(0xEDF3F7, "#,##0.00;(#,##0.00)"),
                    Default(0xE7F0F7, "yyyy-mm-dd"), Default(0xECEFF5, "0.0%"),
                    Default(0xF1F3F6, "#,##0.00;(#,##0.00)"), Default(0xECEFF5, "yyyy-mm-dd") });
            }
        }
        private static AutoFormatStyle Default(int fill, string numberFormat)
        {
            return new AutoFormatStyle(fill, 0x253246, false, false, numberFormat,
                AutoFormatHorizontal.Right, AutoFormatVertical.Center, AutoFormatBorder.Unchanged,
                0xB8C4D0, AutoFormatLineStyle.Continuous, AutoFormatBorderWeight.Thin);
        }
    }
}
