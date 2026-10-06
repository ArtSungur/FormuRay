namespace FormuRay.Core
{
    // Logical pixels at 96 DPI; persistence is independent of the current monitor.
    public sealed class ExplorerColumnWidths
    {
        public int Element { get; private set; }
        public int Value { get; private set; }
        public int Type { get; private set; }
        public double SplitterRatio { get; private set; }
        public static ExplorerColumnWidths Defaults { get { return new ExplorerColumnWidths(390, 160, 120); } }
        public ExplorerColumnWidths(int element, int value, int type, double splitterRatio = 0.7)
        {
            Element = Valid(element, 390); Value = Valid(value, 160); Type = Valid(type, 120);
            SplitterRatio = !double.IsNaN(splitterRatio) && splitterRatio >= .15 && splitterRatio <= .9 ? splitterRatio : .7;
        }
        private static int Valid(int value, int fallback) { return value >= 60 && value <= 3000 ? value : fallback; }
    }
}
