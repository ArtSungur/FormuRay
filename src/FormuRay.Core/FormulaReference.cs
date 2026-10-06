namespace FormuRay.Core
{
    public sealed class FormulaReference
    {
        public string Text { get; set; }
        public string Sheet { get; set; }
        public string Workbook { get; set; }
        public string Address { get; set; }
        // All C# spans are zero-based (the VBA export uses one-based indices).
        public int Start { get; set; }
        public int Occurrence { get; set; }
        public bool IsExternal { get { return Text.Contains("["); } }
        public bool IsWholeRowOrColumn { get { return FormulaReferenceParser.IsWholeRowOrColumn(Address); } }
        public string Key { get { return Sheet.ToUpperInvariant() + "!" + Address.Replace("$", "").ToUpperInvariant(); } }
        public string Display { get { return string.IsNullOrEmpty(Sheet) ? Address : FormulaReferenceParser.QuoteSheet(Sheet) + "!" + Address; } }
    }
}
