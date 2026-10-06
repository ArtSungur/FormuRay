using System.Collections.Generic;

namespace FormuRay.Core
{
    public enum ExplorerMode { Precedents, Dependents }

    // Serializable locations keep the UI independent of COM and do not keep a closed book alive.
    public sealed class RangeLocation
    {
        public string Workbook { get; set; }
        public string WorkbookFullName { get; set; }
        public string Sheet { get; set; }
        public string Address { get; set; }
        public string Key { get { return WorkbookFullName + "|" + Sheet + "|" + Address.Replace("$", ""); } }
    }

    public sealed class ExplorerComponent
    {
        public string Kind { get; set; }
        public string Display { get; set; }
        public string Value { get; set; }
        public string Reference { get; set; }
        public string MatchedReference { get; set; }
        public string Formula { get; set; }
        public string FormulaLocal { get; set; }
        public string SafetyKind { get; set; }
        public RangeLocation Target { get; set; }
        public int Start { get; set; }
        public int Length { get; set; }
        public int Occurrence { get; set; }
        public bool NonNavigable { get; set; }
    }

    public sealed class ExplorerDocument
    {
        public ExplorerMode Mode { get; set; }
        public RangeLocation Source { get; set; }
        public string Formula { get; set; }
        public string FormulaLocal { get; set; }
        public string Status { get; set; }
        public ParsedFormula Expression { get; set; }
        public string LocalDecimalSeparator { get; set; }
        public List<ExplorerComponent> Components { get; private set; }
        public ExplorerDocument() { Components = new List<ExplorerComponent>(); Formula = ""; FormulaLocal = ""; }
        public string DisplayFormula { get { return string.IsNullOrEmpty(FormulaLocal) ? Formula : FormulaLocal; } }
    }
}
