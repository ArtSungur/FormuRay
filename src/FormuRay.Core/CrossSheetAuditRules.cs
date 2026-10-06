using System;

namespace FormuRay.Core
{
    public static class CrossSheetAuditRules
    {
        public const int MaximumRangeCells = 5000;
        public const int MaximumHighlightedRanges = 10000;
        public static bool CouldReferenceSheet(string formula, string sheet)
        {
            return formula.IndexOf(sheet + "!", StringComparison.OrdinalIgnoreCase) >= 0
                || formula.IndexOf(FormulaReferenceParser.QuoteSheet(sheet) + "!", StringComparison.OrdinalIgnoreCase) >= 0;
        }
        public static bool Accepts(FormulaReference reference, string sourceSheet, string formula)
        {
            return !reference.IsExternal && !reference.IsWholeRowOrColumn
                && string.Equals(reference.Sheet, sourceSheet, StringComparison.OrdinalIgnoreCase)
                && !(reference.Start > 0 && formula[reference.Start - 1] == ':');
        }
    }
}
