using System;

namespace FormuRay.Core
{
    /// <summary>The value-type whitelist from HD_IsHardcodedInputCell, without Excel COM.</summary>
    public static class HardcodedInputRules
    {
        public const int MaximumHighlightedCells = 20000;

        // Error status must come from the adapter: an Excel error can marshal as an integer.
        // In particular, do not parse strings or inspect numeric literals inside formulas.
        public static bool IsHardcodedInput(object value, bool hasFormula, bool isError)
        {
            if (hasFormula || isError || value == null)
                return false;

            return value is bool || value is byte || value is short || value is int
                || value is float || value is double || value is decimal || value is DateTime;
        }
    }
}
