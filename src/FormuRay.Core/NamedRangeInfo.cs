using System;
using System.Collections.Generic;
using System.Linq;

namespace FormuRay.Core
{
    public enum NameCleanupKind { Reference, Value, Hidden }

    // A display snapshot only: no Excel objects are exposed to the grid.
    public sealed class NamedRangeInfo
    {
        public string Identity { get; set; }
        public string Name { get; set; }
        public string Scope { get; set; }
        public string Workbook { get; set; }
        public string RefersTo { get; set; }
        public string RefersToLocal { get; set; }
        public string NameValue { get; set; }
        public string Value { get; set; }
        public bool Visible { get; set; }
        public bool HasReferenceError { get { return NamedRangeRules.Contains(this, "#REF", "#ССЫЛКА"); } }
        public bool HasValueError { get { return NamedRangeRules.Contains(this, "#VALUE", "#ЗНАЧ"); } }
        public string Visibility { get { return Visible ? "Visible" : "Hidden"; } }
        public string Status { get { return HasReferenceError && HasValueError ? "#REF! / #VALUE!" : HasReferenceError ? "#REF!" : HasValueError ? "#VALUE!" : "No matching error"; } }
    }

    public static class NamedRangeRules
    {
        // Preserve VBA substring matching, even inside a constant string. Do not Evaluate to classify.
        public static bool Contains(NamedRangeInfo name, params string[] tokens)
        {
            return new[] { name.RefersTo, name.RefersToLocal, name.NameValue }.Any(
                text => tokens.Any(token => (text ?? "").IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0));
        }
        public static bool Matches(NamedRangeInfo name, NameCleanupKind kind)
        {
            return kind == NameCleanupKind.Reference ? name.HasReferenceError
                : kind == NameCleanupKind.Value ? name.HasValueError : kind == NameCleanupKind.Hidden && !name.Visible;
        }
        public static string DisplayName(string identity)
        {
            return identity.Substring(identity.LastIndexOf('!') + 1);
        }
        public static string Scope(string identity, string hint, IEnumerable<string> sheets)
        {
            if (!string.IsNullOrEmpty(hint)) return hint;
            int bang = identity.LastIndexOf('!');
            if (bang < 0) return "Workbook";
            string scope = FormulaReferenceParser.UnquoteSheet(identity.Substring(0, bang).Trim());
            int bracket = scope.LastIndexOf(']');
            if (bracket >= 0) scope = scope.Substring(bracket + 1);
            return sheets.FirstOrDefault(s => string.Equals(s, scope, StringComparison.OrdinalIgnoreCase)) ?? "Workbook";
        }
    }
}
