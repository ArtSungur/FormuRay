using System;
using System.Collections.Generic;

namespace FormuRay.Core
{
    public sealed class FormulaSpan
    {
        public int Start { get; set; }
        public int Length { get; set; }
    }

    public static class FormulaTokenLocator
    {
        public static FormulaSpan Find(string formula, string candidate, int occurrence = 1, int preferredStart = -1, bool maskStrings = true)
        {
            if (string.IsNullOrWhiteSpace(candidate)) return null;
            string clean = maskStrings ? FormulaReferenceParser.MaskStrings(formula) : formula;
            candidate = candidate.Replace("[Audited cell]", "").Trim();
            while (candidate.StartsWith("> ", StringComparison.Ordinal) || candidate.StartsWith("- ", StringComparison.Ordinal)) candidate = candidate.Substring(2).Trim();
            var candidates = new List<string> { candidate, candidate.Replace("$", "") };
            int bang = candidate.LastIndexOf('!');
            if (bang >= 0)
            {
                string sheet = FormulaReferenceParser.UnquoteSheet(candidate.Substring(0, bang)), address = candidate.Substring(bang + 1);
                foreach (string variant in new[] { FormulaReferenceParser.QuoteSheet(sheet) + "!" + address, sheet + "!" + address, address })
                { candidates.Add(variant); candidates.Add(variant.Replace("$", "")); }
            }
            foreach (var token in candidates)
            {
                if (token.Length == 0) continue;
                if (At(clean, token, preferredStart)) return new FormulaSpan { Start = preferredStart, Length = token.Length };
                int start = 0, count = 0, at;
                while (start < clean.Length && (at = clean.IndexOf(token, start, StringComparison.OrdinalIgnoreCase)) >= 0)
                {
                    if (At(clean, token, at) && ++count >= Math.Max(1, occurrence)) return new FormulaSpan { Start = at, Length = token.Length };
                    start = at + token.Length;
                }
            }
            return null;
        }
        private static bool At(string formula, string token, int at)
        {
            if (at < 0 || at + token.Length > formula.Length) return false;
            if (string.Compare(formula, at, token, 0, token.Length, StringComparison.OrdinalIgnoreCase) != 0) return false;
            return (at == 0 || " =+-*/^&(),;:{}<>!%'\t\r\n".IndexOf(formula[at - 1]) >= 0)
                && (at + token.Length == formula.Length || " =+-*/^&(),;:{}<>%'\t\r\n".IndexOf(formula[at + token.Length]) >= 0);
        }
    }
}
