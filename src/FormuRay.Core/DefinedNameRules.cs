using System;
using System.Collections.Generic;
using System.Linq;

namespace FormuRay.Core
{
    public sealed class NameOccurrence
    {
        public string Text { get; set; }
        public int Start { get; set; }
        public int Occurrence { get; set; }
    }

    public static class DefinedNameRules
    {
        public const int MaximumNames = 1000;
        public static string SafetyKind(string refersTo)
        {
            string text = (refersTo ?? "").Trim(), upper = text.ToUpperInvariant();
            if (text.Length == 0 || text.Length > 2048) return "unresolved";
            if (upper.Contains("#REF!")) return "invalid";
            if (text.Contains("[") || upper.Contains("HTTP:") || upper.Contains("HTTPS:")) return "external";
            if (new[] { "INDIRECT", "OFFSET", "GETPIVOTDATA", "ДВССЫЛ", "СМЕЩ", "ПОЛУЧИТЬ.ДАННЫЕ.СВОДНОЙ.ТАБЛИЦЫ" }.Any(n => upper.Contains(n + "("))) return "dynamic";
            string expression = text.StartsWith("=", StringComparison.Ordinal) ? text.Substring(1).Trim() : text;
            int bang = expression.LastIndexOf('!');
            if (bang >= 0 && FormulaReferenceParser.UnquoteSheet(expression.Substring(0, bang).Trim()).Length == 0) return "formula";
            return FormulaReferenceParser.IsSimpleAddress(expression.Substring(bang + 1).Trim()) ? "safe" : "formula";
        }

        public static List<NameOccurrence> Find(string fullName, string localName, string sourceSheet,
            bool assumeSourceScope, string formula, string formulaLocal)
        {
            var candidates = Candidates(fullName ?? "", localName ?? "", sourceSheet, assumeSourceScope);
            var matches = FindInFormula(formulaLocal ?? "", candidates);
            return matches.Count > 0 || formula == formulaLocal ? matches : FindInFormula(formula ?? "", candidates);
        }

        private static List<string> Candidates(string full, string local, string sourceSheet, bool assumeScope)
        {
            string shortName = full.Substring(full.LastIndexOf('!') + 1), shortLocal = local.Substring(local.LastIndexOf('!') + 1);
            string qualified = full.Contains("!") ? full : local;
            string scope = qualified.Contains("!") ? FormulaReferenceParser.UnquoteSheet(qualified.Substring(0, qualified.LastIndexOf('!')).Trim()) : "";
            if (scope.StartsWith("[", StringComparison.Ordinal) && scope.Contains("]")) scope = scope.Substring(scope.IndexOf(']') + 1);
            scope = FormulaReferenceParser.UnquoteSheet(scope.Trim());
            if (scope.Length == 0 && assumeScope) scope = sourceSheet;
            var result = new List<string> { full, local };
            if (scope.Length > 0)
            {
                foreach (string name in new[] { shortName, shortLocal })
                {
                    if (name.Length == 0) continue;
                    result.Add(FormulaReferenceParser.QuoteSheet(scope) + "!" + name);
                    result.Add(scope + "!" + name);
                }
            }
            if (scope.Length == 0 || string.Equals(scope, sourceSheet, StringComparison.OrdinalIgnoreCase)) result.AddRange(new[] { shortName, shortLocal });
            return result.Where(s => s.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static List<NameOccurrence> FindInFormula(string formula, IEnumerable<string> candidates)
        {
            var matches = new SortedDictionary<int, NameOccurrence>();
            string clean = FormulaReferenceParser.MaskStrings(formula);
            foreach (string candidate in candidates)
            {
                int search = 0, at;
                while (search < clean.Length && (at = clean.IndexOf(candidate, search, StringComparison.OrdinalIgnoreCase)) >= 0)
                {
                    int end = at + candidate.Length;
                    bool qualifiedPart = !candidate.Contains("!") && ((at > 0 && clean[at - 1] == '!') || (end < clean.Length && clean[end] == '!'));
                    if (!qualifiedPart && Boundary(clean, at - 1) && Boundary(clean, end) && !matches.ContainsKey(at))
                        matches.Add(at, new NameOccurrence { Text = formula.Substring(at, candidate.Length), Start = at });
                    search = end;
                }
            }
            int occurrence = 0;
            foreach (var item in matches.Values) item.Occurrence = ++occurrence;
            return matches.Values.ToList();
        }
        private static bool Boundary(string text, int at) { return at < 0 || at >= text.Length || " =+-*/^&(),;:{}<>!%\t\r\n".IndexOf(text[at]) >= 0; }
    }
}
