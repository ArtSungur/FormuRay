using System;
using System.Collections.Generic;
using System.Linq;
using FormuRay.Core;

namespace FormuRay.Excel.UI
{
    // Display-only alignment of the two strings returned by Excel. No translation
    // dictionary and no second semantic parser: invariant nodes remain authoritative.
    public sealed class FormulaLocalDisplay
    {
        private sealed class Token { internal int Start, End; internal string Kind, Text; }
        private readonly List<Token> invariant, localized;
        private readonly string source;
        private readonly bool aligned;
        private readonly List<FormulaFunction> invariantFunctions, localFunctions;
        public FormulaLocalDisplay(string formula, string formulaLocal, string decimalSeparator)
        {
            source = string.IsNullOrEmpty(formulaLocal) ? formula ?? "" : formulaLocal;
            invariant = Tokens(formula ?? "", '.');
            localized = Tokens(source, string.IsNullOrEmpty(decimalSeparator) ? '.' : decimalSeparator[0]);
            invariantFunctions = FormulaFunctions.Parse(formula);
            localFunctions = FormulaFunctions.Parse(source);
            aligned = invariant.Count == localized.Count && invariant.Zip(localized, (a, b) =>
                a.Kind == b.Kind && (a.Kind != "Symbol" || a.Text == b.Text || (",;\\|".Contains(a.Text) && ",;\\|".Contains(b.Text)))).All(v => v);
        }
        public FormulaSpan Span(FormulaExpressionNode node)
        {
            if (!aligned)
            {
                var function = LocalFunction(node);
                return function == null ? null : new FormulaSpan { Start = function.Start, Length = function.Length };
            }
            int first = invariant.FindIndex(t => t.Start >= node.Start && t.Start < node.Start + node.Length);
            int last = invariant.FindLastIndex(t => t.End > node.Start && t.End <= node.Start + node.Length);
            return first < 0 || last < first ? null : new FormulaSpan { Start = localized[first].Start, Length = localized[last].End - localized[first].Start };
        }
        public string Element(FormulaExpressionNode node)
        {
            if (node.Kind == "Function")
            {
                var function = LocalFunction(node);
                if (function != null) return function.Name;
            }
            var span = Span(node);
            if (span == null) return node.Element;
            if (node.Kind == "Function")
            {
                int index = invariant.FindIndex(t => t.Start >= node.Start && t.End <= node.Start + node.Length && t.Text == node.Element);
                return index < 0 ? node.Element : localized[index].Text;
            }
            string text = source.Substring(span.Start, span.Length);
            // Grouping parentheses belong to the selection span, not the leaf label.
            if (node.Kind != "Operator")
                while (text.StartsWith("(", StringComparison.Ordinal) && text.EndsWith(")", StringComparison.Ordinal)) text = text.Substring(1, text.Length - 2).Trim();
            return text;
        }
        private FormulaFunction LocalFunction(FormulaExpressionNode node)
        {
            // Function ordinal/nesting is unchanged by localization. This fallback
            // also handles syntax the lightweight token alignment cannot recognize.
            if (node.Kind != "Function" || invariantFunctions.Count != localFunctions.Count) return null;
            int index = invariantFunctions.FindIndex(f => f.Start >= node.Start && f.Start + f.Length <= node.Start + node.Length && f.Name == node.Element);
            if (index < 0 || invariantFunctions[index].ParentOrdinal != localFunctions[index].ParentOrdinal) return null;
            return localFunctions[index];
        }
        private static List<Token> Tokens(string text, char decimalSeparator)
        {
            var result = new List<Token>();
            int arrayDepth = 0;
            for (int i = 0; i < text.Length;)
            {
                if (char.IsWhiteSpace(text[i])) { i++; continue; }
                int start = i; char ch = text[i++]; string kind;
                if (ch == '{') arrayDepth++;
                if (ch == '}') arrayDepth--;
                if (ch == '\'' || ch == '"')
                {
                    kind = ch == '"' ? "String" : "Atom";
                    while (i < text.Length)
                    {
                        if (text[i++] != ch) continue;
                        if (i < text.Length && text[i] == ch) { i++; continue; }
                        break;
                    }
                }
                else if (ch == '[')
                {
                    kind = "Atom"; int depth = 1;
                    while (i < text.Length && depth > 0) { if (text[i] == '[') depth++; if (text[i] == ']') depth--; i++; }
                }
                else if (char.IsDigit(ch) || (ch == decimalSeparator && i < text.Length && char.IsDigit(text[i])))
                {
                    kind = "Number";
                    while (i < text.Length && (char.IsDigit(text[i]) || text[i] == decimalSeparator)) i++;
                    if (i < text.Length && (text[i] == 'E' || text[i] == 'e'))
                    {
                        i++; if (i < text.Length && (text[i] == '+' || text[i] == '-')) i++;
                        while (i < text.Length && char.IsDigit(text[i])) i++;
                    }
                }
                else if (ch == '#')
                {
                    kind = "Error";
                    while (i < text.Length && !char.IsWhiteSpace(text[i]) && "+-*^&=<>(),;:{}".IndexOf(text[i]) < 0) i++;
                }
                else if ("()+-*/^&=<>%,;:{}!@".IndexOf(ch) >= 0 || (arrayDepth > 0 && (ch == '\\' || ch == '|'))) kind = "Symbol";
                else
                {
                    kind = "Atom";
                    while (i < text.Length && !char.IsWhiteSpace(text[i]) && "()+-*/^&=<>%,;:{}!@[]'\"".IndexOf(text[i]) < 0) i++;
                }
                result.Add(new Token { Start = start, End = i, Kind = kind, Text = text.Substring(start, i - start) });
            }
            return result;
        }
    }
}
