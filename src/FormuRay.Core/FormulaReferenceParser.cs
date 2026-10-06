using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace FormuRay.Core
{
    public static class FormulaReferenceParser
    {
        public static ParsedFormula ParseExpression(string formula, string contextWorkbook = "")
        {
            return FormulaExpressionReader.Read(formula, contextWorkbook);
        }
        private const string Cell = @"\$?[A-Za-z]{1,3}\$?\d{1,7}";
        private const string Address = "(" + Cell + "(:" + Cell + @")?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})";
        private const string Sheet = @"('([^']|'')+'|(\[[^\]]+\])?[^'!\[\]\s\+\-\*/\^&=<>(),;:{}]+)!";
        private static readonly Regex References = new Regex(@"(^|[^A-Za-z0-9_.])(?<ref>(" + Sheet + ")?" + Address + @")(?=$|[^A-Za-z0-9_(])", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public static List<FormulaReference> Parse(string formula, string contextWorkbook = "", bool keepDuplicates = true)
        {
            var result = new List<FormulaReference>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var occurrences = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in References.Matches(MaskStrings(formula ?? "")))
            {
                Group token = match.Groups["ref"];
                string raw = token.Value;
                int bang = raw.LastIndexOf('!');
                string qualifier = bang < 0 ? "" : UnquoteSheet(raw.Substring(0, bang));
                string workbook = "", sheet = qualifier;
                int open = qualifier.LastIndexOf('['), close = qualifier.LastIndexOf(']');
                if (open >= 0 && close > open)
                {
                    workbook = qualifier.Substring(open + 1, close - open - 1);
                    sheet = UnquoteSheet(qualifier.Substring(close + 1).Trim());
                }
                int occurrence;
                occurrences.TryGetValue(raw, out occurrence);
                occurrences[raw] = ++occurrence;
                var reference = new FormulaReference { Text = raw, Sheet = sheet, Workbook = workbook.Length == 0 ? contextWorkbook : workbook,
                    Address = bang < 0 ? raw : raw.Substring(bang + 1), Start = token.Index, Occurrence = occurrence };
                if (keepDuplicates || seen.Add(reference.Key)) result.Add(reference);
            }
            return result;
        }

        public static string MaskStrings(string formula)
        {
            var chars = (formula ?? "").ToCharArray();
            bool inside = false;
            for (int i = 0; i < chars.Length; i++)
            {
                char ch = chars[i];
                if (inside)
                {
                    chars[i] = ' ';
                    if (ch == '"')
                    {
                        if (i + 1 < chars.Length && chars[i + 1] == '"') chars[++i] = ' ';
                        else inside = false;
                    }
                }
                else if (ch == '"') { inside = true; chars[i] = ' '; }
            }
            return new string(chars);
        }

        public static string QuoteSheet(string sheet) { return "'" + sheet.Replace("'", "''") + "'"; }
        public static string UnquoteSheet(string sheet)
        {
            return sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\''
                ? sheet.Substring(1, sheet.Length - 2).Replace("''", "'") : sheet;
        }
        public static bool IsWholeRowOrColumn(string address)
        {
            return Regex.IsMatch(address ?? "", @"^(\$?[A-Za-z]+:\$?[A-Za-z]+|\$?\d+:\$?\d+)$");
        }
        public static bool IsSimpleAddress(string address)
        {
            return Regex.IsMatch(address ?? "", "^" + Address + "$", RegexOptions.IgnoreCase);
        }
    }
}
