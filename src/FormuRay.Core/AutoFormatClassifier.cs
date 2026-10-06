using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace FormuRay.Core
{
    public static class AutoFormatClassifier
    {
        public static AutoFormatCellType? Classify(bool hasFormula, object value, string numberFormat, bool isError = false)
        {
            if (isError || value == null || value is ErrorWrapper || !IsNumeric(value)) return null;
            FormatKind format = ClassifyFormat(numberFormat);
            if (hasFormula)
                return format == FormatKind.Percent ? AutoFormatCellType.FormulaPercent
                    : format == FormatKind.Date ? AutoFormatCellType.FormulaDate : AutoFormatCellType.FormulaValue;
            return format == FormatKind.Percent ? AutoFormatCellType.InputPercent
                : format == FormatKind.Date ? AutoFormatCellType.InputDate : AutoFormatCellType.InputValue;
        }

        private enum FormatKind { Value, Percent, Date }
        private static bool IsNumeric(object value)
        {
            return value is double || value is float || value is decimal || value is int || value is long
                || value is short || value is byte;
        }
        private static FormatKind ClassifyFormat(string format)
        {
            if (string.IsNullOrWhiteSpace(format)) return FormatKind.Value;
            var sections = VisibleSections(format);
            FormatKind? found = null;
            for (int i = 0; i < sections.Count && i < 3; i++)
            {
                string section = sections[i];
                if (section.Length == 0) continue;
                bool percent = section.IndexOf('%') >= 0;
                bool date = section.IndexOfAny(new[] { 'y', 'Y', 'd', 'D', 'h', 'H', 's', 'S', 'm', 'M' }) >= 0;
                FormatKind current = percent && date ? FormatKind.Value
                    : percent ? FormatKind.Percent : date ? FormatKind.Date : FormatKind.Value;
                if (found.HasValue && found.Value != current) return FormatKind.Value;
                found = current;
            }
            return found ?? FormatKind.Value;
        }

        private static List<string> VisibleSections(string format)
        {
            var sections = new List<string>();
            var visible = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < format.Length; i++)
            {
                char ch = format[i];
                if (quoted)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < format.Length && format[i + 1] == '"') i++;
                        else quoted = false;
                    }
                    continue;
                }
                if (ch == '"') { quoted = true; continue; }
                if (ch == '\\' || ch == '_' || ch == '*') { i++; continue; }
                if (ch == '[')
                {
                    int end = format.IndexOf(']', i + 1);
                    if (end >= 0)
                    {
                        string directive = format.Substring(i + 1, end - i - 1);
                        if (directive == "h" || directive == "hh" || directive == "m" || directive == "mm"
                            || directive == "s" || directive == "ss") visible.Append(directive);
                        i = end; continue;
                    }
                }
                if (ch == ';') { sections.Add(visible.ToString()); visible.Clear(); continue; }
                visible.Append(ch);
            }
            sections.Add(visible.ToString());
            return sections;
        }
    }
}
