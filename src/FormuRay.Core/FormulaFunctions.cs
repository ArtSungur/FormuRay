using System;
using System.Collections.Generic;
using System.Linq;

namespace FormuRay.Core
{
    public sealed class FormulaFunction
    {
        public string Name { get; set; }
        public string Fragment { get; set; }
        public int Start { get; set; }
        public int Length { get; set; }
        public int Ordinal { get; set; }
        public int ParentOrdinal { get; set; }
        public int Level { get; set; }
    }

    public static class FormulaFunctions
    {
        public static List<FormulaFunction> Parse(string formula)
        {
            formula = formula ?? "";
            var nodes = new List<FormulaFunction>();
            var stack = new Stack<int>();
            char quote = '\0';
            for (int i = 0; i < formula.Length; i++)
            {
                char ch = formula[i];
                if (quote != '\0')
                {
                    if (ch == quote)
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == quote) i++;
                        else quote = '\0';
                    }
                    continue;
                }
                if (ch == '"' || ch == '\'') { quote = ch; continue; }
                if (ch == '(')
                {
                    int start = i;
                    while (start > 0 && IsNameCharacter(formula[start - 1])) start--;
                    if (start < i && !char.IsDigit(formula[start]) && (start == 0 || formula[start - 1] != '['))
                    {
                        nodes.Add(new FormulaFunction { Start = start, Name = formula.Substring(start, i - start), Ordinal = nodes.Count + 1 });
                        stack.Push(nodes.Count - 1);
                    }
                    else stack.Push(-1);
                }
                else if (ch == ')' && stack.Count > 0)
                {
                    int index = stack.Pop();
                    if (index >= 0)
                    {
                        nodes[index].Length = i - nodes[index].Start + 1;
                        nodes[index].Fragment = formula.Substring(nodes[index].Start, nodes[index].Length);
                    }
                }
            }
            // VBA falls back to a flat list if a detected function has no matching close parenthesis.
            if (nodes.Any(n => n.Length == 0)) return new List<FormulaFunction>();
            foreach (var node in nodes)
            {
                var parents = nodes.Where(p => p.Start < node.Start && p.Start + p.Length > node.Start + node.Length).ToList();
                node.Level = parents.Count + 1;
                node.ParentOrdinal = parents.OrderBy(p => p.Length).Select(p => p.Ordinal).FirstOrDefault();
            }
            return nodes;
        }

        private static bool IsNameCharacter(char ch)
        {
            return " =+-*/^&(),;:{}<>!%[]'\"\t\r\n".IndexOf(ch) < 0;
        }
    }
}
