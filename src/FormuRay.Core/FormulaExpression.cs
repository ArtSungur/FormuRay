using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace FormuRay.Core
{
    public sealed class FormulaExpressionNode
    {
        public string Element { get; internal set; }
        public string Kind { get; internal set; }
        public string Value { get; internal set; }
        public int Start { get; internal set; }
        public int Length { get; internal set; }
        public FormulaReference Reference { get; internal set; }
        public ExplorerComponent Component { get; set; }
        public List<FormulaExpressionNode> Children { get; private set; }
        public FormulaExpressionNode() { Children = new List<FormulaExpressionNode>(); }
        public IEnumerable<FormulaExpressionNode> DescendantsAndSelf()
        {
            yield return this;
            foreach (var child in Children) foreach (var node in child.DescendantsAndSelf()) yield return node;
        }
    }

    public sealed class ParsedFormula
    {
        public List<FormulaExpressionNode> Roots { get; internal set; }
        public bool IsComplete { get; internal set; }
        public IEnumerable<FormulaExpressionNode> Nodes { get { return Roots.SelectMany(n => n.DescendantsAndSelf()); } }
        // Both component navigation and the tree use these same node instances.
        public IEnumerable<FormulaExpressionNode> Components
        {
            get { return Nodes.Where(n => n.Kind != "Function" && n.Kind != "Operator").OrderBy(n => n.Start); }
        }
    }

    // Expression layer for the existing reference parser. Never evaluates formulas.
    // Excel precedence (not VBA): reference operators, unary sign, %, ^, */,
    // +-, &, comparisons. Equal-precedence binary operators associate left-to-right.
    internal sealed class FormulaExpressionReader
    {
        private sealed class Token
        {
            internal string Text, Kind;
            internal int Start;
            internal FormulaReference Reference;
        }
        private static readonly Regex Number = new Regex(@"\G(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?", RegexOptions.CultureInvariant);
        private static readonly Regex Error = new Regex(@"\G#(?:REF!|DIV/0!|VALUE!|N/A|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private readonly List<Token> tokens = new List<Token>();
        private int index, depth;

        internal static ParsedFormula Read(string text, string workbook)
        {
            var reader = new FormulaExpressionReader();
            reader.Tokenize(text ?? "", workbook);
            try
            {
                var roots = new List<FormulaExpressionNode>();
                if (reader.tokens.Count > 0) roots.Add(reader.Expression(0, false));
                if (reader.index != reader.tokens.Count) throw new FormatException();
                return new ParsedFormula { Roots = roots, IsComplete = true };
            }
            catch (FormatException)
            {
                // Unsupported or malformed grammar must not hide recognizable components.
                return new ParsedFormula { Roots = reader.tokens.Where(t => t.Kind != "Symbol").Select(Leaf).ToList(), IsComplete = false };
            }
        }

        private void Tokenize(string text, string workbook)
        {
            var references = FormulaReferenceParser.Parse(text, workbook).GroupBy(r => r.Start).ToDictionary(g => g.Key, g => g.First());
            for (int i = text.StartsWith("=", StringComparison.Ordinal) ? 1 : 0; i < text.Length;)
            {
                if (char.IsWhiteSpace(text[i])) { i++; continue; }
                int start = i;
                if (text[i] == '"')
                {
                    i++;
                    while (i < text.Length)
                    {
                        if (text[i++] != '"') continue;
                        if (i < text.Length && text[i] == '"') { i++; continue; }
                        break;
                    }
                    Add(text.Substring(start, i - start), "Constant", start); continue;
                }
                FormulaReference reference;
                if (references.TryGetValue(i, out reference)
                    && (i + reference.Text.Length == text.Length || "![".IndexOf(text[i + reference.Text.Length]) < 0))
                {
                    tokens.Add(new Token { Text = reference.Text, Start = i, Kind = reference.IsExternal ? "External"
                        : reference.Sheet.Length > 0 ? "Cross-sheet" : reference.Address.Contains(":") ? "Range" : "Reference", Reference = reference });
                    i += reference.Text.Length; continue;
                }
                var match = Error.Match(text, i);
                if (!match.Success) match = Number.Match(text, i);
                if (match.Success) { Add(match.Value, "Constant", i); i += match.Length; continue; }
                if (text[i] == '{')
                {
                    bool quoted = false;
                    do { if (text[i] == '"') quoted = !quoted; i++; }
                    while (i < text.Length && (text[i - 1] != '}' || quoted));
                    Add(text.Substring(start, i - start), "Constant", start); continue;
                }
                if ("()+-*/^&=<>%,;:@#".IndexOf(text[i]) >= 0)
                {
                    i++;
                    if (i < text.Length && (text[start] == '<' || text[start] == '>') && (text[i] == '=' || (text[start] == '<' && text[i] == '>'))) i++;
                    Add(text.Substring(start, i - start), "Symbol", start); continue;
                }
                int brackets = 0; bool singleQuoted = false;
                do
                {
                    char ch = text[i];
                    if (ch == '\'') singleQuoted = !singleQuoted;
                    if (!singleQuoted) { if (ch == '[') brackets++; if (ch == ']') brackets--; }
                    i++;
                }
                while (i < text.Length && (singleQuoted || brackets > 0 || (!char.IsWhiteSpace(text[i]) && "()+-*/^&=<>%,;:@#".IndexOf(text[i]) < 0)));
                string atom = text.Substring(start, i - start);
                Add(atom, atom.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || atom.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ? "Constant" : "Name", start);
            }
        }
        private void Add(string text, string kind, int start) { tokens.Add(new Token { Text = text, Kind = kind, Start = start }); }
        private Token Peek { get { return index < tokens.Count ? tokens[index] : null; } }
        private bool Is(string text) { return Peek != null && Peek.Text == text; }
        private static FormulaExpressionNode Leaf(Token token)
        {
            string value = token.Kind == "Constant" ? token.Text : "";
            if (value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"') value = value.Substring(1, value.Length - 2).Replace("\"\"", "\"");
            return new FormulaExpressionNode { Element = token.Text, Kind = token.Kind, Start = token.Start, Length = token.Text.Length, Value = value, Reference = token.Reference };
        }
        private FormulaExpressionNode Expression(int minimum, bool argument)
        {
            if (++depth > 128) throw new FormatException();
            try
            {
                FormulaExpressionNode left = Primary(argument);
                while (Peek != null)
                {
                    if (Is("%") || Is("#"))
                    {
                        int priority = Is("#") ? 12 : 6;
                        if (priority < minimum) break;
                        var token = tokens[index++];
                        left = Operator(token.Text, left.Start, token.Start + token.Text.Length, left); continue;
                    }
                    bool intersection = Peek.Kind != "Symbol" && Peek.Start > left.Start + left.Length;
                    string operation = intersection ? " " : Peek.Text;
                    int precedence = Priority(operation);
                    if (precedence < minimum || (argument && (operation == "," || operation == ";"))) break;
                    if (!intersection) index++;
                    var right = Expression(precedence + 1, argument);
                    left = Operator(operation, left.Start, right.Start + right.Length, left, right);
                }
                return left;
            }
            finally { depth--; }
        }
        private FormulaExpressionNode Primary(bool argument)
        {
            if (Peek == null) throw new FormatException();
            var token = tokens[index++];
            if (token.Text == "(")
            {
                var inner = Expression(0, false);
                if (!Is(")")) throw new FormatException();
                int end = tokens[index++].Start + 1;
                // Parentheses affect parsing, without introducing an artificial tree node.
                inner.Start = token.Start; inner.Length = end - token.Start; return inner;
            }
            if (token.Text == "+" || token.Text == "-" || token.Text == "@")
            {
                var child = Expression(7, argument);
                double number;
                if ((token.Text == "-" || token.Text == "+") && child.Kind == "Constant" && child.Element.Length > 0
                    && child.Element[0] != '-' && child.Element[0] != '+'
                    && double.TryParse(child.Element, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number))
                    return new FormulaExpressionNode { Element = token.Text + child.Element, Kind = "Constant", Value = token.Text + child.Element,
                        Start = token.Start, Length = child.Start + child.Length - token.Start };
                return Operator(token.Text, token.Start, child.Start + child.Length, child);
            }
            if (token.Kind == "Symbol") throw new FormatException();
            if (Is("(") && token.Kind == "Name")
            {
                index++;
                var node = new FormulaExpressionNode { Element = token.Text, Kind = "Function", Start = token.Start };
                while (!Is(")"))
                {
                    if (Peek == null) throw new FormatException();
                    if (Is(",") || Is(";")) { index++; continue; } // Optional omitted argument.
                    node.Children.Add(Expression(0, true));
                    if (Is(",") || Is(";")) index++;
                    else if (!Is(")")) throw new FormatException();
                }
                node.Length = tokens[index++].Start + 1 - node.Start; return node;
            }
            return Leaf(token);
        }
        private static FormulaExpressionNode Operator(string text, int start, int end, params FormulaExpressionNode[] children)
        {
            var node = new FormulaExpressionNode { Element = text, Kind = "Operator", Start = start, Length = end - start };
            node.Children.AddRange(children); return node;
        }
        private static int Priority(string text)
        {
            switch (text)
            {
                case "=": case "<>": case "<": case ">": case "<=": case ">=": return 1;
                case "&": return 2;
                case "+": case "-": return 3;
                case "*": case "/": return 4;
                case "^": return 5;
                case ",": case ";": return 8;
                case " ": return 9;
                case ":": return 10;
                default: return -1;
            }
        }
    }
}
