using System.Collections.Generic;
using FormuRay.Core;

namespace FormuRay.Excel.UI
{
    public static class FormulaExplorerPresentation
    {
        // Project AST edges onto the nearest visible function/root. Keep the original
        // node instances and source order; operators never create a row or a level.
        public static IEnumerable<FormulaExpressionNode> VisibleChildren(IEnumerable<FormulaExpressionNode> nodes)
        {
            foreach (var node in nodes)
            {
                if (node.Kind == "Operator")
                {
                    foreach (var child in VisibleChildren(node.Children)) yield return child;
                }
                else yield return node;
            }
        }
    }
}
