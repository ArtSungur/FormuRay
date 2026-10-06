using System;
using FormuRay.Core;
using FormuRay.Excel.Audit;
using Interop = Microsoft.Office.Interop.Excel;
namespace FormuRay.Excel
{
    internal sealed class CellFillState : IDisposable
    {
        private Interop.Range cell;
        private readonly AuditFillSession fill = new AuditFillSession(true);
        internal CellFillState(Interop.Range cell) { this.cell = cell; }
        internal void Highlight(HighlightStyle style) { fill.Apply(cell, style, 1, true); }
        internal void Restore() { fill.Clear(); }
        public void Dispose() { fill.Dispose(); ComReference.Release(cell); cell = null; }
    }
}
