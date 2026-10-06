using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using FormuRay.Core;
using FormuRay.Excel.Audit;

namespace FormuRay.Excel.UI
{
    internal sealed class FormulaExplorerForm : Form
    {
        private sealed class Row
        {
            internal ExplorerComponent Component;
            internal FormulaExpressionNode Node;
            internal Row Parent;
            internal int Level;
            internal bool LastChild;
            internal FormulaLocalDisplay Local;
            internal string Element { get { return Node == null ? Component.Display : Local.Element(Node); } }
            internal bool HasChildren { get { return Node != null && Node.Kind == "Function" && FormulaExplorerPresentation.VisibleChildren(Node.Children).Any(); } }
        }
        private readonly FormulaExplorerService service;
        private readonly Action<ExplorerColumnWidths> saveWidths;
        private readonly ExplorerColumnWidths initialWidths;
        private readonly DataGridView elements = new DataGridView();
        private readonly TextBox formula = new TextBox();
        private readonly SplitContainer split = new SplitContainer();
        private double splitterRatio;
        private bool fittingSplitter, splitterDragging, splitterChanged;
        private readonly HashSet<FormulaExpressionNode> expanded = new HashSet<FormulaExpressionNode>();
        private readonly List<Row> rows = new List<Row>();
        private ExplorerDocument document;
        private FormulaLocalDisplay localDisplay;
        private ExplorerMode mode;
        private bool binding, layoutReady, widthsChanged;
        private bool fittingColumns;
        private int availableWidth;
        private ExplorerColumnWidths userWidths;
        private float dpiScale = 1;
        internal bool IsNavigating { get { return service.IsNavigating; } }

        internal FormulaExplorerForm(FormulaExplorerService service, ExplorerColumnWidths widths, Action<ExplorerColumnWidths> saveWidths)
        {
            this.service = service;
            this.saveWidths = saveWidths;
            initialWidths = widths ?? ExplorerColumnWidths.Defaults;
            splitterRatio = initialWidths.SplitterRatio;
            AutoScaleDimensions = new SizeF(96, 96);
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Segoe UI", 9);
            ClientSize = new Size(720, 430);
            MinimumSize = new Size(480, 320);
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            KeyPreview = true;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(10) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 39));
            elements.Dock = DockStyle.Fill;
            elements.ReadOnly = true;
            elements.AllowUserToAddRows = elements.AllowUserToDeleteRows = elements.AllowUserToResizeRows = false;
            elements.AllowUserToOrderColumns = false;
            elements.RowHeadersVisible = false;
            elements.MultiSelect = false;
            elements.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            elements.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            elements.BackgroundColor = SystemColors.Window;
            elements.BorderStyle = BorderStyle.FixedSingle;
            elements.CellBorderStyle = DataGridViewCellBorderStyle.None;
            elements.DefaultCellStyle.Font = new Font("Consolas", 10);
            elements.RowTemplate.Height = 20;
            elements.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            elements.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
            elements.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            foreach (string name in new[] { "Element", "Value", "Type" })
                elements.Columns.Add(new DataGridViewTextBoxColumn { Name = name, HeaderText = name, MinimumWidth = 60, SortMode = DataGridViewColumnSortMode.NotSortable });
            formula.Dock = DockStyle.Fill;
            formula.ReadOnly = true;
            formula.Multiline = true;
            formula.ScrollBars = ScrollBars.Vertical;
            formula.HideSelection = false;
            formula.Font = new Font("Consolas", 10);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
            var close = new Button { Text = "Close", AutoSize = true, Height = 29 };
            var refresh = new Button { Text = "Refresh", AutoSize = true, Height = 29 };
            buttons.Controls.AddRange(new Control[] { close, refresh });
            split.Orientation = Orientation.Horizontal;
            split.Dock = DockStyle.Fill;
            split.Size = new Size(700, 360);
            split.Panel1MinSize = 60;
            split.Panel2MinSize = 40;
            split.SplitterWidth = 5;
            split.Panel1.Controls.Add(elements);
            split.Panel2.Controls.Add(formula);
            split.SplitterMoving += (s, e) => { if (layoutReady && !fittingSplitter) splitterDragging = true; };
            split.SplitterMoved += (s, e) =>
            {
                if (!layoutReady || fittingSplitter || !splitterDragging) return;
                splitterDragging = false;
                splitterRatio = (double)split.SplitterDistance / Math.Max(1, split.ClientSize.Height - split.SplitterWidth);
                splitterRatio = Math.Max(.15, Math.Min(.9, splitterRatio));
                splitterChanged = true;
            };
            split.SizeChanged += (s, e) => { if (layoutReady) FitSplitter(); };
            layout.Controls.Add(split, 0, 0); layout.Controls.Add(buttons, 0, 1);
            Controls.Add(layout);
            close.Click += (s, e) => Close();
            refresh.Click += (s, e) => Guard(RefreshActiveCell);
            elements.CurrentCellChanged += (s, e) => { if (!binding) Guard(() => Display(true)); };
            elements.CellPainting += PaintElement;
            elements.CellMouseDown += (s, e) =>
            {
                if (e.ColumnIndex == 0 && e.RowIndex >= 0 && e.Button == MouseButtons.Left && rows[e.RowIndex].HasChildren
                    && e.X >= Indent(rows[e.RowIndex]) && e.X < Indent(rows[e.RowIndex]) + Scale(16))
                    Guard(() => Toggle(rows[e.RowIndex]));
            };
            elements.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) Guard(() => Toggle(rows[e.RowIndex])); };
            elements.ColumnWidthChanged += (s, e) =>
            {
                if (!layoutReady || fittingColumns) return;
                FitColumns(false);
                widthsChanged = true;
                userWidths = new ExplorerColumnWidths(Unscale(elements.Columns[0].Width), Unscale(elements.Columns[1].Width), Unscale(elements.Columns[2].Width));
            };
            elements.ClientSizeChanged += (s, e) => { if (layoutReady) FitColumns(false); };
            FormClosed += (s, e) =>
            {
                if ((widthsChanged || splitterChanged) && saveWidths != null)
                {
                    var retained = widthsChanged ? userWidths : initialWidths;
                    try { saveWidths(new ExplorerColumnWidths(retained.Element, retained.Value, retained.Type, splitterRatio)); }
                    catch (Exception ex) { Debug.WriteLine(ex); MessageBox.Show("Explorer layout could not be saved.", "FormuRay", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                }
                service.Dispose();
            };
        }
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            using (var graphics = CreateGraphics()) dpiScale = graphics.DpiX / 96f;
            int[] widths = { initialWidths.Element, initialWidths.Value, initialWidths.Type };
            for (int i = 0; i < widths.Length; i++)
            {
                elements.Columns[i].MinimumWidth = Scale(60);
                elements.Columns[i].Width = Scale(widths[i]);
            }
            ApplyGridMetrics();
            PositionPanel();
            split.Panel1MinSize = Scale(60);
            split.Panel2MinSize = Scale(40);
            split.SplitterWidth = Scale(5);
            FitSplitter();
            FitColumns(true);
        }
        private void ApplyGridMetrics()
        {
            elements.RowTemplate.Height = Scale(20);
            foreach (DataGridViewRow row in elements.Rows) row.Height = Scale(20);
            // Measure the effective header font, including Office/Windows DPI scaling.
            // Fixed readable height prevents the grid's initial layout shrinking it.
            var headerFont = elements.ColumnHeadersDefaultCellStyle.Font ?? elements.Font;
            elements.ColumnHeadersHeight = Math.Max(Scale(24), TextRenderer.MeasureText("Element Value Type", headerFont).Height + Scale(6));
        }
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            // SplitContainer finishes docking after OnLoad; use its final viewport.
            FitSplitter();
            FitColumns(true);
            layoutReady = true;
        }
        private int Scale(int value) { return (int)Math.Round(value * dpiScale); }
        private int Unscale(int value) { return (int)Math.Round(value / dpiScale); }
        private bool IsContext(Row row) { return row.Component != null && row.Component.Kind == "Audited"; }
        private int Indent(Row row) { return Scale(3 + (row.Level + (IsContext(row) ? 0 : 1)) * 16); }
        private void FitSplitter()
        {
            if (fittingSplitter) return;
            fittingSplitter = true;
            try
            {
                int available = split.ClientSize.Height - split.SplitterWidth;
                if (available >= split.Panel1MinSize + split.Panel2MinSize)
                    split.SplitterDistance = Math.Max(split.Panel1MinSize, Math.Min(available - split.Panel2MinSize, (int)Math.Round(available * splitterRatio)));
            }
            finally { fittingSplitter = false; }
        }

        private void FitColumns(bool initial)
        {
            if (fittingColumns) return;
            fittingColumns = true;
            try
            {
                // Reserve the vertical scrollbar even for short trees, so expanding
                // a branch never pushes Type beyond the viewport.
                int available = Math.Max(Scale(240), elements.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 2);
                int[] widths = elements.Columns.Cast<DataGridViewColumn>().Select(c => c.Width).ToArray();
                var defaults = ExplorerColumnWidths.Defaults;
                bool defaultLayout = initial && initialWidths.Element == defaults.Element && initialWidths.Value == defaults.Value && initialWidths.Type == defaults.Type;
                if (defaultLayout)
                {
                    widths[1] = widths[2] = (int)Math.Round(available * .19);
                    widths[0] = available - widths[1] - widths[2];
                }
                else if (!initial && availableWidth > 0)
                    widths[0] = Math.Max(Scale(80), widths[0] + available - availableWidth);
                int[] minimum = { Scale(80), Scale(70), Scale(100) };
                for (int i = 0; i < 3; i++) widths[i] = Math.Max(minimum[i], widths[i]);
                int total = widths.Sum();
                if (total > available)
                {
                    int extra = Math.Max(0, available - minimum.Sum());
                    int requested = total - minimum.Sum();
                    for (int i = 0; i < 3; i++) widths[i] = minimum[i] + (requested == 0 ? 0 : (int)((long)(widths[i] - minimum[i]) * extra / requested));
                }
                for (int i = 0; i < 3; i++) elements.Columns[i].Width = widths[i];
                availableWidth = available;
                elements.HorizontalScrollingOffset = 0;
            }
            finally { fittingColumns = false; }
        }

        internal void LoadMode(ExplorerMode nextMode)
        {
            mode = nextMode;
            document = service.Refresh(mode);
            localDisplay = new FormulaLocalDisplay(document.Formula, document.DisplayFormula, document.LocalDecimalSeparator);
            Text = "FormuRay - " + mode;
            expanded.Clear();
            Rebuild(null);
            Debug.WriteLine("FormuRay Explorer: " + document.Status);
        }
        internal void RefreshActiveCell() { LoadMode(mode); }
        internal void PositionPanel()
        {
            Rectangle cell = service.Context.ActiveCellScreenBounds;
            Rectangle working = Screen.FromRectangle(cell).WorkingArea;
            if (MinimumSize.Width > working.Width || MinimumSize.Height > working.Height)
                MinimumSize = new Size(Math.Min(MinimumSize.Width, working.Width), Math.Min(MinimumSize.Height, working.Height));
            Bounds = ExplorerWindowPlacement.BesideCell(cell, Size, working, Scale(10));
        }
        private void Rebuild(FormulaExpressionNode keep)
        {
            rows.Clear();
            if (document == null) return;
            var audited = document.Components.FirstOrDefault(c => c.Kind == "Audited");
            if (audited != null) rows.Add(new Row { Component = audited });
            if (document.Expression != null)
                foreach (var root in FormulaExplorerPresentation.VisibleChildren(document.Expression.Roots)) Append(root, null, 0);
            else
                foreach (var component in document.Components.Where(c => c.Kind != "Audited")) rows.Add(new Row { Component = component });
            var lastByParent = new Dictionary<Row, Row>();
            Row lastRoot = null;
            foreach (var row in rows)
            {
                row.LastChild = true;
                Row previous;
                if (row.Parent == null) { if (lastRoot != null) lastRoot.LastChild = false; lastRoot = row; }
                else { if (lastByParent.TryGetValue(row.Parent, out previous)) previous.LastChild = false; lastByParent[row.Parent] = row; }
            }
            binding = true;
            elements.SuspendLayout();
            try
            {
                elements.Rows.Clear();
                foreach (var row in rows)
                {
                    string kind = row.Node == null ? (row.Component.Target != null && document.Source != null
                        && row.Component.Target.Sheet != document.Source.Sheet ? "Cross-sheet" : "Reference") : row.Node.Kind;
                    string value = row.Component == null ? row.Node.Value : row.Component.Value;
                    if (row.Component != null && row.Component.Kind == "Audited") kind = "";
                    if (row.Node != null && row.Node.Kind == "Constant" && !row.Node.Element.StartsWith("\"", StringComparison.Ordinal)) value = row.Element;
                    int index = elements.Rows.Add(row.Element, value ?? "", kind);
                    elements.Rows[index].Height = Scale(20);
                    elements.Rows[index].Cells[0].ToolTipText = row.Element;
                }
                int selected = keep == null ? 0 : rows.FindIndex(r => r.Node == keep);
                if (rows.Count > 0) elements.CurrentCell = elements.Rows[Math.Max(0, selected)].Cells[0];
            }
            finally { elements.ResumeLayout(); binding = false; }
            Display(false);
        }
        private void Append(FormulaExpressionNode node, Row parent, int level)
        {
            var row = new Row { Node = node, Component = node.Component, Parent = parent, Level = level, Local = localDisplay };
            rows.Add(row);
            if (expanded.Contains(node)) foreach (var child in FormulaExplorerPresentation.VisibleChildren(node.Children)) Append(child, row, level + 1);
        }
        private Row Selected { get { return elements.CurrentCell == null || elements.CurrentCell.RowIndex >= rows.Count ? null : rows[elements.CurrentCell.RowIndex]; } }
        private void PaintElement(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex != 0 || e.RowIndex < 0 || e.RowIndex >= rows.Count) return;
            Row row = rows[e.RowIndex];
            e.Paint(e.ClipBounds, e.PaintParts & ~DataGridViewPaintParts.ContentForeground);
            var saved = e.Graphics.Save();
            e.Graphics.SetClip(Rectangle.Intersect(e.CellBounds, e.ClipBounds));
            int x = e.CellBounds.Left + Indent(row), middle = e.CellBounds.Top + e.CellBounds.Height / 2;
            Color color = (e.State & DataGridViewElementStates.Selected) != 0 ? e.CellStyle.SelectionForeColor : e.CellStyle.ForeColor;
            using (var connector = new Pen(Color.FromArgb(155, 155, 155)))
            {
                if (!IsContext(row))
                {
                    int trunk = x - Scale(11);
                    e.Graphics.DrawLine(connector, trunk, e.CellBounds.Top, trunk, row.LastChild ? middle : e.CellBounds.Bottom);
                    e.Graphics.DrawLine(connector, trunk, middle, x + Scale(5), middle);
                    for (Row ancestor = row.Parent; ancestor != null; ancestor = ancestor.Parent)
                    {
                        if (ancestor.LastChild) continue;
                        int ancestorTrunk = e.CellBounds.Left + Indent(ancestor) - Scale(11);
                        e.Graphics.DrawLine(connector, ancestorTrunk, e.CellBounds.Top, ancestorTrunk, e.CellBounds.Bottom);
                    }
                }
                if ((row.HasChildren && expanded.Contains(row.Node)) || (IsContext(row) && rows.Count > 1))
                    e.Graphics.DrawLine(connector, x + Scale(5), IsContext(row) ? middle + Scale(7) : middle, x + Scale(5), e.CellBounds.Bottom);
            }
            if (row.HasChildren)
            {
                int size = Scale(10), inset = Scale(2);
                using (var background = new SolidBrush((e.State & DataGridViewElementStates.Selected) != 0 ? e.CellStyle.SelectionBackColor : e.CellStyle.BackColor))
                    e.Graphics.FillRectangle(background, x, middle - size / 2, size, size);
                using (var pen = new Pen(color))
                {
                    e.Graphics.DrawRectangle(pen, x, middle - size / 2, size, size);
                    e.Graphics.DrawLine(pen, x + inset, middle, x + size - inset, middle);
                    if (!expanded.Contains(row.Node)) e.Graphics.DrawLine(pen, x + size / 2, middle - size / 2 + inset, x + size / 2, middle + size / 2 - inset);
                }
            }
            int textLeft = x + (IsContext(row) ? 0 : Scale(16));
            var textBounds = new Rectangle(textLeft, e.CellBounds.Top, Math.Max(0, e.CellBounds.Right - textLeft - Scale(3)), e.CellBounds.Height);
            TextRenderer.DrawText(e.Graphics, row.Element, e.CellStyle.Font, textBounds, color,
                TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.PreserveGraphicsClipping);
            e.Graphics.Restore(saved);
            e.Handled = true;
        }
        private void Display(bool navigate)
        {
            Row row = Selected;
            formula.Text = document == null ? "" : document.DisplayFormula;
            if (row == null) return;
            FormulaSpan span = null;
            if (row.Node != null)
            {
                span = localDisplay.Span(row.Node);
                if (navigate) service.Navigate(row.Component); // Constants/operators never select an Excel cell.
            }
            else if (row.Component != null)
            {
                var component = row.Component;
                if (!string.IsNullOrEmpty(component.FormulaLocal)) formula.Text = component.FormulaLocal;
                else if (!string.IsNullOrEmpty(component.Formula)) formula.Text = component.Formula;
                if (navigate) service.Navigate(component);
                if (component.Kind != "Audited")
                    span = FormulaTokenLocator.Find(formula.Text, component.MatchedReference, component.Occurrence, -1, false);
            }
            formula.Select(span == null ? 0 : span.Start, span == null ? 0 : span.Length);
            if (navigate) elements.Focus(); // Keep arrow-key browsing after Excel selects a reference.
        }
        private void Toggle(Row row)
        {
            if (!row.HasChildren) { Display(true); return; }
            if (!expanded.Add(row.Node)) expanded.Remove(row.Node);
            Rebuild(row.Node);
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape) { Close(); return true; }
            if (elements.ContainsFocus && (keyData == Keys.Enter || keyData == Keys.Left || keyData == Keys.Right))
            {
                Guard(() =>
                {
                    Row row = Selected;
                    if (row == null) return;
                    if (keyData == Keys.Enter) Toggle(row);
                    else if (keyData == Keys.Right && row.HasChildren)
                    {
                        if (expanded.Add(row.Node)) Rebuild(row.Node);
                        else if (elements.CurrentCell.RowIndex + 1 < rows.Count) elements.CurrentCell = elements.Rows[elements.CurrentCell.RowIndex + 1].Cells[0];
                    }
                    else if (keyData == Keys.Left)
                    {
                        if (row.Node != null && expanded.Remove(row.Node)) Rebuild(row.Node);
                        else if (row.Parent != null) elements.CurrentCell = elements.Rows[rows.IndexOf(row.Parent)].Cells[0];
                    }
                });
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        private void Guard(Action action)
        {
            try { action(); }
            catch (Exception ex)
            {
                binding = false;
                Debug.WriteLine(ex);
                MessageBox.Show(this, "Formula Audit Explorer could not complete. Check that the workbook is still open.", "FormuRay", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}
