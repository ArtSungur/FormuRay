using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using FormuRay.Core;
using FormuRay.Excel.Names;

namespace FormuRay.Excel.UI
{
    public sealed class NamedRangesManagerForm : Form
    {
        private readonly NamedRangesService service;
        private readonly DataGridView grid = new DataGridView();
        private readonly ComboBox filter = new ComboBox();
        private readonly TextBox search = new TextBox();
        private readonly Label workbook = new Label();
        private List<NamedRangeInfo> snapshot = new List<NamedRangeInfo>();
        public NamedRangesManagerForm(NamedRangesService service)
        {
            this.service = service;
            Text = "FormuRay - Named Ranges Manager";
            Size = new Size(1100, 530);
            MinimumSize = new Size(760, 350);
            StartPosition = FormStartPosition.CenterParent;
            Font = SystemFonts.MessageBoxFont;
            var top = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8) };
            workbook.AutoSize = true; workbook.Margin = new Padding(3, 7, 20, 3);
            filter.DropDownStyle = ComboBoxStyle.DropDownList;
            filter.Items.AddRange(new object[] { "All", "Visible", "Hidden", "#REF!", "#VALUE!", "No matching error" });
            filter.SelectedIndex = 0;
            search.Width = 240;
            top.Controls.Add(workbook); top.Controls.Add(filter);
            top.Controls.Add(new Label { Text = "Search:", AutoSize = true, Margin = new Padding(10, 7, 0, 0) });
            top.Controls.Add(search);
            grid.Dock = DockStyle.Fill; grid.ReadOnly = true; grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false; grid.RowHeadersVisible = false; grid.AutoGenerateColumns = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = true;
            grid.AllowUserToOrderColumns = true; grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
            foreach (string column in new[] { "Name", "Value", "Refers To", "Scope", "Visibility", "Status" })
                grid.Columns.Add(new DataGridViewTextBoxColumn { Name = column, HeaderText = column,
                    FillWeight = column == "Refers To" ? 230 : column == "Status" ? 150 : 100,
                    SortMode = DataGridViewColumnSortMode.Automatic });
            grid.FontChanged += (s, e) => SizeHeaders();
            Shown += (s, e) => SizeHeaders();
            SizeHeaders();
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(8) };
            AddButton(buttons, "Clean #REF", () => Cleanup(NameCleanupKind.Reference));
            AddButton(buttons, "Clean #VALUE", () => Cleanup(NameCleanupKind.Value));
            AddButton(buttons, "Clean Hidden", () => Cleanup(NameCleanupKind.Hidden));
            AddButton(buttons, "Go to Range", Navigate);
            AddButton(buttons, "Refresh", RefreshNames);
            var close = AddButton(buttons, "Close", Close); close.DialogResult = DialogResult.Cancel; CancelButton = close;
            Controls.Add(grid); Controls.Add(top); Controls.Add(buttons);
            filter.SelectedIndexChanged += (s, e) => Render(); search.TextChanged += (s, e) => Render();
            grid.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) Safe(Navigate); };
            FormClosed += (s, e) => service.Clear();
        }
        private void SizeHeaders()
        {
            var font = grid.ColumnHeadersDefaultCellStyle.Font ?? grid.Font;
            grid.ColumnHeadersHeight = Math.Max(34, TextRenderer.MeasureText("Visibility", font).Height + 14);
            foreach (DataGridViewColumn column in grid.Columns)
            {
                int floor = column.Name == "Refers To" ? 210 : column.Name == "Status" ? 180
                    : column.Name == "Name" ? 145 : column.Name == "Value" ? 125 : 115;
                column.MinimumWidth = Math.Max(floor, TextRenderer.MeasureText(column.HeaderText, font).Width + 34);
            }
        }
        private Button AddButton(FlowLayoutPanel panel, string text, Action action)
        {
            var button = new Button { Text = text, AutoSize = true, Height = 30 };
            button.Click += (s, e) => Safe(action); panel.Controls.Add(button); return button;
        }
        private void Safe(Action action)
        {
            try { action(); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Named Ranges Manager", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
        }
        public void RefreshNames()
        {
            snapshot.Clear(); Render(); workbook.Text = "Workbook: none";
            snapshot = service.Refresh();
            workbook.Text = "Workbook: " + service.WorkbookName;
            Render();
        }
        private void Render()
        {
            grid.Rows.Clear();
            foreach (var item in snapshot)
            {
                if (filter.SelectedIndex == 1 && !item.Visible || filter.SelectedIndex == 2 && item.Visible
                    || filter.SelectedIndex == 3 && !item.HasReferenceError || filter.SelectedIndex == 4 && !item.HasValueError
                    || filter.SelectedIndex == 5 && (item.HasReferenceError || item.HasValueError)) continue;
                if ((item.Name + " " + item.Scope + " " + item.RefersTo).IndexOf(search.Text, StringComparison.OrdinalIgnoreCase) < 0) continue;
                int index = grid.Rows.Add(item.Name, item.Value, item.RefersTo, item.Scope, item.Visibility, item.Status);
                grid.Rows[index].Tag = item;
            }
        }
        private void Navigate()
        {
            if (grid.CurrentRow != null) service.Navigate(grid.CurrentRow.Tag as NamedRangeInfo);
        }
        private void Cleanup(NameCleanupKind kind)
        {
            // Like VBA, cleanup targets the current active user workbook, independently of the display filter.
            // Refresh the display first so the workbook named in the confirmation is never ambiguous.
            RefreshNames();
            var matches = snapshot.Where(n => NamedRangeRules.Matches(n, kind)).ToList();
            string criterion = kind == NameCleanupKind.Reference ? "#REF" : kind == NameCleanupKind.Value ? "#VALUE" : "hidden";
            if (matches.Count == 0) { MessageBox.Show(this, "No " + criterion + " names were found.", Text); return; }
            string warning = kind == NameCleanupKind.Hidden
                ? "\nThis includes internal names created by Excel." : "";
            string message = "Delete all " + matches.Count + " " + criterion + " names from " + service.WorkbookName
                + "?\nThis applies to the whole workbook, regardless of the filter or selection." + warning;
            if (MessageBox.Show(this, message, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            var result = service.Delete(matches, kind);
            try { RefreshNames(); }
            finally { MessageBox.Show(this, "Found: " + matches.Count + "\nDeleted: " + result.Deleted + "\nFailed: " + result.Failed, Text); }
        }
    }
}
