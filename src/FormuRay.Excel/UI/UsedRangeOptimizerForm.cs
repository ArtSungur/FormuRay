using System;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FormuRay.Core;

namespace FormuRay.Excel.UI
{
    public sealed class UsedRangeOptimizerForm : Form
    {
        private readonly UsedRangeOptimizerService service;
        private readonly UsedRangeInspection detected;
        private readonly NumericUpDown keepRow = new NumericUpDown();
        private readonly NumericUpDown keepColumn = new NumericUpDown();
        private readonly Label proposal = new Label();
        private readonly Label excessRows = new Label();
        private readonly Label excessColumns = new Label();
        private readonly Label keepColumnLetter = new Label();
        private readonly Button optimize = new Button();

        public UsedRangeOptimizerForm(UsedRangeOptimizerService service)
        {
            this.service = service;
            detected = service.Inspect();
            var plan = detected.Plan;
            Text = "FormuRay - Used Range Optimizer";
            Font = SystemFonts.MessageBoxFont;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            ClientSize = new Size(780, 650);
            MinimumSize = new Size(700, 610);
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(14), ColumnCount = 2, RowCount = 13 };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int row = 0; row <= 10; row++) panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            AddLine(panel, 0, "Sheet", service.SheetName);
            AddLine(panel, 1, "Current Used Range", Address(plan.Used));
            AddLine(panel, 2, "Last content row", plan.HasContent ? plan.ContentRow.ToString() : "No cell content detected");
            AddLine(panel, 3, "Last content column", plan.HasContent ? ColumnLetters(plan.ContentColumn) : "No cell content detected");
            AddLine(panel, 4, "Main Content Zone", detected.Regions.Main == null
                ? "No cell content detected" : Address(detected.Regions.Main.Bounds));
            panel.Controls.Add(new Label { Text = "Excess rows", AutoSize = true }, 0, 5);
            excessRows.AutoSize = true; panel.Controls.Add(excessRows, 1, 5);
            panel.Controls.Add(new Label { Text = "Excess columns", AutoSize = true }, 0, 6);
            excessColumns.AutoSize = true; panel.Controls.Add(excessColumns, 1, 6);
            keepRow.Minimum = 1; keepRow.Maximum = UsedRangeCleanupPlan.MaxRow;
            keepRow.Value = plan.RetainedRow; keepRow.Width = 125;
            keepColumn.Minimum = 1; keepColumn.Maximum = UsedRangeCleanupPlan.MaxColumn;
            keepColumn.Value = plan.RetainedColumn; keepColumn.Width = 125;
            panel.Controls.Add(new Label { Text = "Keep through row", AutoSize = true }, 0, 7);
            panel.Controls.Add(keepRow, 1, 7);
            panel.Controls.Add(new Label { Text = "Keep through column", AutoSize = true }, 0, 8);
            var columnInput = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = Padding.Empty };
            columnInput.Controls.Add(keepColumn);
            keepColumnLetter.AutoSize = true;
            keepColumnLetter.Margin = new Padding(8, 5, 0, 0);
            columnInput.Controls.Add(keepColumnLetter);
            panel.Controls.Add(columnInput, 1, 8);
            panel.Controls.Add(new Label { Text = "Proposed retained boundary", AutoSize = true }, 0, 9);
            proposal.AutoSize = true; panel.Controls.Add(proposal, 1, 9);
            var heading = new Label { Text = "Content outside main area", AutoSize = true,
                Font = new Font(Font, FontStyle.Bold), Margin = new Padding(3, 10, 3, 5) };
            panel.SetColumnSpan(heading, 2); panel.Controls.Add(heading, 0, 10);
            var blocks = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true,
                AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect, ScrollBars = ScrollBars.Both };
            blocks.ColumnHeadersHeight = Math.Max(32, Font.Height + 14);
            blocks.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
            blocks.Columns.Add("Range", "Range");
            blocks.Columns.Add("Cells", "Cells");
            blocks.Columns.Add("Content", "Content");
            blocks.Columns[0].MinimumWidth = 190;
            blocks.Columns[1].MinimumWidth = 90;
            blocks.Columns[2].MinimumWidth = 200;
            blocks.Columns[0].FillWeight = 130;
            blocks.Columns[1].FillWeight = 70;
            blocks.Columns[2].FillWeight = 140;
            blocks.MinimumSize = new Size(480, blocks.ColumnHeadersHeight + blocks.RowTemplate.Height * 4 + 3);
            foreach (ContentBlock block in detected.Regions.Distant)
                blocks.Rows.Add(Address(block.Bounds), block.Cells, block.ContentDescription);
            panel.SetColumnSpan(blocks, 2); panel.Controls.Add(blocks, 0, 11);
            keepRow.ValueChanged += (s, e) => UpdateProposal();
            keepColumn.ValueChanged += (s, e) => UpdateProposal();
            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill, AutoSize = true };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            optimize.Text = "Optimize"; optimize.AutoSize = true;
            optimize.Click += (s, e) => Optimize();
            buttons.Controls.Add(cancel); buttons.Controls.Add(optimize);
            panel.SetColumnSpan(buttons, 2); panel.Controls.Add(buttons, 0, 12);
            Controls.Add(panel); CancelButton = cancel;
            UpdateProposal();
        }

        private static string ColumnLetters(int column)
        {
            return UsedRangeOptimizerService.Address(1, column).TrimEnd('1');
        }
        private static string Address(CellRectangle area)
        {
            string first = UsedRangeOptimizerService.Address(area.FirstRow, area.FirstColumn);
            string last = UsedRangeOptimizerService.Address(area.LastRow, area.LastColumn);
            return first == last ? first : first + ":" + last;
        }
        private static void AddLine(TableLayoutPanel panel, int row, string title, string value)
        {
            panel.Controls.Add(new Label { Text = title, AutoSize = true }, 0, row);
            panel.Controls.Add(new Label { Text = value, AutoSize = true }, 1, row);
        }
        private void UpdateProposal()
        {
            var plan = detected.Plan;
            proposal.Text = UsedRangeOptimizerService.Address((int)keepRow.Value, (int)keepColumn.Value);
            keepColumnLetter.Text = ColumnLetters((int)keepColumn.Value);
            excessRows.Text = Math.Max(0, plan.Used.LastRow - (int)keepRow.Value).ToString();
            excessColumns.Text = Math.Max(0, plan.Used.LastColumn - (int)keepColumn.Value).ToString();
            optimize.Enabled = (int)keepRow.Value < plan.Used.LastRow || (int)keepColumn.Value < plan.Used.LastColumn;
        }

        private void Optimize()
        {
            int row = (int)keepRow.Value, column = (int)keepColumn.Value;
            bool manual = row != detected.Plan.RetainedRow || column != detected.Plan.RetainedColumn;
            try
            {
                UsedRangeOptimizationPreview preview = service.Prepare(row, column, manual);
                if (!preview.Plan.HasExcess)
                {
                    MessageBox.Show(this, "No excess area remains. Nothing was cleared.", Text);
                    return;
                }
                string warning = "Clear the sheet area beyond "
                    + UsedRangeOptimizerService.Address(preview.Plan.RetainedRow, preview.Plan.RetainedColumn) + "?\n\n"
                    + "This may remove formatting, data validation, comments/notes, conditional formatting, "
                    + "merged-cell properties, and other cell-level properties in that area.\n\n"
                    + "This operation may not appear in Excel's Undo history. Save a backup first.";
                if (!preview.Plan.HasContent) warning = "No cell content detected on this sheet.\n\n" + warning;
                if (MessageBox.Show(this, warning, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
                if (preview.Removal.RequiresConfirmation && !ConfirmContentRemoval(preview.Removal)) return;
                string result = service.Optimize(preview, preview.Removal.RequiresConfirmation);
                MessageBox.Show(this, result, Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.OK; Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private bool ConfirmContentRemoval(ContentRemovalPreview removal)
        {
            using (var dialog = new Form { Text = "Confirm content removal", Font = Font,
                StartPosition = FormStartPosition.CenterParent, Size = new Size(590, 410),
                MinimumSize = new Size(490, 330) })
            {
                var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), RowCount = 3 };
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.Controls.Add(new Label { AutoSize = true, Text = "The selected cutoff is inside detected content.\n"
                    + "The following populated ranges will be removed:" }, 0, 0);
                var details = new StringBuilder();
                foreach (ContentBlock block in removal.Blocks.Take(40))
                    details.AppendLine(Address(block.Bounds) + " - " + block.Cells + " cells ("
                        + block.FormulaCells + " formulas)");
                if (removal.Blocks.Count > 40)
                    details.AppendLine("... and " + (removal.Blocks.Count - 40) + " more ranges");
                details.AppendLine();
                details.AppendLine("Total: " + removal.Cells + " non-empty cells, including "
                    + removal.FormulaCells + " formulas.");
                details.AppendLine("This action may not be undoable.");
                layout.Controls.Add(new TextBox { Text = details.ToString(), Multiline = true, ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill }, 0, 1);
                var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft,
                    Dock = DockStyle.Fill };
                var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
                var proceed = new Button { Text = "Optimize Anyway", AutoSize = true, DialogResult = DialogResult.Yes };
                buttons.Controls.Add(cancel); buttons.Controls.Add(proceed);
                layout.Controls.Add(buttons, 0, 2);
                dialog.Controls.Add(layout);
                dialog.AcceptButton = cancel;
                dialog.CancelButton = cancel;
                return dialog.ShowDialog(this) == DialogResult.Yes;
            }
        }
    }
}
