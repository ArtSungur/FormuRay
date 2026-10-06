using System;
using System.Drawing;
using System.Windows.Forms;
using FormuRay.Core;

namespace FormuRay.Excel.UI
{
    public sealed class AutoFormatterForm : Form
    {
        private readonly AutoFormatterService service;
        private readonly Action<AutoFormatSettings> save;
        private AutoFormatSettings settings;
        private readonly RadioButton worksheet = new RadioButton { Text = "Worksheet", AutoSize = true, Checked = true };
        private readonly RadioButton workbook = new RadioButton { Text = "Workbook", AutoSize = true };
        private readonly RadioButton fill = new RadioButton { Text = "Fill", AutoSize = true, Checked = true };
        private readonly RadioButton fillAndFormat = new RadioButton { Text = "Fill and Format", AutoSize = true };
        private readonly DataGridView styles = new DataGridView();
        private static readonly string[] TypeNames = { "Input %", "Input Value", "Input Date",
            "Formula %", "Formula Value", "Formula Date" };

        public AutoFormatterForm(AutoFormatterService service, AutoFormatSettings settings,
            Action<AutoFormatSettings> save)
        {
            this.service = service; this.settings = settings; this.save = save;
            Text = "FormuRay - AutoFormatter";
            Font = SystemFonts.MessageBoxFont;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            ClientSize = new Size(720, 530);
            MinimumSize = new Size(650, 500);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12),
                ColumnCount = 1, RowCount = 2 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var content = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1,
                RowCount = 3, Margin = Padding.Empty };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            content.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            content.Controls.Add(ChoiceGroup("Scope", worksheet, workbook), 0, 0);
            content.Controls.Add(ChoiceGroup("Mode", fill, fillAndFormat), 0, 1);
            styles.Dock = DockStyle.Fill;
            styles.ReadOnly = true; styles.AllowUserToAddRows = false; styles.AllowUserToDeleteRows = false;
            styles.RowHeadersVisible = false; styles.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            styles.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            styles.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            styles.ColumnHeadersHeight = Math.Max(32, Font.Height + 14);
            styles.RowTemplate.Height = Math.Max(36, Font.Height + 18);
            styles.Columns.Add(new DataGridViewTextBoxColumn { Name = "Type", HeaderText = "Type", MinimumWidth = 140, FillWeight = 80 });
            styles.Columns.Add(new DataGridViewTextBoxColumn { Name = "Preview", HeaderText = "Preview", MinimumWidth = 260, FillWeight = 190 });
            styles.Columns.Add(new DataGridViewButtonColumn { Name = "Edit", HeaderText = "Edit Style", MinimumWidth = 120,
                FillWeight = 80, Text = "Edit Style...", UseColumnTextForButtonValue = true });
            styles.MinimumSize = new Size(520, styles.ColumnHeadersHeight + styles.RowTemplate.Height * 6 + 4);
            styles.CellContentClick += (s, e) => { if (e.RowIndex >= 0 && e.ColumnIndex == 2) Edit((AutoFormatCellType)e.RowIndex); };
            content.Controls.Add(styles, 0, 2);
            var buttons = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false,
                Margin = new Padding(0, 10, 0, 0) };
            var close = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.Cancel };
            var resetFormatting = new Button { Text = "Reset Formatting", AutoSize = true };
            var format = new Button { Text = "Format", AutoSize = true };
            var resetStyles = new Button { Text = "Reset Styles", AutoSize = true };
            format.Font = new Font(Font, FontStyle.Bold);
            foreach (Button button in new[] { format, resetStyles, resetFormatting, close })
                button.Margin = new Padding(0, 0, 10, 0);
            format.Click += (s, e) => Run(() => service.Format(Scope, Mode, settings));
            resetFormatting.Click += (s, e) => ResetFormatting();
            resetStyles.Click += (s, e) => ResetStyles();
            buttons.Controls.Add(format); buttons.Controls.Add(resetStyles);
            buttons.Controls.Add(resetFormatting); buttons.Controls.Add(close);
            root.Controls.Add(content, 0, 0);
            root.Controls.Add(buttons, 0, 1);
            Controls.Add(root); AcceptButton = format; CancelButton = close;
            RenderStyles();
        }

        private AutoFormatScope Scope { get { return worksheet.Checked ? AutoFormatScope.Worksheet : AutoFormatScope.Workbook; } }
        private AutoFormatMode Mode { get { return fill.Checked ? AutoFormatMode.Fill : AutoFormatMode.FillAndFormat; } }
        private static GroupBox ChoiceGroup(string title, params RadioButton[] choices)
        {
            var group = new GroupBox { Text = title, Dock = DockStyle.Fill, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(12, 8, 12, 10),
                Margin = new Padding(0, 0, 0, 7) };
            var row = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false, Margin = Padding.Empty };
            foreach (RadioButton choice in choices)
            { choice.Margin = new Padding(4, 4, 24, 4); row.Controls.Add(choice); }
            group.Controls.Add(row);
            return group;
        }
        private void RenderStyles()
        {
            styles.Rows.Clear();
            foreach (AutoFormatCellType type in AutoFormatSettings.Types)
            {
                AutoFormatStyle style = settings.Get(type);
                int row = styles.Rows.Add(TypeNames[(int)type], "Aa 123  |  " + style.NumberFormat);
                var preview = styles.Rows[row].Cells[1].Style;
                preview.BackColor = Rgb(style.FillRgb);
                preview.ForeColor = Rgb(style.FontRgb);
                preview.SelectionBackColor = preview.BackColor;
                preview.SelectionForeColor = preview.ForeColor;
                preview.Font = new Font(Font, (style.Bold ? FontStyle.Bold : FontStyle.Regular)
                    | (style.Italic ? FontStyle.Italic : FontStyle.Regular));
                styles.Rows[row].Cells[1].ToolTipText = style.NumberFormat;
            }
        }
        private static Color Rgb(int rgb) { return Color.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255); }
        private void Edit(AutoFormatCellType type)
        {
            using (var editor = new AutoFormatStyleEditor(TypeNames[(int)type], settings.Get(type)))
            {
                if (editor.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    var next = settings.With(type, editor.Result);
                    save(next); settings = next; RenderStyles();
                }
                catch (Exception ex) { ShowError(ex); }
            }
        }
        private void ResetStyles()
        {
            if (MessageBox.Show(this, "Restore all six AutoFormatter styles to FormuRay defaults?", Text,
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.OK) return;
            try { var defaults = AutoFormatSettings.Defaults; save(defaults); settings = defaults; RenderStyles(); }
            catch (Exception ex) { ShowError(ex); }
        }
        private void ResetFormatting()
        {
            const string warning = "Reset Formatting will clear formatting from input and formula cells in the selected scope. "
                + "Existing fills, fonts, number formats, borders and alignment may be removed. "
                + "This operation may not be undoable.";
            if (MessageBox.Show(this, warning, Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2) != DialogResult.OK) return;
            Run(() => service.ResetFormatting(Scope));
        }
        private void Run(Func<string> operation)
        {
            try { MessageBox.Show(this, operation(), Text, MessageBoxButtons.OK, MessageBoxIcon.Information); }
            catch (Exception ex) { ShowError(ex); }
        }
        private void ShowError(Exception ex)
        {
            MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }

    internal sealed class AutoFormatStyleEditor : Form
    {
        private readonly Button fill = new Button(), font = new Button(), borderColor = new Button();
        private readonly CheckBox bold = new CheckBox { Text = "Bold", AutoSize = true };
        private readonly CheckBox italic = new CheckBox { Text = "Italic", AutoSize = true };
        private readonly TextBox numberFormat = new TextBox();
        private readonly ComboBox horizontal = new ComboBox(), vertical = new ComboBox(), border = new ComboBox();
        private readonly ComboBox line = new ComboBox(), weight = new ComboBox();
        private int fillRgb, fontRgb, borderRgb;
        internal AutoFormatStyle Result { get; private set; }

        internal AutoFormatStyleEditor(string typeName, AutoFormatStyle style)
        {
            Text = "Edit Style - " + typeName;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            ClientSize = new Size(510, 480);
            MinimumSize = new Size(460, 440);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12),
                ColumnCount = 1, RowCount = 2 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill,
                ColumnCount = 2, RowCount = 10, AutoScroll = true, Margin = Padding.Empty };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 10; i++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            fillRgb = style.FillRgb; fontRgb = style.FontRgb; borderRgb = style.BorderRgb;
            AddColor(layout, 0, "Fill Color", fill, () => fillRgb, v => fillRgb = v);
            AddColor(layout, 1, "Font Color", font, () => fontRgb, v => fontRgb = v);
            var flags = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            bold.Checked = style.Bold; italic.Checked = style.Italic;
            flags.Controls.Add(bold); flags.Controls.Add(italic);
            Add(layout, 2, "Font", flags);
            numberFormat.Text = style.NumberFormat; numberFormat.Dock = DockStyle.Fill;
            Add(layout, 3, "Number Format", numberFormat);
            Setup(horizontal, new[] { "Unchanged", "General", "Left", "Center", "Right" }, (int)style.Horizontal);
            Setup(vertical, new[] { "Unchanged", "Top", "Center", "Bottom" }, (int)style.Vertical);
            Setup(border, new[] { "Unchanged", "None", "Outline", "Bottom", "Top", "Top + Bottom", "All Borders" }, (int)style.Border);
            Setup(line, new[] { "Continuous", "Dash", "Dot", "Double" }, (int)style.BorderLineStyle);
            Setup(weight, new[] { "Thin", "Medium", "Thick" }, (int)style.BorderWeight);
            Add(layout, 4, "Horizontal Alignment", horizontal);
            Add(layout, 5, "Vertical Alignment", vertical);
            Add(layout, 6, "Border Preset", border);
            AddColor(layout, 7, "Border Color", borderColor, () => borderRgb, v => borderRgb = v);
            Add(layout, 8, "Border Line Style", line);
            Add(layout, 9, "Border Weight", weight);
            var buttons = new FlowLayoutPanel { AutoSize = false, Height = 48,
                MinimumSize = new Size(0, 48), Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft, WrapContents = false,
                Margin = new Padding(0, 10, 0, 0) };
            var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
            var ok = new Button { Text = "OK", AutoSize = true };
            ok.Click += (s, e) => Save();
            buttons.Controls.Add(cancel); buttons.Controls.Add(ok);
            root.Controls.Add(layout, 0, 0);
            root.Controls.Add(buttons, 0, 1);
            Controls.Add(root); AcceptButton = ok; CancelButton = cancel;
            FontChanged += (s, e) => FitInitialHeight(root, layout, buttons);
            FitInitialHeight(root, layout, buttons);
        }
        private void FitInitialHeight(TableLayoutPanel root, TableLayoutPanel fields, FlowLayoutPanel actions)
        {
            root.PerformLayout(); fields.PerformLayout();
            int fieldsBottom = 0;
            foreach (Control control in fields.Controls) fieldsBottom = Math.Max(fieldsBottom, control.Bottom);
            int needed = root.Padding.Vertical + fieldsBottom + fields.Padding.Bottom
                + actions.GetPreferredSize(Size.Empty).Height + actions.Margin.Vertical + 8;
            int available = Screen.FromControl(this).WorkingArea.Height - 60;
            if (ClientSize.Height < needed && available > ClientSize.Height)
                ClientSize = new Size(ClientSize.Width, Math.Min(needed, available));
        }
        private static void Add(TableLayoutPanel layout, int row, string caption, Control control)
        {
            layout.Controls.Add(new Label { Text = caption, AutoSize = true, Margin = new Padding(3, 8, 3, 8) }, 0, row);
            control.Margin = new Padding(3, 3, 3, 6);
            layout.Controls.Add(control, 1, row);
        }
        private static void Setup(ComboBox box, string[] names, int selected)
        {
            box.DropDownStyle = ComboBoxStyle.DropDownList;
            box.Items.AddRange(names);
            box.SelectedIndex = selected;
            box.Dock = DockStyle.Fill;
        }
        private void AddColor(TableLayoutPanel layout, int row, string caption, Button button,
            Func<int> get, Action<int> set)
        {
            button.AutoSize = true;
            Action update = () =>
            {
                int rgb = get(), red = (rgb >> 16) & 255, green = (rgb >> 8) & 255, blue = rgb & 255;
                button.Text = "#" + rgb.ToString("X6");
                button.BackColor = Color.FromArgb(red, green, blue);
                button.ForeColor = (299 * red + 587 * green + 114 * blue) / 1000 < 140
                    ? Color.White : Color.Black;
            };
            update();
            button.Click += (s, e) =>
            {
                using (var picker = new ColorDialog { Color = button.BackColor, FullOpen = true })
                    if (picker.ShowDialog(this) == DialogResult.OK)
                    { set((picker.Color.R << 16) | (picker.Color.G << 8) | picker.Color.B); update(); }
            };
            Add(layout, row, caption, button);
        }
        private void Save()
        {
            try
            {
                Result = new AutoFormatStyle(fillRgb, fontRgb, bold.Checked, italic.Checked, numberFormat.Text,
                    (AutoFormatHorizontal)horizontal.SelectedIndex, (AutoFormatVertical)vertical.SelectedIndex,
                    (AutoFormatBorder)border.SelectedIndex, borderRgb, (AutoFormatLineStyle)line.SelectedIndex,
                    (AutoFormatBorderWeight)weight.SelectedIndex);
                DialogResult = DialogResult.OK; Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}
