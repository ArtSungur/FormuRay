using System;
using System.Drawing;
using System.Windows.Forms;
using FormuRay.Core;

namespace FormuRay.Excel.UI
{
    public sealed class HotkeySettingsForm : Form
    {
        private readonly ComboBox precedentsModifier = new ComboBox(), dependentsModifier = new ComboBox(), pasteModifier = new ComboBox();
        private readonly TextBox precedentsKey = new TextBox(), dependentsKey = new TextBox(), pasteKey = new TextBox();
        private readonly Action<ExplorerHotkeySettings> save;
        private readonly Button[] colors = { new Button(), new Button(), new Button(), new Button() };
        private readonly NumericUpDown[] opacities = { new NumericUpDown(), new NumericUpDown(), new NumericUpDown(), new NumericUpDown() };
        public HotkeySettingsForm(ExplorerHotkeySettings settings, Action<ExplorerHotkeySettings> save)
        {
            this.save = save;
            SuspendLayout();
            Text = "FormuRay - Settings";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleDimensions = new SizeF(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(630, 510); MinimumSize = new Size(480, 330);
            FormBorderStyle = FormBorderStyle.Sizable; MaximizeBox = true; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(12) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Margin = Padding.Empty };
            var layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 3, RowCount = 7, Padding = new Padding(2, 2, 18, 12) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 43));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 23));
            for (int i = 0; i < 7; i++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            AddHeading(layout, 0, "Formula Audit Explorer");
            AddRow(layout, 1, "Precedents", precedentsModifier, precedentsKey, false);
            AddRow(layout, 2, "Dependents", dependentsModifier, dependentsKey, false);
            AddRow(layout, 3, "Paste Formulas", pasteModifier, pasteKey, true);
            var help = new Label { Text = "Paste Formulas: modifiers + A-Z, 0-9 or F1-F24 (except F12). Copy cells with Ctrl + C first. Excel's own shortcuts may be overridden while the worksheet has focus.",
                AutoSize = true, Dock = DockStyle.Fill };
            layout.Controls.Add(help, 0, 4); layout.SetColumnSpan(help, 3);
            AddHeading(layout, 5, "Highlight colors");
            string[] names = { "HardCode Detector", "Formula Consistency", "Cross Sheet Inflow", "Cross Sheet Outflow" };
            var highlights = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill, ColumnCount = 5, RowCount = names.Length, Margin = new Padding(3, 0, 3, 0) };
            highlights.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int column = 1; column < 5; column++) highlights.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            for (int row = 0; row < names.Length; row++) highlights.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            for (int i = 0; i < names.Length; i++)
            {
                int index = i;
                highlights.Controls.Add(new Label { Text = names[i], AutoSize = true, Anchor = AnchorStyles.Left,
                    Margin = new Padding(3, 4, 8, 4) }, 0, i);
                colors[i].Size = new Size(48, 24); colors[i].Anchor = AnchorStyles.Left;
                colors[i].Margin = new Padding(4, 4, 12, 4);
                colors[i].UseVisualStyleBackColor = false; colors[i].AccessibleName = names[i] + " color";
                colors[i].Click += (s, e) =>
                {
                    using (var dialog = new ColorDialog { Color = colors[index].BackColor, FullOpen = true })
                        if (dialog.ShowDialog(this) == DialogResult.OK) colors[index].BackColor = dialog.Color;
                };
                highlights.Controls.Add(colors[i], 1, i);
                highlights.Controls.Add(new Label { Text = "Opacity", AutoSize = true, Anchor = AnchorStyles.Left,
                    Margin = new Padding(3, 4, 8, 4) }, 2, i);
                opacities[i].Minimum = 0; opacities[i].Maximum = 100; opacities[i].Width = 60;
                opacities[i].Anchor = AnchorStyles.Left; opacities[i].Margin = new Padding(3, 4, 3, 4);
                opacities[i].AccessibleName = names[i] + " opacity";
                highlights.Controls.Add(opacities[i], 3, i);
                highlights.Controls.Add(new Label { Text = "%", AutoSize = true, Anchor = AnchorStyles.Left,
                    Margin = new Padding(3, 4, 3, 4) }, 4, i);
            }
            layout.Controls.Add(highlights, 0, 6); layout.SetColumnSpan(highlights, 3);
            scroll.Controls.Add(layout); root.Controls.Add(scroll, 0, 0);
            var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true, Margin = new Padding(0, 10, 0, 0) };
            var reset = new Button { Text = "Reset Defaults", AutoSize = true };
            var apply = new Button { Text = "Save", AutoSize = true };
            var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
            buttons.Controls.AddRange(new Control[] { reset, apply, cancel });
            root.Controls.Add(buttons, 0, 1); Controls.Add(root);
            CancelButton = cancel; AcceptButton = apply;
            LoadSettings(settings);
            apply.Click += (s, e) => Attempt(() => { this.save(new ExplorerHotkeySettings(Build(precedentsModifier, precedentsKey), Build(dependentsModifier, dependentsKey), Build(pasteModifier, pasteKey), new HighlightSettings(ReadStyle(0), ReadStyle(2), ReadStyle(3), ReadStyle(1)))); DialogResult = DialogResult.OK; Close(); });
            // Reset edits the form; Save applies all shortcuts together and Cancel discards them.
            reset.Click += (s, e) => LoadSettings(ExplorerHotkeySettings.Defaults);
            ResumeLayout(true);
        }
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // Respect laptop working areas after WinForms applies the host's DPI scaling.
            var working = Screen.FromControl(this).WorkingArea;
            MinimumSize = new Size(Math.Min(MinimumSize.Width, working.Width), Math.Min(MinimumSize.Height, working.Height));
            Size = new Size(Math.Min(Width, working.Width), Math.Min(Height, working.Height));
        }
        private static void AddHeading(TableLayoutPanel layout, int row, string text)
        {
            var label = new Label { Text = text, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(3, 12, 3, 8) };
            layout.Controls.Add(label, 0, row); layout.SetColumnSpan(label, 3);
        }
        private static void AddRow(TableLayoutPanel layout, int row, string label, ComboBox modifier, TextBox key, bool extended)
        {
            modifier.DropDownStyle = ComboBoxStyle.DropDownList; modifier.Dock = DockStyle.Fill;
            modifier.Items.AddRange(new object[] { "Ctrl+", "Alt+", "Shift+", "Ctrl+Shift+", "Ctrl+Alt+", "Alt+Shift+", "Ctrl+Alt+Shift+" });
            key.CharacterCasing = CharacterCasing.Upper; key.MaxLength = extended ? 3 : 1; key.Dock = DockStyle.Fill;
            key.AccessibleName = label + " key"; modifier.AccessibleName = label + " modifiers";
            layout.Controls.Add(new Label { Text = label, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(3, 6, 3, 6) }, 0, row);
            layout.Controls.Add(modifier, 1, row); layout.Controls.Add(key, 2, row);
        }
        private static string Build(ComboBox modifier, TextBox key)
        {
            return modifier.Text.Replace("Ctrl+", "^").Replace("Alt+", "%").Replace("Shift+", "+") + key.Text;
        }
        private static void LoadShortcut(ExplorerHotkey shortcut, ComboBox modifier, TextBox key)
        {
            modifier.SelectedItem = shortcut.ModifierLabel; key.Text = shortcut.KeyName;
        }
        private void LoadSettings(ExplorerHotkeySettings settings)
        {
            var styles = new[] { settings.Highlights.HardCode, settings.Highlights.FormulaConsistency,
                settings.Highlights.Inflow, settings.Highlights.Outflow };
            for (int i = 0; i < styles.Length; i++) { colors[i].BackColor = Color.FromArgb(unchecked((int)0xFF000000) | styles[i].Rgb); opacities[i].Value = styles[i].OpacityPercent; }
            LoadShortcut(settings.Precedents, precedentsModifier, precedentsKey);
            LoadShortcut(settings.Dependents, dependentsModifier, dependentsKey);
            LoadShortcut(settings.PasteFormulas, pasteModifier, pasteKey);
        }
        private HighlightStyle ReadStyle(int index) { return new HighlightStyle(colors[index].BackColor.ToArgb() & 0xFFFFFF, (int)opacities[index].Value); }
        private void Attempt(Action action)
        {
            try { action(); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
        }
    }
}
