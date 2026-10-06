using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Globalization;
using FormuRay.Core;

namespace FormuRay.AddIn
{
    internal sealed class HotkeySettingsStore
    {
        private readonly string path;
        internal ExplorerHotkeySettings Current { get; private set; }
        internal HotkeySettingsStore() : this(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FormuRay", "hotkeys.txt")) { }
        internal HotkeySettingsStore(string path)
        {
            this.path = path;
            Current = ExplorerHotkeySettings.Defaults;
            try
            {
                if (File.Exists(path))
                {
                    var lines = File.ReadAllLines(path);
                    if (lines.Length == 2 || lines.Length == 3 || lines.Length == 9 || lines.Length == 11)
                    {
                        Current = new ExplorerHotkeySettings(lines[0], lines[1], lines.Length >= 3 ? lines[2] : "%+v");
                        if (lines.Length == 9)
                            Current = new ExplorerHotkeySettings(lines[0], lines[1], lines[2], new HighlightSettings(ReadStyle(lines, 3), ReadStyle(lines, 5), ReadStyle(lines, 7)));
                        if (lines.Length == 11)
                            Current = new ExplorerHotkeySettings(lines[0], lines[1], lines[2],
                                new HighlightSettings(ReadStyle(lines, 3), ReadStyle(lines, 7), ReadStyle(lines, 9), ReadStyle(lines, 5)));
                    }
                }
            }
            catch (Exception ex) { Debug.WriteLine("FormuRay hotkey settings: " + ex.Message); }
        }
        internal void Save(ExplorerHotkeySettings settings)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                var lines = new System.Collections.Generic.List<string> { settings.Precedents.Canonical, settings.Dependents.Canonical, settings.PasteFormulas.Canonical };
                foreach (var style in new[] { settings.Highlights.HardCode, settings.Highlights.FormulaConsistency,
                    settings.Highlights.Inflow, settings.Highlights.Outflow })
                { lines.Add(style.Rgb.ToString("X6", CultureInfo.InvariantCulture)); lines.Add(style.OpacityPercent.ToString(CultureInfo.InvariantCulture)); }
                File.WriteAllLines(temporary, lines, Encoding.UTF8);
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
                Current = settings; // Runtime keys change only after persistence succeeds.
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
        private static HighlightStyle ReadStyle(string[] lines, int index)
        {
            return new HighlightStyle(int.Parse(lines[index], NumberStyles.HexNumber, CultureInfo.InvariantCulture), int.Parse(lines[index + 1], CultureInfo.InvariantCulture));
        }
    }
}
