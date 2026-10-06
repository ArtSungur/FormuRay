using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using FormuRay.Core;

namespace FormuRay.AddIn
{
    // Same per-user directory and atomic replacement pattern as HotkeySettingsStore.
    internal sealed class ExplorerLayoutSettingsStore
    {
        private readonly string path;
        internal ExplorerColumnWidths Current { get; private set; }
        internal ExplorerLayoutSettingsStore() : this(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FormuRay", "explorer-layout.txt")) { }
        internal ExplorerLayoutSettingsStore(string path)
        {
            this.path = path;
            Current = ExplorerColumnWidths.Defaults;
            try
            {
                if (!File.Exists(path)) return;
                var lines = File.ReadAllLines(path);
                int element, value, type;
                double ratio = .7;
                if (lines.Length >= 4) double.TryParse(lines[3], NumberStyles.Float, CultureInfo.InvariantCulture, out ratio);
                if ((lines.Length == 3 || lines.Length == 4) && int.TryParse(lines[0], out element) && int.TryParse(lines[1], out value) && int.TryParse(lines[2], out type))
                    Current = new ExplorerColumnWidths(element, value, type, ratio);
            }
            catch (Exception ex) { Debug.WriteLine("FormuRay Explorer layout: " + ex.Message); }
        }
        internal void Save(ExplorerColumnWidths widths)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllLines(temporary, new[] { widths.Element.ToString(CultureInfo.InvariantCulture), widths.Value.ToString(CultureInfo.InvariantCulture), widths.Type.ToString(CultureInfo.InvariantCulture), widths.SplitterRatio.ToString("R", CultureInfo.InvariantCulture) }, Encoding.UTF8);
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
                Current = widths;
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
