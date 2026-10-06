using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using FormuRay.Core;

namespace FormuRay.AddIn
{
    // Uses the same per-user directory and atomic replacement pattern as the other FormuRay stores.
    internal sealed class AutoFormatSettingsStore
    {
        private readonly string path;
        internal AutoFormatSettings Current { get; private set; }
        internal AutoFormatSettingsStore() : this(Path.Combine(Environment.GetFolderPath(
            Environment.SpecialFolder.LocalApplicationData), "FormuRay", "autoformatter.txt")) { }
        internal AutoFormatSettingsStore(string path)
        {
            this.path = path;
            Current = AutoFormatSettings.Defaults;
            try { if (File.Exists(path)) Current = AutoFormatSettingsCodec.Deserialize(File.ReadAllLines(path)); }
            catch (Exception ex) { Debug.WriteLine("FormuRay AutoFormatter settings: " + ex.Message); }
        }
        internal void Save(AutoFormatSettings settings)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllLines(temporary, AutoFormatSettingsCodec.Serialize(settings), Encoding.UTF8);
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
                Current = settings;
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
