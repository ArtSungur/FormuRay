using System;
using System.Windows.Forms;
using FormuRay.Core;
using FormuRay.Excel.UI;

namespace FormuRay.AddIn
{
    internal sealed class HotkeySettingsController
    {
        private sealed class Owner : IWin32Window { public IntPtr Handle { get; set; } }
        private readonly HotkeySettingsStore store = new HotkeySettingsStore();
        internal Action<ExplorerHotkeySettings> Validate { get; set; }
        internal ExplorerHotkeySettings Current { get { return store.Current; } }
        internal void Show(IntPtr owner)
        {
            using (var form = new HotkeySettingsForm(Current, Save))
                form.ShowDialog(new Owner { Handle = owner });
        }
        private void Save(ExplorerHotkeySettings settings)
        {
            if (Validate == null) throw new InvalidOperationException("Excel shortcuts are unavailable. Reload the add-in before changing shortcuts.");
            Validate(settings);
            store.Save(settings); // Hook reads Current; old mapping remains until this succeeds.
        }
    }
}
