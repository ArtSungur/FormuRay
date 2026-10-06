using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using FormuRay.Core;

namespace FormuRay.AddIn
{
    // WH_KEYBOARD is installed ONLY on Excel's UI thread, never globally. OnKey
    // requires an Excel procedure and cannot directly call a managed VSTO method.
    internal sealed class ExplorerKeyboardShortcuts : IDisposable
    {
        private delegate IntPtr KeyboardProc(int code, IntPtr key, IntPtr flags);
        private readonly KeyboardProc callback;
        private readonly Control dispatcher = new Control();
        private readonly Action precedents, dependents, paste, captureCopy, confirmCopy, clearCopy;
        private readonly Func<ExplorerHotkeySettings> settings;
        private readonly uint excelThread;
        private IntPtr hook;
        private bool pending, disposed, capturing;
        private int consumedKey, copyGeneration;

        internal ExplorerKeyboardShortcuts(IntPtr excelWindow, Action precedents, Action dependents,
            Func<ExplorerHotkeySettings> settings, Action paste, Action captureCopy, Action confirmCopy, Action clearCopy)
        {
            this.precedents = precedents; this.dependents = dependents; this.settings = settings;
            this.paste = paste; this.captureCopy = captureCopy; this.confirmCopy = confirmCopy; this.clearCopy = clearCopy;
            IntPtr dispatchWindow = dispatcher.Handle;
            callback = OnKey;
            uint process;
            excelThread = GetWindowThreadProcessId(excelWindow, out process);
            // Never pass zero: that would request a global hook.
            if (excelThread == 0 || excelThread != GetCurrentThreadId())
            {
                dispatcher.Dispose();
                throw new InvalidOperationException("FormuRay shortcuts require Excel's UI thread.");
            }
            hook = SetWindowsHookEx(2, callback, IntPtr.Zero, excelThread);
            if (hook == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                dispatcher.Dispose();
                throw new InvalidOperationException("FormuRay could not activate Excel shortcuts (Windows error " + error + ").");
            }
        }

        internal void ValidateSettings(ExplorerHotkeySettings proposed)
        {
            if (disposed || hook == IntPtr.Zero) throw new InvalidOperationException("Excel shortcuts are unavailable. Reload the add-in before changing shortcuts.");
            // A short-lived registration probes existing Windows ownership; it is released
            // immediately. Actual delivery stays local to Excel, with NO global registration.
            // Excel built-ins/other add-ins using hooks cannot be enumerated by Windows.
            var shortcut = proposed.PasteFormulas;
            uint modifiers = (uint)((shortcut.Alt ? 1 : 0) | (shortcut.Control ? 2 : 0) | (shortcut.Shift ? 4 : 0));
            ushort atom = GlobalAddAtom("FormuRay.HotkeyProbe." + Guid.NewGuid().ToString("N"));
            if (atom == 0) throw new InvalidOperationException("Could not validate the shortcut. The previous shortcut is unchanged.");
            bool registered = false;
            try
            {
                registered = RegisterHotKey(dispatcher.Handle, atom, modifiers, (uint)shortcut.VirtualKey);
                if (!registered) throw new InvalidOperationException(shortcut.DisplayName + " is unavailable or registered by another application. Choose another shortcut. The previous shortcut is unchanged.");
            }
            finally
            {
                if (registered) UnregisterHotKey(dispatcher.Handle, atom);
                GlobalDeleteAtom(atom);
            }
        }

        private bool WorksheetHasFocus()
        {
            uint process;
            if (GetWindowThreadProcessId(GetForegroundWindow(), out process) != excelThread) return false;
            var className = new StringBuilder(64);
            GetClassName(GetFocus(), className, className.Capacity);
            return className.ToString() == "EXCEL7";
        }

        private IntPtr OnKey(int code, IntPtr key, IntPtr flags)
        {
            try
            {
                if (code != 0 || disposed || capturing) return CallNextHookEx(hook, code, key, flags);
                int value = key.ToInt32();
                bool up = (flags.ToInt64() & 0x80000000L) != 0;
                bool repeat = (flags.ToInt64() & 0x40000000L) != 0;
                if (value == consumedKey)
                {
                    if (up) consumedKey = 0;
                    return new IntPtr(1);
                }
                if (up || !WorksheetHasFocus()) return CallNextHookEx(hook, code, key, flags);
                bool control = Down(Keys.ControlKey), alt = Down(Keys.Menu), shift = Down(Keys.ShiftKey);
                if (control && !alt && !shift && (value == 'C' || value == 'X'))
                {
                    if (!repeat)
                    {
                        int generation = ++copyGeneration;
                        if (value == 'X') clearCopy();
                        else
                        {
                            // Bounded, read-only bulk snapshot BEFORE native Copy handles the
                            // key. Deferring the read could capture a rapidly changed Selection.
                            // The guard prevents reentrancy; exceptions never cancel Excel Copy.
                            capturing = true;
                            try { captureCopy(); }
                            finally { capturing = false; }
                            dispatcher.BeginInvoke(new Action(() =>
                            {
                                if (disposed || generation != copyGeneration) return;
                                try { confirmCopy(); } catch (Exception ex) { Debug.WriteLine(ex); clearCopy(); }
                            }));
                        }
                    }
                    return CallNextHookEx(hook, code, key, flags);
                }
                var current = settings();
                Action command = current.Precedents.Matches(value, control, alt, shift) ? precedents
                    : current.Dependents.Matches(value, control, alt, shift) ? dependents
                    : current.PasteFormulas.Matches(value, control, alt, shift) ? paste : null;
                if (command != null)
                {
                    consumedKey = value;
                    if (!pending && !repeat)
                    {
                        pending = true;
                        dispatcher.BeginInvoke(new Action(() =>
                        {
                            try { if (!disposed && WorksheetHasFocus()) command(); }
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex);
                                MessageBox.Show(ex is InvalidOperationException ? ex.Message : "FormuRay could not complete the command. Check the active worksheet.",
                                    "FormuRay", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                            finally { pending = false; }
                        }));
                    }
                    return new IntPtr(1);
                }
            }
            catch (Exception ex) { Debug.WriteLine(ex); }
            return CallNextHookEx(hook, code, key, flags);
        }
        private static bool Down(Keys key) { return (GetKeyState((int)key) & 0x8000) != 0; }
        public void Dispose()
        {
            if (disposed) return;
            disposed = true; ++copyGeneration;
            if (hook != IntPtr.Zero) { UnhookWindowsHookEx(hook); hook = IntPtr.Zero; }
            clearCopy(); dispatcher.Dispose();
        }
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetWindowsHookEx(int hook, KeyboardProc proc, IntPtr module, uint thread);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr key, IntPtr flags);
        [DllImport("user32.dll")] private static extern short GetKeyState(int key);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetFocus();
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int size);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint key);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr window, int id);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern ushort GlobalAddAtom(string text);
        [DllImport("kernel32.dll")] private static extern ushort GlobalDeleteAtom(ushort atom);
    }
}
