using System;
using System.Windows.Forms;
using FormuRay.Core;
using FormuRay.Excel.Audit;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel.UI
{
    public sealed class FormulaExplorerController : IDisposable
    {
        private sealed class ExcelOwner : IWin32Window
        {
            public IntPtr Handle { get; private set; }
            internal ExcelOwner(IntPtr handle) { Handle = handle; }
        }
        private readonly X.Application application;
        private readonly Func<ExplorerColumnWidths> loadWidths;
        private readonly Action<ExplorerColumnWidths> saveWidths;
        private FormulaExplorerForm form;
        private bool refreshing, pending;
        public FormulaExplorerController(X.Application application, Func<ExplorerColumnWidths> loadWidths = null, Action<ExplorerColumnWidths> saveWidths = null)
        {
            this.application = application; this.loadWidths = loadWidths; this.saveWidths = saveWidths;
        }
        public void Show(ExplorerMode mode)
        {
            if (form == null || form.IsDisposed)
            {
                var service = new FormulaExplorerService(application);
                form = new FormulaExplorerForm(service, loadWidths == null ? null : loadWidths(), saveWidths);
                application.SheetSelectionChange += SelectionChanged;
                application.SheetActivate += SheetActivated;
                application.WindowActivate += WindowActivated;
                form.FormClosed += (s, e) =>
                {
                    application.SheetSelectionChange -= SelectionChanged;
                    application.SheetActivate -= SheetActivated;
                    application.WindowActivate -= WindowActivated;
                    pending = false;
                };
            }
            try
            {
                refreshing = true;
                form.LoadMode(mode);
                // OnLoad positions once, after WinForms scaling. Refresh never moves it.
                if (!form.Visible) form.Show(new ExcelOwner(new IntPtr(application.Hwnd)));
                form.Activate();
            }
            catch { form.Close(); throw; }
            finally { refreshing = false; }
        }
        private void SelectionChanged(object sheet, X.Range target) { QueueRefresh(); }
        private void SheetActivated(object sheet) { QueueRefresh(); }
        private void WindowActivated(X.Workbook workbook, X.Window window) { QueueRefresh(); }
        private void QueueRefresh()
        {
            // Navigation intentionally keeps the original audit root. Excel can emit
            // several synchronous sheet/window/selection events for one navigation.
            if (refreshing || pending || form == null || !form.Visible || form.IsNavigating) return;
            var current = form;
            pending = true;
            current.BeginInvoke(new Action(() =>
            {
                pending = false;
                if (current.IsDisposed || current != form) return;
                refreshing = true;
                try { current.RefreshActiveCell(); }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine(ex); }
                finally { refreshing = false; }
            }));
        }
        public void Dispose() { if (form != null) { form.Close(); form.Dispose(); form = null; } }
    }
}
