using System;
using System.Windows.Forms;
using FormuRay.Excel.Names;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel.UI
{
    public sealed class NamedRangesManagerController : IDisposable
    {
        private sealed class Owner : IWin32Window { public IntPtr Handle { get; set; } }
        private readonly X.Application application;
        private NamedRangesManagerForm form;
        public NamedRangesManagerController(X.Application application) { this.application = application; }
        public void Show()
        {
            if (form == null || form.IsDisposed) form = new NamedRangesManagerForm(new NamedRangesService(application));
            try
            {
                form.RefreshNames();
                if (!form.Visible) form.Show(new Owner { Handle = new IntPtr(application.Hwnd) });
                form.Activate();
            }
            catch { form.Close(); throw; }
        }
        public void Dispose() { if (form != null) { form.Close(); form.Dispose(); form = null; } }
    }
}
