using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormuRay.Excel;
using FormuRay.Excel.Audit;
using Office = Microsoft.Office.Core;

namespace FormuRay.AddIn.Ribbon
{
    [ComVisible(true)]
    public sealed class FormuRayRibbon : Office.IRibbonExtensibility, IDisposable
    {
        private readonly Func<HardcodedDetectorService> getDetector;
        private readonly Func<CrossSheetAuditService> getCrossSheetAudit;
        private readonly Func<CrossSheetInflowService> getCrossSheetInflow;
        private readonly Func<FormulaConsistencyService> getFormulaConsistency;
        private readonly Action<bool> toggleFormulaConsistency;
        private readonly Action closeExplorer, showNamedRanges, showKeySettings, pasteFormulas, showUsedRangeOptimizer, showAutoFormatter;
        private readonly Action showPrecedents;
        private readonly Action showDependents;
        private Office.IRibbonUI ribbon;
        private readonly RibbonImages images = new RibbonImages();

        public System.Drawing.Image GetImage(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "tglCrossSheetInflowVsto": return images.Get("Inflow");
                case "tglCrossSheetOutflowVsto": return images.Get("Outflow");
                case "btnFormulaExplorer": return images.Get("FormulaExplorer");
                case "btnPasteFormulasVsto": return images.Get("PasteFormulas");
                case "tglFormulaConsistencyVsto": return images.Get("FormulaConsistency");
                default: return null;
            }
        }
        public void Dispose() { images.Dispose(); }

        public FormuRayRibbon(Func<HardcodedDetectorService> getDetector,
            Func<CrossSheetAuditService> getCrossSheetAudit, Func<CrossSheetInflowService> getCrossSheetInflow,
            Action showPrecedents, Action showDependents, Action closeExplorer, Action showNamedRanges, Action showKeySettings,
            Action pasteFormulas, Action showUsedRangeOptimizer, Action showAutoFormatter,
            Func<FormulaConsistencyService> getFormulaConsistency, Action<bool> toggleFormulaConsistency)
        {
            this.getDetector = getDetector;
            this.getCrossSheetAudit = getCrossSheetAudit;
            this.getCrossSheetInflow = getCrossSheetInflow;
            this.closeExplorer = closeExplorer;
            this.showNamedRanges = showNamedRanges;
            this.showKeySettings = showKeySettings;
            this.pasteFormulas = pasteFormulas;
            this.showUsedRangeOptimizer = showUsedRangeOptimizer;
            this.showAutoFormatter = showAutoFormatter;
            this.getFormulaConsistency = getFormulaConsistency;
            this.toggleFormulaConsistency = toggleFormulaConsistency;
            this.showPrecedents = showPrecedents;
            this.showDependents = showDependents;
        }

        public string GetCustomUI(string ribbonId)
        {
            using (var stream = typeof(FormuRayRibbon).Assembly.GetManifestResourceStream(
                "FormuRay.AddIn.Ribbon.FormuRayRibbon.xml"))
            {
                if (stream == null) throw new InvalidOperationException("FormuRay Ribbon resource is missing.");
                using (var reader = new StreamReader(stream)) return reader.ReadToEnd();
            }
        }

        public void OnLoad(Office.IRibbonUI ribbonUi) { ribbon = ribbonUi; }

        public void OnPrecedents(Office.IRibbonControl control) { RunAuditAction(showPrecedents); InvalidateCrossSheet(); }
        public void OnDependents(Office.IRibbonControl control) { RunAuditAction(showDependents); InvalidateCrossSheet(); }
        public void OnNamedRanges(Office.IRibbonControl control) { RunAuditAction(showNamedRanges); }
        public void OnKeySettings(Office.IRibbonControl control) { RunAuditAction(showKeySettings); }
        public void OnPasteFormulas(Office.IRibbonControl control) { RunAuditAction(pasteFormulas); }
        public void OnUsedRangeOptimizer(Office.IRibbonControl control) { RunAuditAction(showUsedRangeOptimizer); }
        public void OnAutoFormatter(Office.IRibbonControl control) { RunAuditAction(showAutoFormatter); }
        public bool GetFormulaConsistencyPressed(Office.IRibbonControl control)
        {
            try { return getFormulaConsistency().IsActive; }
            catch (Exception ex) { Debug.WriteLine(ex); return false; }
        }
        public void OnFormulaConsistency(Office.IRibbonControl control, bool pressed)
        {
            try { RunAuditAction(() => toggleFormulaConsistency(pressed)); }
            finally { InvalidateCrossSheet(); }
        }
        public bool GetCrossSheetOutflowPressed(Office.IRibbonControl control)
        {
            try { return getCrossSheetAudit().IsActive; }
            catch (Exception ex) { Debug.WriteLine(ex); return false; }
        }
        public void OnCrossSheetOutflow(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                RunAuditAction(() =>
                {
                    if (pressed) { toggleFormulaConsistency(false); getCrossSheetInflow().Clear(); }
                    getCrossSheetAudit().Toggle(pressed);
                });
            }
            finally { InvalidateCrossSheet(); }
        }
        public bool GetCrossSheetInflowPressed(Office.IRibbonControl control)
        {
            try { return getCrossSheetInflow().IsActive; }
            catch (Exception ex) { Debug.WriteLine(ex); return false; }
        }
        public void OnCrossSheetInflow(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                RunAuditAction(() =>
                {
                    if (pressed) { toggleFormulaConsistency(false); closeExplorer(); getCrossSheetAudit().Clear(); }
                    var inflow = getCrossSheetInflow();
                    inflow.Toggle(pressed);
                    if (pressed && inflow.Highlighted < inflow.Matched)
                        MessageBox.Show("Matched: " + inflow.Matched + "\nHighlighted: " + inflow.Highlighted
                            + "\nSome cells could not be highlighted. Gradient fills are preserved without highlighting; unavailable cells are skipped.", "Cross Sheet Inflow",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                });
            }
            finally { InvalidateCrossSheet(); }
        }
        public void InvalidateCrossSheet()
        {
            try
            {
                if (ribbon == null) return;
                ribbon.InvalidateControl("tglCrossSheetOutflowVsto");
                ribbon.InvalidateControl("tglCrossSheetInflowVsto");
                ribbon.InvalidateControl("tglFormulaConsistencyVsto");
            }
            catch (Exception ex) { Debug.WriteLine(ex); }
        }
        private static void RunAuditAction(Action action)
        {
            try { action(); }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                string message = ex is InvalidOperationException ? ex.Message : "FormuRay could not complete the audit. Check the active workbook and worksheet.";
#if DEBUG
                message += Environment.NewLine + ex;
#endif
                MessageBox.Show(message, "FormuRay", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        public bool GetHardcodedDetectorPressed(Office.IRibbonControl control)
        {
            try { return getDetector().IsActive; }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return false;
            }
        }

        public void OnHardcodedDetector(Office.IRibbonControl control, bool pressed)
        {
            try { getDetector().Toggle(pressed); }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                string message = ex is InvalidOperationException ? ex.Message
                    : "HardCode Detector could not complete. Check that the workbook and worksheet are available.";
#if DEBUG
                message += Environment.NewLine + ex;
#endif
                MessageBox.Show(message, "FormuRay", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                try
                {
                    if (ribbon != null) ribbon.InvalidateControl("tglHardcodedDetectorVsto");
                }
                catch (Exception ex) { Debug.WriteLine(ex); }
            }
        }
    }
}
