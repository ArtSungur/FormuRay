using System;
using System.Diagnostics;
using FormuRay.AddIn.Ribbon;
using FormuRay.Excel;
using FormuRay.Core;
using FormuRay.Excel.Audit;
using FormuRay.Excel.UI;
using Office = Microsoft.Office.Core;

namespace FormuRay.AddIn
{
    public partial class ThisAddIn
    {
        private HardcodedDetectorService hardcodedDetector;
        private FormulaConsistencyService formulaConsistency;
        private CrossSheetAuditService crossSheetAudit;
        private FormulaExplorerController explorer;
        private readonly ExplorerLayoutSettingsStore explorerLayout = new ExplorerLayoutSettingsStore();
        private ExplorerKeyboardShortcuts explorerShortcuts;
        private CrossSheetInflowService crossSheetInflow;
        private NamedRangesManagerController namedRanges;
        private PasteFormulasService pasteFormulas;
        private readonly HotkeySettingsController hotkeys = new HotkeySettingsController();
        private readonly AutoFormatSettingsStore autoFormatStyles = new AutoFormatSettingsStore();
        private FormuRayRibbon formuRayRibbon;

        private PasteFormulasService GetPasteFormulas()
        {
            return pasteFormulas ?? (pasteFormulas = new PasteFormulasService(Application));
        }

        private CrossSheetInflowService GetCrossSheetInflow()
        {
            return crossSheetInflow ?? (crossSheetInflow = new CrossSheetInflowService(Application, () => hotkeys.Current.Highlights.Inflow));
        }

        private void ShowNamedRanges()
        {
            if (namedRanges == null) namedRanges = new NamedRangesManagerController(Application);
            namedRanges.Show();
        }

        private void ShowUsedRangeOptimizer()
        {
            using (var service = new UsedRangeOptimizerService(Application))
            using (var form = new UsedRangeOptimizerForm(service))
                form.ShowDialog(new ExcelWindowOwner(new IntPtr(Application.Hwnd)));
        }

        private void ShowAutoFormatter()
        {
            using (var form = new AutoFormatterForm(new AutoFormatterService(Application),
                autoFormatStyles.Current, autoFormatStyles.Save))
                form.ShowDialog(new ExcelWindowOwner(new IntPtr(Application.Hwnd)));
        }

        private FormulaConsistencyService GetFormulaConsistency()
        {
            return formulaConsistency ?? (formulaConsistency = new FormulaConsistencyService(Application,
                () => hotkeys.Current.Highlights.FormulaConsistency));
        }

        private void ToggleFormulaConsistency(bool pressed)
        {
            if (!pressed)
            {
                if (formulaConsistency != null) formulaConsistency.Toggle(false);
                return;
            }
            if (crossSheetInflow != null) crossSheetInflow.Clear();
            if (crossSheetAudit != null) crossSheetAudit.Clear();
            CloseExplorer();
            int count = GetFormulaConsistency().Toggle(true);
            string message = count == 0 ? "Formula Consistency: no formula changes found."
                : "Formula Consistency: " + count + " formula changes found.";
            System.Windows.Forms.MessageBox.Show(message, "FormuRay", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private sealed class ExcelWindowOwner : System.Windows.Forms.IWin32Window
        {
            public IntPtr Handle { get; private set; }
            public ExcelWindowOwner(IntPtr handle) { Handle = handle; }
        }

        private void ShowExplorer(ExplorerMode mode)
        {
            if (crossSheetInflow != null) crossSheetInflow.Clear();
            if (formulaConsistency != null) formulaConsistency.Clear();
            try { GetExplorer().Show(mode); }
            finally { if (formuRayRibbon != null) formuRayRibbon.InvalidateCrossSheet(); }
        }

        private void CloseExplorer() { if (explorer != null) explorer.Dispose(); }

        private CrossSheetAuditService GetCrossSheetAudit()
        {
            return crossSheetAudit ?? (crossSheetAudit = new CrossSheetAuditService(Application, () => hotkeys.Current.Highlights.Outflow));
        }

        private FormulaExplorerController GetExplorer()
        {
            return explorer ?? (explorer = new FormulaExplorerController(Application, () => explorerLayout.Current, explorerLayout.Save));
        }

        private HardcodedDetectorService GetHardcodedDetector()
        {
            return hardcodedDetector ?? (hardcodedDetector = new HardcodedDetectorService(Application, () => hotkeys.Current.Highlights.HardCode));
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return formuRayRibbon = new FormuRayRibbon(GetHardcodedDetector, GetCrossSheetAudit, GetCrossSheetInflow,
                () => ShowExplorer(ExplorerMode.Precedents), () => ShowExplorer(ExplorerMode.Dependents),
                CloseExplorer, ShowNamedRanges, () => hotkeys.Show(new IntPtr(Application.Hwnd)),
                () => GetPasteFormulas().Paste(), ShowUsedRangeOptimizer, ShowAutoFormatter,
                GetFormulaConsistency, ToggleFormulaConsistency);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            GetHardcodedDetector();
            try
            {
                explorerShortcuts = new ExplorerKeyboardShortcuts(new IntPtr(Application.Hwnd),
                    () => ShowExplorer(ExplorerMode.Precedents), () => ShowExplorer(ExplorerMode.Dependents), () => hotkeys.Current,
                    () => GetPasteFormulas().Paste(), () => GetPasteFormulas().CaptureCopy(),
                    () => GetPasteFormulas().ConfirmCopy(), () => GetPasteFormulas().Clear());
                hotkeys.Validate = explorerShortcuts.ValidateSettings;
                explorerShortcuts.ValidateSettings(hotkeys.Current);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                System.Windows.Forms.MessageBox.Show(ex.Message + "\nPaste Formulas is also available on the Ribbon.", "FormuRay", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            hotkeys.Validate = null;
            try { if (formuRayRibbon != null) formuRayRibbon.Dispose(); }
            catch (Exception ex) { Debug.WriteLine(ex); }
            try { if (explorerShortcuts != null) explorerShortcuts.Dispose(); }
            catch (Exception ex) { Debug.WriteLine(ex); }
            try { if (pasteFormulas != null) pasteFormulas.Dispose(); }
            catch (Exception ex) { Debug.WriteLine(ex); }
            try { if (namedRanges != null) namedRanges.Dispose(); }
            catch (Exception ex) { Debug.WriteLine(ex); }
            try { if (crossSheetInflow != null) crossSheetInflow.Dispose(); }
            catch (Exception ex) { Debug.WriteLine(ex); }
            try { if (explorer != null) explorer.Dispose(); }
            catch (Exception ex) { Debug.WriteLine(ex); }
            try { if (crossSheetAudit != null) crossSheetAudit.Dispose(); }
            catch (Exception ex) { Debug.WriteLine(ex); }
            try
            {
                if (hardcodedDetector != null) hardcodedDetector.Dispose();
            }
            catch (Exception ex) { Debug.WriteLine(ex); }
            try { if (formulaConsistency != null) formulaConsistency.Dispose(); }
            catch (Exception ex) { Debug.WriteLine(ex); }
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
