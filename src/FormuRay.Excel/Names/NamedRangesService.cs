using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using FormuRay.Core;
using FormuRay.Excel.Audit;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel.Names
{
    public sealed class NameCleanupResult
    {
        public int Deleted { get; internal set; }
        public int Failed { get; internal set; }
    }

    public sealed class NamedRangesService
    {
        private readonly X.Application application;
        private X.Workbook workbook;
        private Dictionary<string, X.Name> names = new Dictionary<string, X.Name>(StringComparer.OrdinalIgnoreCase);
        public string WorkbookName { get; private set; }
        public NamedRangesService(X.Application application) { this.application = application; }

        public List<NamedRangeInfo> Refresh(bool evaluateValues = true)
        {
            workbook = null;
            names.Clear();
            WorkbookName = "none";
            var active = application.ActiveWorkbook;
            if (active == null || active.IsAddin)
                throw new InvalidOperationException("Activate a user workbook, then open or refresh Named Ranges Manager.");
            workbook = active;
            WorkbookName = active.Name;
            var sheets = ExcelNameReader.Sheets(active);
            var result = new List<NamedRangeInfo>();
            ExcelNameReader.Visit(active, (name, hint) =>
            {
                string identity = ExcelNameReader.Identity(name);
                bool visible = true;
                try { visible = name.Visible; } catch (COMException) { } // Same safe default as VBA.
                var row = new NamedRangeInfo { Identity = identity, Name = NamedRangeRules.DisplayName(identity),
                    Workbook = WorkbookName, Scope = ExcelNameReader.Scope(name, identity, hint, sheets),
                    RefersTo = ExcelNameReader.Text(() => name.RefersTo), RefersToLocal = ExcelNameReader.Text(() => name.RefersToLocal),
                    NameValue = ExcelNameReader.Text(() => name.Value), Visible = visible };
                row.Value = evaluateValues ? EvaluatedValue(name, row) : "";
                result.Add(row);
                names[identity] = name;
            });
            return result;
        }

        private string EvaluatedValue(X.Name name, NamedRangeInfo info)
        {
            X.Range range = null;
            try
            {
                range = name.RefersToRange;
                if (range != null)
                {
                    double count = Convert.ToDouble(range.CountLarge);
                    return count == 1 ? ValueText(range.Value2) : "[Range: " + count + " cells]";
                }
            }
            catch (COMException) { }
            finally { ComReference.Release(range); }
            try
            {
                object evaluated = application.Evaluate(info.Identity);
                // Some Excel COM builds return #NAME? for a dynamic Name although VBA resolves it.
                // Evaluate its existing definition in the same scope; this is display-only, never cleanup classification.
                int error = evaluated is int ? (int)evaluated : evaluated is ErrorWrapper ? ((ErrorWrapper)evaluated).ErrorCode : 0;
                if ((error & 0xFFFF) == 2029 && !string.IsNullOrEmpty(info.RefersTo))
                {
                    try
                    {
                        evaluated = info.Scope != "Workbook"
                            ? ((X.Worksheet)workbook.Worksheets[info.Scope]).Evaluate(info.RefersTo)
                            : application.Evaluate(info.RefersTo);
                    }
                    catch (COMException) { }
                }
                string value = ValueText(evaluated);
                if (value.Length > 0) return value;
            }
            catch (COMException) { }
            return info.NameValue;
        }
        private static string ValueText(object value)
        {
            if (value == null || value == DBNull.Value) return "";
            if (value is Array) return "[Array]";
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            if (value is int || value is ErrorWrapper)
            {
                string error = ExplorerValueFormatter.Error(value is int ? (int)value : ((ErrorWrapper)value).ErrorCode);
                return error == "#REF!" || error == "#VALUE!" ? error : "#ERROR";
            }
            var range = value as X.Range;
            if (range != null)
            {
                try { return Convert.ToDouble(range.CountLarge) == 1 ? ValueText(range.Value2) : "[Range: " + range.CountLarge + " cells]"; }
                finally { ComReference.Release(range); }
            }
            return Convert.ToString(value);
        }

        private void EnsureAvailable()
        {
            try
            {
                if (workbook != null && application.Workbooks.Cast<X.Workbook>().Any(b => b == workbook) && !workbook.IsAddin) return;
            }
            catch (COMException) { }
            throw new InvalidOperationException("The displayed workbook is no longer available. Refresh the list.");
        }
        public NameCleanupResult Delete(IList<NamedRangeInfo> confirmed, NameCleanupKind kind)
        {
            EnsureAvailable();
            var result = new NameCleanupResult();
            // Only delete the names in the confirmed snapshot; never add new matches after confirmation.
            foreach (var item in confirmed)
            {
                X.Name name;
                if (!names.TryGetValue(item.Identity, out name)) { result.Failed++; continue; }
                try
                {
                    var current = new NamedRangeInfo { RefersTo = Convert.ToString(name.RefersTo),
                        RefersToLocal = ExcelNameReader.Text(() => name.RefersToLocal),
                        NameValue = ExcelNameReader.Text(() => name.Value), Visible = name.Visible };
                    if (!NamedRangeRules.Matches(current, kind)) { result.Failed++; continue; }
                    name.Delete();
                    result.Deleted++;
                }
                catch (COMException) { result.Failed++; }
                catch (InvalidComObjectException) { result.Failed++; }
            }
            return result;
        }
        public void Navigate(NamedRangeInfo item)
        {
            EnsureAvailable();
            X.Name name;
            if (item == null || !names.TryGetValue(item.Identity, out name)) return;
            try
            {
                var range = name.RefersToRange;
                if (range == null) throw new InvalidOperationException("This name does not refer to an available range.");
                try { new ExcelAuditContext(application).Navigate(range); }
                finally { ComReference.Release(range); }
            }
            catch (COMException) { throw new InvalidOperationException("This name does not refer to an available range. It may be broken, a constant, or on an unavailable sheet."); }
        }
        public void Clear() { names.Clear(); workbook = null; WorkbookName = "none"; }
    }
}
