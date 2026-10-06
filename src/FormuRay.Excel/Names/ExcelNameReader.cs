using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using FormuRay.Core;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel.Names
{
    internal static class ExcelNameReader
    {
        internal static string Text(Func<object> read)
        {
            try { return Convert.ToString(read()); }
            catch (COMException) { return ""; }
        }
        internal static void Visit(X.Workbook book, Action<X.Name, string> visit)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Action<X.Names, string> scan = (names, hint) =>
            {
                try
                {
                    foreach (X.Name name in names)
                    {
                        try
                        {
                            string identity = Identity(name);
                            if (identity.Length > 0 && seen.Add(identity)) visit(name, hint);
                        }
                        catch (COMException ex) { Debug.WriteLine(ex); }
                    }
                }
                catch (COMException ex) { Debug.WriteLine(ex); }
            };
            scan(book.Names, "");
            foreach (X.Worksheet sheet in book.Worksheets) scan(sheet.Names, sheet.Name);
        }
        internal static string Identity(X.Name name)
        {
            string full = Text(() => name.Name);
            return full.Length > 0 ? full : Text(() => name.NameLocal);
        }
        internal static List<string> Sheets(X.Workbook book)
        {
            return book.Worksheets.Cast<X.Worksheet>().Select(sheet => sheet.Name).ToList();
        }
        internal static string Scope(X.Name name, string identity, string hint, List<string> sheets)
        {
            if (hint.Length == 0)
            {
                try { var sheet = name.Parent as X.Worksheet; if (sheet != null) hint = sheet.Name; }
                catch (COMException) { }
            }
            return NamedRangeRules.Scope(identity, hint, sheets);
        }
        internal static List<InflowName> InflowNames(X.Workbook book, List<string> sheets)
        {
            var result = new List<InflowName>();
            Visit(book, (name, hint) =>
            {
                string full = Identity(name), scope = Scope(name, full, hint, sheets);
                var item = new InflowName { FullName = full, LocalName = Text(() => name.NameLocal),
                    ScopeSheet = full.Contains("!") && sheets.Contains(scope, StringComparer.OrdinalIgnoreCase) ? scope : "" };
                string refers = Text(() => name.RefersTo);
                if (CrossSheetInflowAnalyzer.CanResolveName(refers))
                {
                    X.Range range = null;
                    try
                    {
                        range = name.RefersToRange;
                        if (range != null && ((X.Workbook)range.Worksheet.Parent).FullName == book.FullName)
                            item.TargetSheet = range.Worksheet.Name;
                    }
                    catch (COMException) { }
                    finally { ComReference.Release(range); }
                }
                result.Add(item);
            });
            return result;
        }
    }
}
