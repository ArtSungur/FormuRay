using System;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using FormuRay.Core;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel.Audit
{
    public sealed class ExcelAuditContext
    {
        internal X.Application Application { get; private set; }
        public ExcelAuditContext(X.Application application) { Application = application; }
        public IntPtr WindowHandle
        {
            get
            {
                X.Window window = null;
                try { window = Application.ActiveWindow; return new IntPtr(window == null ? Application.Hwnd : window.Hwnd); }
                finally { ComReference.Release(window); }
            }
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRectangle { internal int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr window, out NativeRectangle rectangle);
        public Rectangle WindowBounds
        {
            get { NativeRectangle rectangle; return GetWindowRect(WindowHandle, out rectangle)
                ? Rectangle.FromLTRB(rectangle.Left, rectangle.Top, rectangle.Right, rectangle.Bottom) : new Rectangle(0, 0, 800, 600); }
        }
        public Rectangle ActiveCellScreenBounds
        {
            get
            {
                X.Range cell = null;
                X.Window window = null;
                X.Pane pane = null;
                try
                {
                    cell = Application.ActiveCell as X.Range;
                    window = Application.ActiveWindow;
                    if (cell != null && window != null)
                    {
                        pane = window.ActivePane;
                        // Excel transforms document points through the active pane's
                        // scroll/zoom to screen pixels. Do not change process DPI awareness
                        // inside an Office add-in; use the host UI thread's coordinate context.
                        int left = pane.PointsToScreenPixelsX((int)Math.Floor(Convert.ToDouble(cell.Left)));
                        int top = pane.PointsToScreenPixelsY((int)Math.Floor(Convert.ToDouble(cell.Top)));
                        int right = pane.PointsToScreenPixelsX((int)Math.Ceiling(Convert.ToDouble(cell.Left) + Convert.ToDouble(cell.Width)));
                        int bottom = pane.PointsToScreenPixelsY((int)Math.Ceiling(Convert.ToDouble(cell.Top) + Convert.ToDouble(cell.Height)));
                        return new Rectangle(left, top, Math.Max(1, right - left), Math.Max(1, bottom - top));
                    }
                }
                catch (COMException) { }
                finally { ComReference.Release(pane); ComReference.Release(window); ComReference.Release(cell); }
                Rectangle owner = WindowBounds;
                return new Rectangle(owner.Left + owner.Width / 2, owner.Top + owner.Height / 2, 1, 1);
            }
        }

        internal static RangeLocation Location(X.Range range)
        {
            var sheet = range.Worksheet;
            var book = (X.Workbook)sheet.Parent;
            return new RangeLocation { Workbook = book.Name, WorkbookFullName = book.FullName, Sheet = sheet.Name, Address = range.Address[false, false, X.XlReferenceStyle.xlA1] };
        }

        internal X.Range Resolve(RangeLocation location)
        {
            if (location == null || string.IsNullOrEmpty(location.Address)) return null;
            try
            {
                var book = Application.Workbooks[location.Workbook];
                if (!string.IsNullOrEmpty(location.WorkbookFullName) && !string.Equals(book.FullName, location.WorkbookFullName, StringComparison.OrdinalIgnoreCase)) return null;
                return ((X.Worksheet)book.Worksheets[location.Sheet]).Range[location.Address];
            }
            catch (COMException) { return null; }
        }

        internal X.Range ResolveReference(FormulaReference reference, X.Worksheet source, bool legacyDefaultWorkbook = false)
        {
            try
            {
                // Navigation can resolve open external books. Closed books are never opened implicitly.
                var book = reference.IsExternal && !legacyDefaultWorkbook ? Application.Workbooks[reference.Workbook] : (X.Workbook)source.Parent;
                var sheet = string.IsNullOrEmpty(reference.Sheet) ? source : (X.Worksheet)book.Worksheets[reference.Sheet];
                return sheet.Range[reference.Address];
            }
            catch (COMException) { return null; }
        }

        internal string Value(X.Range range)
        {
            if (range == null) return "";
            try
            {
                if (Convert.ToDouble(range.CountLarge) != 1) return "<array>";
                object value = range.Value; // VBA uses Value here, preserving dates/currency.
                if (value is int || value is ErrorWrapper)
                {
                    string displayed = Convert.ToString(range.Text);
                    if (displayed.StartsWith("#", StringComparison.Ordinal)) return displayed;
                    return ExplorerValueFormatter.Error(value is int ? (int)value : ((ErrorWrapper)value).ErrorCode);
                }
                return ExplorerValueFormatter.Format(value, CultureInfo.CurrentCulture);
            }
            catch (COMException) { return ""; }
        }

        public string Evaluate(RangeLocation source, string fragment)
        {
            if (string.IsNullOrEmpty(fragment)) return "";
            X.Range cell = Resolve(source);
            if (cell == null) return "";
            string formula = fragment.StartsWith("=", StringComparison.Ordinal) ? fragment : "=" + fragment;
            try
            {
                string result = EvaluatedValue(cell.Worksheet.Evaluate(formula));
                if (result.Length > 0) return result;
            }
            catch (COMException) { }
            try { return EvaluatedValue(Application.Evaluate(formula)); }
            catch (COMException) { return ""; }
        }

        private string EvaluatedValue(object value)
        {
            var range = value as X.Range;
            if (range != null) return Value(range);
            if (value is int) return ExplorerValueFormatter.Error((int)value);
            if (value is ErrorWrapper) return ExplorerValueFormatter.Error(((ErrorWrapper)value).ErrorCode);
            return ExplorerValueFormatter.Format(value);
        }

        internal void Navigate(X.Range range)
        {
            ((X.Workbook)range.Worksheet.Parent).Activate();
            range.Worksheet.Activate();
            var window = Application.ActiveWindow;
            if (window != null)
            {
                X.Range intersection = Application.Intersect(range, window.VisibleRange);
                if (intersection == null || Convert.ToDouble(intersection.CountLarge) != Convert.ToDouble(range.CountLarge))
                {
                    window.ScrollRow = Math.Max(1, range.Row - 5);
                    window.ScrollColumn = Math.Max(1, range.Column - 3);
                }
            }
            range.Select();
        }
    }
}
