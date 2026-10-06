using System;
using System.Globalization;

namespace FormuRay.Core
{
    public static class ExplorerValueFormatter
    {
        public static string Format(object value, CultureInfo culture = null)
        {
            culture = culture ?? CultureInfo.CurrentCulture;
            if (value == null || value == DBNull.Value) return "";
            if (value is Array) return "<array>";
            if (value is string) return (string)value;
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            if (value is DateTime)
            {
                var date = (DateTime)value;
                return date.ToString(date.TimeOfDay == TimeSpan.Zero ? "d" : "G", culture);
            }
            double number;
            try { number = Convert.ToDouble(value, culture); }
            catch (Exception) { return Convert.ToString(value, culture); }
            if (number == 0) return "0";
            double absolute = Math.Abs(number);
            if (absolute < 0.000001) return number.ToString("0.#####E+00", culture).Replace("E+", "E");
            int magnitude = (int)Math.Floor(Math.Log10(absolute));
            int decimals = absolute >= 100 ? 2 : 6 - magnitude - 1;
            if (absolute < 1) decimals = Math.Max(6, decimals);
            decimals = Math.Max(0, Math.Min(10, decimals));
            double rounded = Math.Round(number, decimals, MidpointRounding.AwayFromZero);
            return rounded == 0 ? "0" : rounded.ToString(decimals == 0 ? "0" : "0." + new string('#', decimals), culture);
        }
        public static string Error(int error)
        {
            switch (error & 0xFFFF)
            {
                case 2000: return "#NULL!";
                case 2007: return "#DIV/0!";
                case 2015: return "#VALUE!";
                case 2023: return "#REF!";
                case 2029: return "#NAME?";
                case 2036: return "#NUM!";
                case 2042: return "#N/A";
                default: return "#ERROR";
            }
        }
    }
}
