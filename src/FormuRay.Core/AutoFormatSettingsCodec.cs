using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace FormuRay.Core
{
    public static class AutoFormatSettingsCodec
    {
        private const string Header = "FormuRay AutoFormatter 1";
        public static string[] Serialize(AutoFormatSettings settings)
        {
            var lines = new List<string> { Header };
            foreach (AutoFormatCellType type in AutoFormatSettings.Types)
            {
                AutoFormatStyle style = settings.Get(type);
                lines.Add(string.Join("\t", new[] {
                    ((int)type).ToString(CultureInfo.InvariantCulture),
                    style.FillRgb.ToString("X6", CultureInfo.InvariantCulture),
                    style.FontRgb.ToString("X6", CultureInfo.InvariantCulture),
                    style.Bold ? "1" : "0", style.Italic ? "1" : "0",
                    Convert.ToBase64String(Encoding.UTF8.GetBytes(style.NumberFormat)),
                    ((int)style.Horizontal).ToString(CultureInfo.InvariantCulture),
                    ((int)style.Vertical).ToString(CultureInfo.InvariantCulture),
                    ((int)style.Border).ToString(CultureInfo.InvariantCulture),
                    style.BorderRgb.ToString("X6", CultureInfo.InvariantCulture),
                    ((int)style.BorderLineStyle).ToString(CultureInfo.InvariantCulture),
                    ((int)style.BorderWeight).ToString(CultureInfo.InvariantCulture) }));
            }
            return lines.ToArray();
        }
        public static AutoFormatSettings Deserialize(string[] lines)
        {
            if (lines == null || lines.Length != 7 || lines[0] != Header)
                throw new FormatException("Invalid AutoFormatter settings header.");
            var styles = new AutoFormatStyle[6];
            for (int i = 0; i < styles.Length; i++)
            {
                string[] fields = lines[i + 1].Split('\t');
                if (fields.Length != 12 || ParseInt(fields[0]) != i) throw new FormatException("Invalid AutoFormatter style row.");
                styles[i] = new AutoFormatStyle(ParseHex(fields[1]), ParseHex(fields[2]), ParseBool(fields[3]),
                    ParseBool(fields[4]), Encoding.UTF8.GetString(Convert.FromBase64String(fields[5])),
                    (AutoFormatHorizontal)ParseInt(fields[6]), (AutoFormatVertical)ParseInt(fields[7]),
                    (AutoFormatBorder)ParseInt(fields[8]), ParseHex(fields[9]),
                    (AutoFormatLineStyle)ParseInt(fields[10]), (AutoFormatBorderWeight)ParseInt(fields[11]));
            }
            return new AutoFormatSettings(styles);
        }
        private static int ParseInt(string value) { return int.Parse(value, NumberStyles.None, CultureInfo.InvariantCulture); }
        private static int ParseHex(string value) { return int.Parse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture); }
        private static bool ParseBool(string value)
        {
            if (value == "1") return true;
            if (value == "0") return false;
            throw new FormatException("Invalid AutoFormatter Boolean field.");
        }
    }
}
