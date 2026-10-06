using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using FormuRay.Core;
using X = Microsoft.Office.Interop.Excel;

namespace FormuRay.Excel
{
    public enum AutoFormatScope { Worksheet, Workbook }
    public enum AutoFormatMode { Fill, FillAndFormat }

    public sealed class AutoFormatterService
    {
        private readonly X.Application application;
        public AutoFormatterService(X.Application application) { this.application = application; }

        private sealed class Target
        {
            internal int Row, Column, LastRow, LastColumn;
            internal AutoFormatCellType Type;
            internal Target(int row, int column, int lastRow, int lastColumn, AutoFormatCellType type)
            { Row = row; Column = column; LastRow = lastRow; LastColumn = lastColumn; Type = type; }
            internal long Cells { get { return (long)(LastRow - Row + 1) * (LastColumn - Column + 1); } }
        }
        private sealed class SheetPlan
        {
            internal X.Worksheet Sheet;
            internal string Name;
            internal readonly List<Target> Targets = new List<Target>();
            internal bool Protected;
            internal bool Failed;
        }

        public string Format(AutoFormatScope scope, AutoFormatMode mode, AutoFormatSettings settings)
        {
            if (settings == null) throw new ArgumentNullException("settings");
            return Execute(scope, mode, settings, false);
        }
        public string ResetFormatting(AutoFormatScope scope)
        {
            return Execute(scope, AutoFormatMode.Fill, null, true);
        }

        private string Execute(AutoFormatScope scope, AutoFormatMode mode, AutoFormatSettings settings, bool reset)
        {
            var plans = CaptureSheets(scope);
            int skippedProtected = 0, failed = 0, completed = 0;
            long cells = 0;
            var protectedNames = new List<string>();
            var failedNames = new List<string>();
            try
            {
                // Every sheet is classified before any NumberFormat or other formatting changes.
                foreach (SheetPlan plan in plans)
                {
                    try
                    {
                        if (plan.Sheet.ProtectContents || plan.Sheet.ProtectDrawingObjects || plan.Sheet.ProtectScenarios)
                        { plan.Protected = true; skippedProtected++; protectedNames.Add(plan.Name); continue; }
                        CaptureTargets(plan, reset);
                    }
                    catch (COMException) { plan.Failed = true; failed++; failedNames.Add(plan.Name); }
                }
                using (new ExcelState(application, reset ? "FormuRay: resetting formatting..." : "FormuRay: formatting cells..."))
                {
                    foreach (SheetPlan plan in plans)
                    {
                        if (plan.Protected || plan.Failed) continue;
                        try
                        {
                            foreach (Target target in plan.Targets)
                            {
                                X.Range range = null;
                                try
                                {
                                    range = GetRange(plan.Sheet, target);
                                    if (reset) range.ClearFormats();
                                    else ApplyStyle(range, settings.Get(target.Type), mode);
                                    cells += target.Cells;
                                }
                                finally { ComReference.Release(range); }
                            }
                            completed++;
                        }
                        catch (COMException) { failed++; failedNames.Add(plan.Name); }
                    }
                }
                return (reset ? "Reset Formatting" : "AutoFormatter") + " completed.\n"
                    + "Worksheets processed: " + completed + "\nCells affected: " + cells
                    + "\nProtected worksheets skipped: " + skippedProtected + "\nWorksheets with errors: " + failed
                    + (protectedNames.Count > 0 ? "\nProtected: " + string.Join(", ", protectedNames) : "")
                    + (failedNames.Count > 0 ? "\nErrors: " + string.Join(", ", failedNames) : "")
                    + (failed > 0 ? "\nA worksheet with an error may have been partially processed." : "");
            }
            finally { foreach (SheetPlan plan in plans) ComReference.Release(plan.Sheet); }
        }

        private List<SheetPlan> CaptureSheets(AutoFormatScope scope)
        {
            var result = new List<SheetPlan>();
            if (scope == AutoFormatScope.Worksheet)
            {
                X.Worksheet sheet = application.ActiveSheet as X.Worksheet;
                if (sheet == null) throw new InvalidOperationException("Select a worksheet before using AutoFormatter.");
                result.Add(new SheetPlan { Sheet = sheet, Name = sheet.Name });
                return result;
            }
            X.Workbook book = null;
            X.Sheets sheets = null;
            try
            {
                book = application.ActiveWorkbook;
                if (book == null) throw new InvalidOperationException("Open a workbook before using AutoFormatter.");
                sheets = book.Worksheets; // Excludes chart sheets.
                for (int i = 1; i <= sheets.Count; i++)
                {
                    X.Worksheet sheet = (X.Worksheet)sheets[i];
                    result.Add(new SheetPlan { Sheet = sheet, Name = sheet.Name });
                }
                return result;
            }
            catch { foreach (SheetPlan plan in result) ComReference.Release(plan.Sheet); throw; }
            finally { ComReference.Release(sheets); ComReference.Release(book); }
        }

        private static void CaptureTargets(SheetPlan plan, bool reset)
        {
            X.Range used = null;
            try
            {
                used = plan.Sheet.UsedRange;
                ReadCandidateAreas(used, X.XlCellType.xlCellTypeConstants, X.XlSpecialCellsValue.xlNumbers, false, reset, plan.Targets);
                ReadCandidateAreas(used, X.XlCellType.xlCellTypeFormulas,
                    reset ? Type.Missing : (object)X.XlSpecialCellsValue.xlNumbers, true, reset, plan.Targets);
            }
            finally { ComReference.Release(used); }
        }

        private static void ReadCandidateAreas(X.Range used, X.XlCellType cellType, object valueType,
            bool formula, bool reset, List<Target> targets)
        {
            X.Range found = null;
            X.Areas areas = null;
            try
            {
                try { found = used.SpecialCells(cellType, valueType); }
                catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x800A03EC)) { return; }
                areas = found.Areas;
                for (int i = 1; i <= areas.Count; i++)
                {
                    X.Range area = null;
                    try
                    {
                        area = areas[i];
                        if (reset)
                            targets.Add(new Target(area.Row, area.Column, area.Row + area.Rows.Count - 1,
                                area.Column + area.Columns.Count - 1, AutoFormatCellType.InputValue));
                        else CaptureArea(area, formula, targets);
                    }
                    finally { ComReference.Release(area); }
                }
            }
            finally { ComReference.Release(areas); ComReference.Release(found); }
        }

        private static void CaptureArea(X.Range area, bool formula, List<Target> targets)
        {
            X.Range rows = null, columns = null, cells = null;
            try
            {
                rows = area.Rows; columns = area.Columns;
                int height = rows.Count, width = columns.Count;
                string common = area.NumberFormat as string;
                if (common != null)
                {
                    AutoFormatCellType? kind = AutoFormatClassifier.Classify(formula, 0d, common);
                    if (kind.HasValue)
                        targets.Add(new Target(area.Row, area.Column, area.Row + height - 1,
                            area.Column + width - 1, kind.Value));
                    return;
                }
                cells = area.Cells;
                for (int r = 1; r <= height; r++)
                {
                    AutoFormatCellType? run = null;
                    int runStart = 1;
                    for (int c = 1; c <= width + 1; c++)
                    {
                        AutoFormatCellType? current = null;
                        if (c <= width)
                        {
                            X.Range cell = null;
                            try
                            {
                                cell = (X.Range)cells[r, c];
                                current = AutoFormatClassifier.Classify(formula, 0d, cell.NumberFormat as string);
                            }
                            finally { ComReference.Release(cell); }
                        }
                        if (current == run) continue;
                        if (run.HasValue)
                            targets.Add(new Target(area.Row + r - 1, area.Column + runStart - 1,
                                area.Row + r - 1, area.Column + c - 2, run.Value));
                        run = current; runStart = c;
                    }
                }
            }
            finally { ComReference.Release(cells); ComReference.Release(columns); ComReference.Release(rows); }
        }

        private static X.Range GetRange(X.Worksheet sheet, Target target)
        {
            X.Range cells = null, first = null, last = null;
            try
            {
                cells = sheet.Cells;
                first = (X.Range)cells[target.Row, target.Column];
                last = (X.Range)cells[target.LastRow, target.LastColumn];
                return sheet.Range[first, last];
            }
            finally { ComReference.Release(last); ComReference.Release(first); ComReference.Release(cells); }
        }

        private static int Ole(int rgb)
        {
            return ColorTranslator.ToOle(Color.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255));
        }
        private static void ApplyStyle(X.Range range, AutoFormatStyle style, AutoFormatMode mode)
        {
            X.Interior interior = null; X.Font font = null;
            try
            {
                interior = range.Interior;
                interior.Pattern = X.XlPattern.xlPatternSolid;
                interior.Color = Ole(style.FillRgb);
                if (mode == AutoFormatMode.Fill) return;
                font = range.Font;
                font.Color = Ole(style.FontRgb);
                font.Bold = style.Bold;
                font.Italic = style.Italic;
                range.NumberFormat = style.NumberFormat;
                if (style.Horizontal != AutoFormatHorizontal.Unchanged) range.HorizontalAlignment = Horizontal(style.Horizontal);
                if (style.Vertical != AutoFormatVertical.Unchanged) range.VerticalAlignment = Vertical(style.Vertical);
                ApplyBorders(range, style);
            }
            finally { ComReference.Release(font); ComReference.Release(interior); }
        }
        private static X.XlHAlign Horizontal(AutoFormatHorizontal alignment)
        {
            switch (alignment)
            {
                case AutoFormatHorizontal.General: return X.XlHAlign.xlHAlignGeneral;
                case AutoFormatHorizontal.Left: return X.XlHAlign.xlHAlignLeft;
                case AutoFormatHorizontal.Center: return X.XlHAlign.xlHAlignCenter;
                default: return X.XlHAlign.xlHAlignRight;
            }
        }
        private static X.XlVAlign Vertical(AutoFormatVertical alignment)
        {
            switch (alignment)
            {
                case AutoFormatVertical.Top: return X.XlVAlign.xlVAlignTop;
                case AutoFormatVertical.Center: return X.XlVAlign.xlVAlignCenter;
                default: return X.XlVAlign.xlVAlignBottom;
            }
        }
        private static void ApplyBorders(X.Range range, AutoFormatStyle style)
        {
            if (style.Border == AutoFormatBorder.Unchanged) return;
            X.Borders borders = null;
            try
            {
                borders = range.Borders;
                borders.LineStyle = X.XlLineStyle.xlLineStyleNone;
                if (style.Border == AutoFormatBorder.None) return;
                if (style.Border == AutoFormatBorder.Outline || style.Border == AutoFormatBorder.All)
                {
                    SetBorder(borders, X.XlBordersIndex.xlEdgeLeft, style);
                    SetBorder(borders, X.XlBordersIndex.xlEdgeRight, style);
                }
                if (style.Border == AutoFormatBorder.Top || style.Border == AutoFormatBorder.TopAndBottom
                    || style.Border == AutoFormatBorder.Outline || style.Border == AutoFormatBorder.All)
                    SetBorder(borders, X.XlBordersIndex.xlEdgeTop, style);
                if (style.Border == AutoFormatBorder.Bottom || style.Border == AutoFormatBorder.TopAndBottom
                    || style.Border == AutoFormatBorder.Outline || style.Border == AutoFormatBorder.All)
                    SetBorder(borders, X.XlBordersIndex.xlEdgeBottom, style);
                if (style.Border == AutoFormatBorder.All)
                {
                    SetBorder(borders, X.XlBordersIndex.xlInsideHorizontal, style);
                    SetBorder(borders, X.XlBordersIndex.xlInsideVertical, style);
                }
            }
            finally { ComReference.Release(borders); }
        }
        private static void SetBorder(X.Borders borders, X.XlBordersIndex edge, AutoFormatStyle style)
        {
            X.Border border = null;
            try
            {
                border = borders[edge];
                border.LineStyle = style.BorderLineStyle == AutoFormatLineStyle.Dash ? X.XlLineStyle.xlDash
                    : style.BorderLineStyle == AutoFormatLineStyle.Dot ? X.XlLineStyle.xlDot
                    : style.BorderLineStyle == AutoFormatLineStyle.Double ? X.XlLineStyle.xlDouble
                    : X.XlLineStyle.xlContinuous;
                border.Color = Ole(style.BorderRgb);
                border.Weight = style.BorderWeight == AutoFormatBorderWeight.Medium ? X.XlBorderWeight.xlMedium
                    : style.BorderWeight == AutoFormatBorderWeight.Thick ? X.XlBorderWeight.xlThick
                    : X.XlBorderWeight.xlThin;
            }
            finally { ComReference.Release(border); }
        }

        private sealed class ExcelState : IDisposable
        {
            private readonly X.Application application;
            private readonly bool screenUpdating, enableEvents;
            private readonly object statusBar;
            internal ExcelState(X.Application application, string status)
            {
                this.application = application;
                screenUpdating = application.ScreenUpdating;
                enableEvents = application.EnableEvents;
                statusBar = application.StatusBar;
                try { application.ScreenUpdating = false; application.EnableEvents = false; application.StatusBar = status; }
                catch { Dispose(); throw; }
            }
            public void Dispose()
            {
                try { application.StatusBar = statusBar; }
                finally { try { application.EnableEvents = enableEvents; } finally { application.ScreenUpdating = screenUpdating; } }
            }
        }
    }
}
