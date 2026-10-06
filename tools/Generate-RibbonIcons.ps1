# Reproducible code-native Ribbon artwork. Render at 4x, then downsample to a
# transparent 64px PNG (Office scales it for 32px large buttons / 125-200% DPI).
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
public static class FormuRayIconArt {
    public static void Render(string directory) {
        foreach (string name in new[] { "Inflow", "Outflow", "FormulaExplorer", "PasteFormulas", "FormulaConsistency" }) {
            using (var large = new Bitmap(256, 256, PixelFormat.Format32bppArgb))
            using (var g = Graphics.FromImage(large)) {
                g.SmoothingMode = SmoothingMode.AntiAlias; g.ScaleTransform(4, 4);
                using (var ink = new Pen(Color.FromArgb(55, 73, 89), 2.5f))
                using (var grid = new Pen(Color.FromArgb(165, 180, 191), 1.6f))
                using (var white = new SolidBrush(Color.FromArgb(248, 251, 253)))
                using (var mint = new SolidBrush(Color.FromArgb(209, 235, 220)))
                using (var blue = new SolidBrush(Color.FromArgb(207, 226, 243))) {
                    ink.LineJoin = LineJoin.Round; ink.StartCap = ink.EndCap = LineCap.Round;
                    if (name == "Inflow" || name == "Outflow") {
                        bool incoming = name == "Inflow";
                        g.FillRectangle(white, 23, 7, 34, 49); g.DrawRectangle(ink, 23, 7, 34, 49);
                        g.FillRectangle(incoming ? mint : blue, 25, 9, 30, 10);
                        g.DrawLine(grid, 24, 20, 56, 20); g.DrawLine(grid, 24, 31, 56, 31);
                        g.DrawLine(grid, 24, 43, 56, 43); g.DrawLine(grid, 40, 20, 40, 55);
                        using (var arrow = new SolidBrush(incoming ? Color.FromArgb(41, 139, 100) : Color.FromArgb(49, 113, 184))) {
                            PointF[] points = incoming
                                ? new[] { new PointF(5,29),new PointF(29,29),new PointF(29,22),new PointF(43,36),new PointF(29,50),new PointF(29,43),new PointF(5,43) }
                                : new[] { new PointF(39,29),new PointF(18,29),new PointF(18,22),new PointF(4,36),new PointF(18,50),new PointF(18,43),new PointF(39,43) };
                            using (var halo = new Pen(Color.White, 3)) g.DrawPolygon(halo, points);
                            g.FillPolygon(arrow, points);
                        }
                    } else if (name == "FormulaExplorer") {
                        g.FillRectangle(white, 7, 6, 34, 39); g.DrawRectangle(ink, 7, 6, 34, 39);
                        using (var font = new Font("Segoe UI", 18, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Pixel))
                        using (var text = new SolidBrush(Color.FromArgb(75, 94, 150))) g.DrawString("fx", font, text, 11, 9);
                        g.DrawLine(grid, 13, 36, 34, 36);
                        using (var lens = new SolidBrush(Color.FromArgb(227, 235, 249))) g.FillEllipse(lens, 28, 28, 23, 23);
                        g.DrawEllipse(ink, 28, 28, 23, 23);
                        using (var handle = new Pen(Color.FromArgb(75, 94, 150), 6)) { handle.EndCap = LineCap.Round; g.DrawLine(handle, 49, 49, 58, 58); }
                        g.DrawLine(grid, 34, 40, 45, 40); g.DrawLine(grid, 40, 34, 40, 46);
                    } else if (name == "FormulaConsistency") {
                        g.FillRectangle(white, 5, 8, 54, 48); g.DrawRectangle(ink, 5, 8, 54, 48);
                        g.FillRectangle(blue, 7, 10, 50, 12);
                        using (var amber = new SolidBrush(Color.FromArgb(246, 218, 164)))
                            g.FillRectangle(amber, 38, 24, 10, 18);
                        for (int x = 17; x <= 48; x += 10) g.DrawLine(grid, x, 23, x, 55);
                        g.DrawLine(grid, 6, 23, 58, 23); g.DrawLine(grid, 6, 43, 58, 43);
                        using (var text = new SolidBrush(Color.FromArgb(75, 94, 150))) {
                            using (var font = new Font("Segoe UI", 12, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Pixel))
                                g.DrawString("fx", font, text, 10, 9);
                            using (var mark = new Pen(Color.FromArgb(179, 116, 38), 3)) {
                                g.DrawLine(mark, 41, 29, 45, 37);
                                g.DrawLine(mark, 45, 29, 41, 37);
                            }
                        }
                    } else {
                        g.FillRectangle(blue, 8, 6, 30, 39); g.DrawRectangle(ink, 8, 6, 30, 39);
                        g.FillRectangle(white, 22, 18, 33, 40); g.DrawRectangle(ink, 22, 18, 33, 40);
                        g.FillRectangle(blue, 30, 14, 17, 8); g.DrawRectangle(ink, 30, 14, 17, 8);
                        using (var font = new Font("Segoe UI", 20, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Pixel))
                        using (var text = new SolidBrush(Color.FromArgb(49, 113, 184))) g.DrawString("fx", font, text, 26, 29);
                    }
                }
                using (var output = new Bitmap(64, 64, PixelFormat.Format32bppArgb)) {
                    using (var final = Graphics.FromImage(output)) { final.InterpolationMode = InterpolationMode.HighQualityBicubic; final.DrawImage(large, 0, 0, 64, 64); }
                    output.Save(System.IO.Path.Combine(directory, name + ".png"), ImageFormat.Png);
                }
            }
        }
    }
}
"@
$iconDirectory = Join-Path (Split-Path $PSScriptRoot -Parent) 'src\FormuRay.AddIn\Ribbon\Icons'
New-Item -ItemType Directory -Path $iconDirectory -Force | Out-Null
[FormuRayIconArt]::Render($iconDirectory)
