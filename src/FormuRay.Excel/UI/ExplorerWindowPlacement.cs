using System;
using System.Drawing;
using System.Linq;

namespace FormuRay.Excel.UI
{
    public static class ExplorerWindowPlacement
    {
        public static Rectangle BesideCell(Rectangle cell, Size size, Rectangle working, int gap)
        {
            size = new Size(Math.Min(size.Width, working.Width), Math.Min(size.Height, working.Height));
            int y = cell.Top + (cell.Height - size.Height) / 2;
            if (cell.Right + gap + size.Width <= working.Right)
                return Clamp(new Rectangle(cell.Right + gap, y, size.Width, size.Height), working);
            if (cell.Left - gap - size.Width >= working.Left)
                return Clamp(new Rectangle(cell.Left - gap - size.Width, y, size.Width, size.Height), working);
            // Prefer zero overlap, then the closest clamped alternative (including above/below).
            var candidates = new[] {
                new Rectangle(cell.Right + gap, y, size.Width, size.Height),
                new Rectangle(cell.Left - gap - size.Width, y, size.Width, size.Height),
                new Rectangle(cell.Left, cell.Top - gap - size.Height, size.Width, size.Height),
                new Rectangle(cell.Left, cell.Bottom + gap, size.Width, size.Height) };
            return candidates.Select(r => Clamp(r, working)).OrderBy(r => Overlap(r, cell))
                .ThenBy(r => Math.Abs((long)r.Left - cell.Right) + Math.Abs((long)r.Top - y)).First();
        }
        private static long Overlap(Rectangle a, Rectangle b)
        {
            var intersection = Rectangle.Intersect(a, b);
            return (long)Math.Max(0, intersection.Width) * Math.Max(0, intersection.Height);
        }
        private static Rectangle Clamp(Rectangle rectangle, Rectangle working)
        {
            return new Rectangle(Math.Max(working.Left, Math.Min(rectangle.Left, working.Right - rectangle.Width)),
                Math.Max(working.Top, Math.Min(rectangle.Top, working.Bottom - rectangle.Height)), rectangle.Width, rectangle.Height);
        }
    }
}
