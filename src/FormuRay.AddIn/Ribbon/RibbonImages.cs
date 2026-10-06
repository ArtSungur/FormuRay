using System;
using System.Collections.Generic;
using System.Drawing;

namespace FormuRay.AddIn.Ribbon
{
    internal sealed class RibbonImages : IDisposable
    {
        private readonly Dictionary<string, Image> images = new Dictionary<string, Image>();
        internal Image Get(string name)
        {
            Image result;
            if (!images.TryGetValue(name, out result))
            {
                using (var stream = typeof(RibbonImages).Assembly.GetManifestResourceStream("FormuRay.AddIn.Ribbon.Icons." + name + ".png"))
                {
                    if (stream == null) throw new InvalidOperationException("Missing Ribbon image: " + name);
                    using (var source = Image.FromStream(stream)) result = new Bitmap(source);
                }
                images.Add(name, result);
            }
            return result;
        }
        public void Dispose() { foreach (var image in images.Values) image.Dispose(); images.Clear(); }
    }
}
