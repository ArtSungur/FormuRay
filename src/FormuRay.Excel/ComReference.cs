using System.Runtime.InteropServices;

namespace FormuRay.Excel
{
    internal static class ComReference
    {
        // Release only references acquired by this service, never the borrowed VSTO Application.
        // Do not FinalReleaseComObject: another component may share the RCW.
        internal static void Release(object value)
        {
            if (value == null || !Marshal.IsComObject(value))
                return;
            try { Marshal.ReleaseComObject(value); }
            catch (InvalidComObjectException) { }
        }
    }
}
