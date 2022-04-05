using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointSlideHtmlLayoutDemo
{
    public class PowerPointHelper
    {
        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out Object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        public static bool TryGetRunningApplication(out Microsoft.Office.Interop.PowerPoint.Application powerPoint)
        {
            CLSIDFromProgID("PowerPoint.Application", out Guid clsid);
            var result = GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
            powerPoint = obj as Microsoft.Office.Interop.PowerPoint.Application;
            return result == 0;
        }

        public static bool TryGetOpenPresentation(Microsoft.Office.Interop.PowerPoint.Presentations presentations,string pptxFilePath, out Microsoft.Office.Interop.PowerPoint.Presentation openPresentation)
        {
            foreach (Microsoft.Office.Interop.PowerPoint.Presentation presentation in presentations)
            {
                if (String.Equals(presentation.FullName, pptxFilePath, StringComparison.OrdinalIgnoreCase))
                {
                    openPresentation = presentation;
                    return true;
                }
            }
            openPresentation = null;
            return false;
        }
    }
}
