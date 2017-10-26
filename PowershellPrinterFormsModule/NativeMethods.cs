using Microsoft.Win32.SafeHandles;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;

namespace PowershellPrinterFormsModule
{
    internal static class NativeMethods
    {
        #region Constants

        internal const int ERROR_INSUFFICIENT_BUFFER = 0x7A;

        #endregion

        #region winspool.drv

        private const string Winspool = "winspool.drv";

        [DllImport(Winspool, SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool OpenPrinter(string szPrinter, out SafePrinterHandle hPrinter, IntPtr pd);

        public static SafePrinterHandle OpenPrinter(string szPrinter)
        {
            SafePrinterHandle hServer;
            if (!OpenPrinter(null, out hServer, IntPtr.Zero))
            {
                throw new Win32Exception();
            }
            return hServer;
        }

        [DllImport(Winspool, SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport(Winspool, SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool EnumForms(SafePrinterHandle hPrinter, int level, IntPtr pBuf, int cbBuf, out int pcbNeeded, out int pcReturned);

        [DllImport(Winspool, SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool AddForm(SafePrinterHandle hPrinter, int level, [In] ref FORM_INFO_1 form);

        [DllImport(Winspool, SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool DeleteForm(SafePrinterHandle hPrinter, string formName);

        #endregion

        #region Structs

        [StructLayout(LayoutKind.Sequential)]
        internal struct FORM_INFO_1
        {
            public int Flags;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string Name;
            public SIZEL Size;
            public RECTL ImageableArea;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct SIZEL
        {
            public int cx;
            public int cy;

            public static explicit operator SIZEL(Size r)
                => new SIZEL { cx = (int)r.Width, cy = (int)r.Height };
            public static explicit operator Size(SIZEL r)
                => new Size(r.cx, r.cy);
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct RECTL
        {
            public int left;
            public int top;
            public int right;
            public int bottom;

            public static explicit operator RECTL(Rect r)
                => new RECTL { left = (int)r.Left, top = (int)r.Top, right = (int)r.Right, bottom = (int)r.Bottom };
            public static explicit operator Rect(RECTL r)
                => new Rect(new Point(r.left, r.top), new Point(r.right, r.bottom));
        }

        #endregion
    }

    internal sealed class SafePrinterHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        public SafePrinterHandle()
            : base(true)
        {
        }

        protected override bool ReleaseHandle()
        {
            return NativeMethods.ClosePrinter(handle);
        }
    }

    internal static class UnitConverter
    {
        public static Point MicronsToMM(Point s) => (Point)((Vector)s * (1d / 1000d));
        public static Size MicronsToMM(Size s) => (Size)((Vector)s * (1d / 1000d));

        public static Point MicronsToInches(Point s) => (Point)((Vector)s * (1d / (1000d * 25.4d)));
        public static Size MicronsToInches(Size s) => (Size)((Vector)s * (1d / (1000d * 25.4d)));

        public static Size InchesToMicrons(Size s) => (Size)((Vector)s * (1000d * 25.4d));
        public static Point InchesToMicrons(Point s) => (Point)((Vector)s * (1000d * 25.4d));

        public static Size MMToMicrons(Size s) => (Size)((Vector)s * 1000d);
        public static Point MMToMicrons(Point s) => (Point)((Vector)s * 1000d);

        public static Thickness MarginFromRect(Size parent, Rect child) => new Thickness(child.Left, child.Top, parent.Width - child.Right, parent.Height - child.Bottom);
    }
}