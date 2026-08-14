using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Excel = NetOffice.ExcelApi;
using Enums = NetOffice.ExcelApi.Enums;

namespace ExcelTab.CLASSES
{
    public static class CaptureHelper
    {
        [DllImport("User32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PrintWindow(IntPtr hwnd, IntPtr hDC, uint nFlags);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr handle, ref Rectangle rect);

        [DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeleteObject([In] IntPtr hObject);

        public static Bitmap CaptureWindow(IntPtr handle)
        {
            Rectangle rect = new Rectangle();
            GetWindowRect(handle, ref rect);
            rect.Width -= rect.X;
            rect.Height -= rect.Y;
            if (rect.Width <= 0 || rect.Height <= 0)
            {
                return null;
            }

            using (Bitmap bitmap = new Bitmap(rect.Width, rect.Height))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    IntPtr hdc = g.GetHdc();
                    if (!PrintWindow(handle, hdc, 0))
                    {
                        int error = Marshal.GetLastWin32Error();
                        AppLog.Warn("PrintWindow に失敗しました。code=" + error);
                    }
                    g.ReleaseHdc(hdc);
                }
                return (Bitmap)bitmap.Clone();
            }
        }

        public static ImageSource ImageSourceFromBitmap(Bitmap bmp)
        {
            if (bmp == null)
            {
                return null;
            }
            var handle = bmp.GetHbitmap();
            try
            {
                return Imaging.CreateBitmapSourceFromHBitmap(handle, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            }
            finally
            {
                DeleteObject(handle);
            }
        }

        public static ImageSource BitmapFromUri(Uri source)
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.UriSource = source;
            bitmap.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.EndInit();
            return bitmap;
        }

        public static Bitmap GetWorkbookCapture(Excel.Workbook book)
        {
            Bitmap rtn = null;
            try
            {
                var windows = book.Windows;
                for (int i = 1; i <= windows.Count; i++)
                {
                    Excel.Window window = windows[i];
                    IntPtr handler = (IntPtr)window.Hwnd;
                    if (window.WindowState == Enums.XlWindowState.xlMinimized)
                    {
                        window.WindowState = Enums.XlWindowState.xlNormal;
                        rtn?.Dispose();
                        rtn = CaptureWindow(handler);
                        window.WindowState = Enums.XlWindowState.xlMinimized;
                    }
                    else
                    {
                        rtn?.Dispose();
                        rtn = CaptureWindow(handler);
                    }
                }
            }
            catch (Exception ex)
            {
                AppLog.Warn("ブックキャプチャに失敗しました。", ex);
            }
            return rtn;
        }
    }
}
