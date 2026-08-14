using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media;

namespace ExcelTab.CLASSES
{
    public class ScreenManager
    {
        public static IEnumerable<ScreenManager> AllScreens()
        {
            foreach (Screen screen in System.Windows.Forms.Screen.AllScreens)
            {
                yield return new ScreenManager(screen);
            }
        }

        public static ScreenManager GetScreenFrom(Window window)
        {
            WindowInteropHelper windowInteropHelper = new WindowInteropHelper(window);
            Screen screen = System.Windows.Forms.Screen.FromHandle(windowInteropHelper.Handle);
            ScreenManager screenManager = new ScreenManager(screen, window);
            return screenManager;
        }

        public static ScreenManager GetScreenFrom(System.Drawing.Point point)
        {
            //int x = (int)Math.Round(point.X);
            //int y = (int)Math.Round(point.Y);
            int x = point.X;
            int y = point.Y;

            // are x,y device-independent-pixels ??
            System.Drawing.Point drawingPoint = new System.Drawing.Point(x, y);
            Screen screen = System.Windows.Forms.Screen.FromPoint(drawingPoint);
            ScreenManager screenManager = new ScreenManager(screen);

            return screenManager;
        }

        public static ScreenManager Primary
        {
            get { return new ScreenManager(System.Windows.Forms.Screen.PrimaryScreen); }
        }

        private readonly Screen screen;

        private readonly Window win;

        private readonly double scale;

        public ScreenManager(System.Windows.Forms.Screen screen, Window win = null)
        {
            this.screen = screen;
            this.win = win;
            this.scale = GetScale(win);
        }

        public Rect DeviceBounds
        {
            get { return this.GetRect(this.screen.Bounds); }
        }

        public Rect WorkingArea
        {
            get { return this.GetRect(this.screen.WorkingArea); }
        }

        private Rect GetRect(Rectangle value)
        {
            // should x, y, width, height be device-independent-pixels ??
            return new Rect
            {
                X = value.X,
                Y = value.Y,
                Width = value.Width,
                Height = value.Height
            };
        }

        public bool IsPrimary
        {
            get { return this.screen.Primary; }
        }

        public string DeviceName
        {
            get { return this.screen.DeviceName; }
        }

        public double GetScale()
        {
            double dpiX = 96.0;
            try
            {
                PresentationSource source = PresentationSource.FromVisual(win);
                if (source != null)
                {
                    dpiX = 96.0 * source.CompositionTarget.TransformToDevice.M11;
                }
            }
            catch (Exception ex)
            {
                //Log.Exception(ex);
            }
            return dpiX;
        }

        public static double GetScale(Visual w = null)
        {
            if (w == null)
            {
                return 1;
            }
            var dpi = VisualTreeHelper.GetDpi(w).PixelsPerDip;
            return dpi;
        }
    }
}
