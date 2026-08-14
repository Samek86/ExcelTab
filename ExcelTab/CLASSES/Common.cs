using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using ExcelTab.CLASSES;
using ExcelTab.ITEM;
using static ExcelTab.CLASSES.Win32Helper;

namespace ExcelTab
{
    public static class Common
    {
        public static bool TestMode = false;
        public static string RootPath = AppDomain.CurrentDomain.BaseDirectory;
        public static string TempFolderPath = Path.Combine(RootPath, "TEMP");
        public static bool IsGetKeyData = false;
        public static List<string> SortList = new List<string>();
        public static bool IsIgnoreCtrlAltV = false;
        public static bool IsExcelActive = false;
        public static double TooltipWidth = 900;
        public static int CurrentTabIndex = 1;
        public static double Magnification = 1;

        public static System.Windows.Media.Brush[] BasicTabColor = new System.Windows.Media.Brush[]
        {
            new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 3, 163, 76)),
            new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 0, 67, 124)),
            new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 74, 0, 163)),
            new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 163, 0, 88)),
            new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 163, 74, 0)),
        };

        public static SolidColorBrush NormalBackground = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 45, 45, 48));
        public static SolidColorBrush SuccessBackground = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 9, 241, 32));
        public static SolidColorBrush FailBackground = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 241, 65, 46));

        public static void Invoke(this Action act, FrameworkElement fe = null, bool exceptionThrowFlg = false, DispatcherPriority priority = DispatcherPriority.Normal)
        {
            try
            {
                Dispatcher dis = fe == null ? App.Current.Dispatcher : fe.Dispatcher;
                if (dis == null || dis.CheckAccess())
                {
                    act();
                }
                else
                {
                    dis.Invoke(priority, act);
                }
            }
            catch (Exception ex)
            {
                AppLog.Warn("UI スレッド実行に失敗しました。", ex);
                if (exceptionThrowFlg)
                {
                    throw;
                }
            }
        }

        public static string GetNextText()
        {
            string pasteText = ExcelHelper.PasteText;
            if (pasteText.Contains("-"))
            {
                int hyphenIndex = pasteText.LastIndexOf("-");
                string numText = pasteText.Substring(hyphenIndex + 1);
                if (int.TryParse(numText, out int num))
                {
                    ExcelHelper.PasteText = pasteText.Substring(0, hyphenIndex + 1) + (num + 1);
                }
            }
            return ExcelHelper.PasteText;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        public static bool IsActive(IntPtr handler)
        {
            return GetForegroundWindow() == handler;
        }

        public static void GetScale()
        {
            var screenScale = ScreenManager.GetScale(Application.Current.MainWindow);
            if (screenScale == 1.25d)
            {
                Magnification = 4d / 5d;
            }
            else if (screenScale == 1.50d)
            {
                Magnification = 2d / 3d;
            }
        }

        public static Rectangle ModifyRect(Rectangle r)
        {
            return new Rectangle(
                (int)(r.X * Magnification),
                (int)(r.Y * Magnification),
                (int)(r.Width * Magnification),
                (int)(r.Height * Magnification));
        }

        public static Rectangle GetWorkingArea()
        {
            GetScale();
            WindowInteropHelper windowInteropHelper = new WindowInteropHelper(Application.Current.MainWindow);
            System.Windows.Forms.Screen screen = System.Windows.Forms.Screen.FromHandle(windowInteropHelper.Handle);
            return ModifyRect(screen.WorkingArea);
        }

        public static Rectangle GetBounds()
        {
            GetScale();
            WindowInteropHelper windowInteropHelper = new WindowInteropHelper(Application.Current.MainWindow);
            System.Windows.Forms.Screen screen = System.Windows.Forms.Screen.FromHandle(windowInteropHelper.Handle);
            return ModifyRect(screen.Bounds);
        }

        public static bool CheckBright(System.Windows.Media.Color c)
        {
            return c.R * 0.2126 + c.G * 0.7152 + c.B * 0.0722 <= 255 / 2 + 1;
        }

        public static bool CheckBright(System.Windows.Media.Brush c)
        {
            if (c is SolidColorBrush scb)
            {
                return CheckBright(scb.Color);
            }
            return false;
        }

        public static System.Windows.Media.Color ToColor(System.Windows.Media.Brush brush)
        {
            if (brush is SolidColorBrush solid)
            {
                return solid.Color;
            }
            return Colors.Black;
        }

        public static SolidColorBrush ToBrush(System.Windows.Media.Color color)
        {
            return new SolidColorBrush(color);
        }

        public static string GetVersion(Assembly asb, string prefix = "v", int step = 4)
        {
            var thisVer = asb.GetName().Version;
            switch (step)
            {
                case 1:
                    return prefix + thisVer.Major;
                case 2:
                    return prefix + thisVer.Major + "." + thisVer.Minor;
                case 3:
                    return prefix + thisVer.Major + "." + thisVer.Minor + "." + thisVer.Build;
                default:
                    return prefix + thisVer;
            }
        }

        public static void SetTimeout(Action act, int milliseconds, bool invokeFlg = false)
        {
            Task.Run(() =>
            {
                for (int i = 0; i < 20; i++)
                {
                    try
                    {
                        Thread.Sleep(milliseconds);
                        if (invokeFlg)
                        {
                            App.Current.Dispatcher.Invoke(() => act?.Invoke());
                        }
                        else
                        {
                            act?.Invoke();
                        }
                    }
                    catch (Exception ex)
                    {
                        AppLog.Warn("SetTimeout 実行に失敗しました。", ex);
                        milliseconds += 100;
                        continue;
                    }
                    break;
                }
            });
        }

        public static void Wink(WinkEnum winkEnum)
        {
            Task.Run(() =>
            {
                try
                {
                    switch (winkEnum)
                    {
                        case WinkEnum.Success:
                            App.Current.Dispatcher.Invoke(() => ((MainWindow)App.Current.MainWindow).MainBd.Background = SuccessBackground);
                            Thread.Sleep(500);
                            App.Current.Dispatcher.Invoke(() => ((MainWindow)App.Current.MainWindow).MainBd.Background = NormalBackground);
                            break;
                        case WinkEnum.Fail:
                            App.Current.Dispatcher.Invoke(() => ((MainWindow)App.Current.MainWindow).MainBd.Background = FailBackground);
                            Thread.Sleep(500);
                            App.Current.Dispatcher.Invoke(() => ((MainWindow)App.Current.MainWindow).MainBd.Background = NormalBackground);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    AppLog.Warn("Wink 表示に失敗しました。", ex);
                }
            });
        }

        public static void Try(this Action act)
        {
            try
            {
                act();
            }
            catch (Exception ex)
            {
                AppLog.Warn("例外を無視しました。", ex);
            }
        }

        public static void CreateDirectory(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        public static void CreateFile(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }
            if (!File.Exists(path))
            {
                File.WriteAllText(path, "");
            }
        }

        public static void Kill(params string[] processNames)
        {
            for (int i = 0; i < processNames.Length; i++)
            {
                try
                {
                    foreach (var proc in Process.GetProcessesByName(processNames[i]))
                    {
                        proc.Kill();
                    }
                }
                catch (Exception ex)
                {
                    AppLog.Warn("プロセス終了に失敗しました: " + processNames[i], ex);
                }
            }
        }
    }
}
