using ExcelTab.CLASSES;
using ExcelTab.VIEW;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;
using static ExcelTab.CLASSES.Win32Helper;
using Excel = NetOffice.ExcelApi;

namespace ExcelTab
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static Excel.Application oExcelApp = null;
        public static int ExcelCount { get; set; } = 0;
        public static int scrollSize = 250;
        public static KeyHookManager _keyHook;
        public static List<TabItemControl> TabList = new List<TabItemControl>();
        public static WindowInteropHelper wndHelper = null;
        public static IntPtr Handle { get; set; } = IntPtr.Zero;

        public static bool TopmostFlg { get; set; } = true;

        private DispatcherTimer _bookTimer;
        private DispatcherTimer _layoutTimer;
        private bool _isClosing;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //TestBt.Visibility = Visibility.Collapsed;

            wndHelper = new WindowInteropHelper(this);
            Handle = wndHelper.Handle;

            System.Drawing.Rectangle area = Common.GetWorkingArea();
            Top = area.Height - (Height * Common.Magnification);
            MaxWidth = area.Width - (300 * Common.Magnification);
            Width = 0;

            MoveBottom();
            GetExcelApp();
            SetTabSp();

            try
            {
                // Alt + Tabで非表示
                int exStyle = (int)GetWindowLong(Handle, (int)GetWindowLongFields.GWL_EXSTYLE);
                exStyle |= (int)ExtendedWindowStyles.WS_EX_TOOLWINDOW;
                SetWindowLong(Handle, (int)GetWindowLongFields.GWL_EXSTYLE, (IntPtr)exStyle);

                // キャプチャ禁止
                SetWindowDisplayAffinity(Handle, WDA_EXCLUDEFROMCAPTURE);

                // Topmost
                SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            }
            catch (Exception)
            {
            }


            string version = Common.GetVersion(Assembly.GetExecutingAssembly());
            VersionMi.Header = $"ExcelTab {version}";
            VersionMi2.Header = $"ExcelTab {version}";
            _keyHook = new KeyHookManager();

            if (CheckStartup())
            {
                MenuStartupBt.Icon = "✔";
                MenuStartupBt2.Icon = "✔";
            }

            Topmost = true;
            StartMonitors();
        }

        public Excel.Application GetExcelApp()
        {
            oExcelApp = ExcelAppHelper.EnsureAlive(oExcelApp);
            if (oExcelApp == null)
            {
                ExcelCount = 0;
            }
            return oExcelApp;
        }

        public void MoveBottom()
        {
            int topPlus = 0;
            if (Common.Magnification == 0.8)
            {
                topPlus = 1;
            }
            else if (Common.Magnification <= 0.7)
            {
                topPlus = 0;
            }

            System.Drawing.Rectangle area = Common.GetWorkingArea();
            MaxWidth = area.Width;
            MaxWidth = area.Width - 300;
            Top = area.Height - Height + area.Top + topPlus;
            Left = (area.Width - ActualWidth) / 2 - (40 * Common.Magnification) + area.Left;
            Common.TooltipWidth = area.Width * 0.6;
        }

        public void SetTabSp()
        {
            ClearTabs();
            try
            {
                //if(oExcelApp == null)
                //{
                //    Visibility = Visibility.Collapsed;
                //    return;
                //}
                Visibility = Visibility.Visible;
                //Excel.Workbooks oWorkBooks = oExcelApp.Workbooks;
                var allWb = ROTManager.GetOpendLocalWorkbooks();

                //if (oWorkBooks.Count == 0)
                if (allWb.Count == 0)
                {
                    Visibility = Visibility.Collapsed;
                    return;
                }

                //for (int i = 1; i <= oWorkBooks.Count; i++)
                for (int i = 0; i < allWb.Count; i++)
                {
                    //Excel.Workbook book = oWorkBooks[i];
                    Excel.Workbook book = allWb[i];
                    TabItemControl item = new TabItemControl();
                    //item.Index = i;
                    item.SetWorkBook(book);
                    item.SetBackgroundColor(i % 5);
                    TabList.Add(item);
                }
                //ExcelCount = oWorkBooks.Count;
                //ExcelCount = allWb.Count;

                if (Common.SortList.Count == 0)
                {
                    for (int i = 0; i < TabList.Count; i++)
                    {
                        TabItemControl tab = TabList[i];
                        tab.Order = i;
                        tab.SetIndex(i + 1);
                        TabSp.Children.Add(tab);
                        Common.SortList.Add(tab.FullName);
                    }
                }
                else
                {
                    List<string> AddedList = new List<string>();
                    foreach (var sl in Common.SortList)
                    {
                        foreach (var tab in TabList)
                        {
                            if(tab.FullName == sl)
                            {
                                TabSp.Children.Add(tab);
                                AddedList.Add(tab.FullName);
                            }
                        }
                    }
                    foreach (var tab in TabList)
                    {
                        if (!AddedList.Contains(tab.FullName))
                        {
                            TabSp.Children.Add(tab);
                            AddedList.Add(tab.FullName);
                        }
                    }
                    Common.SortList.Clear();

                    for (int i = 0; i < TabSp.Children.Count; i++)
                    {
                        TabItemControl tab = (TabItemControl)TabSp.Children[i];
                        tab.Order = i;
                        tab.SetIndex(i + 1);
                        Common.SortList.Add(tab.FullName);
                    }
                }

            }
            catch (Exception e)
            {
                AppLog.Warn("タブ一覧の更新に失敗しました。", e);
            }
            finally
            {
                Dispatcher.BeginInvoke(new Action(UpdateTabBarLayout), DispatcherPriority.Loaded);
            }
        }

        public void ClearTabs()
        {
            try
            {
                TabSp.Children.Clear();
                foreach (var item in TabList)
                {
                    item.ReleaseResources();
                }
                TabList.Clear();
            }
            catch (Exception)
            {
            }
        }

        private void StartMonitors()
        {
            _bookTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _bookTimer.Tick += BookTimer_Tick;
            _bookTimer.Start();

            _layoutTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10) };
            _layoutTimer.Tick += LayoutTimer_Tick;
            _layoutTimer.Start();
        }

        private void StopMonitors()
        {
            if (_bookTimer != null)
            {
                _bookTimer.Stop();
                _bookTimer.Tick -= BookTimer_Tick;
                _bookTimer = null;
            }
            if (_layoutTimer != null)
            {
                _layoutTimer.Stop();
                _layoutTimer.Tick -= LayoutTimer_Tick;
                _layoutTimer = null;
            }
        }

        private void BookTimer_Tick(object sender, EventArgs e)
        {
            if (_isClosing)
            {
                return;
            }
            try
            {
                ExcelCount = TabList.Count;
                int openedCount = ROTManager.GetOpenedLocalWorkbookCount();
                if (openedCount != ExcelCount)
                {
                    SetTabSp();
                    Topmost = false;
                    if (TopmostFlg)
                    {
                        Topmost = true;
                    }
                    SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                }

                bool isExcelActive = false;
                foreach (TabItemControl item in TabList)
                {
                    if (item.ActiveCheck())
                    {
                        isExcelActive = true;
                    }
                }
                Common.IsExcelActive = isExcelActive;
            }
            catch (Exception ex)
            {
                AppLog.Warn("ブック監視に失敗しました。", ex);
                ClearTabs();
            }
        }

        private void LayoutTimer_Tick(object sender, EventArgs e)
        {
            if (_isClosing)
            {
                return;
            }
            try
            {
                GetExcelApp();
                Topmost = false;
                if (TopmostFlg)
                {
                    Topmost = true;
                }
                UpdateTabBarLayout();
            }
            catch (Exception ex)
            {
                AppLog.Warn("レイアウト監視に失敗しました。", ex);
                oExcelApp = null;
            }
        }

        private void UpdateTabBarLayout()
        {
            if (TabList.Count == 0)
            {
                Visibility = Visibility.Collapsed;
                return;
            }

            Visibility = Visibility.Visible;
            double width = 24 + 22;
            foreach (TabItemControl item in TabList)
            {
                width += item.ActualWidth + 5;
            }
            if (MaxWidth > width)
            {
                LeftScrollWidth.Width = new GridLength(0);
                RightScrollWidth.Width = new GridLength(0);
            }
            else
            {
                LeftScrollWidth.Width = new GridLength(20);
                RightScrollWidth.Width = new GridLength(20);
            }
            Width = width;
            MoveBottom();
        }

        public void ActiveClear()
        {
            foreach (TabItemControl item in TabList)
            {
                item.Opacity = 0.5;
            }
        }

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (e.ChangedButton == MouseButton.Left)
                {
                    DragMove();
                }

                //if (e.ClickCount == 2)
                //{
                //    MoveBottom();
                //}
            }
            catch { }

            
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_isClosing)
            {
                return;
            }
            _isClosing = true;
            StopMonitors();
            try
            {
                _keyHook?.Close();
            }
            catch (Exception)
            {
            }
            ClearTabs();
            if (oExcelApp != null)
            {
                ComHelper.ReleaseCom(oExcelApp);
                oExcelApp = null;
            }
            try
            {
                App.Setting.Save();
            }
            catch (Exception)
            {
            }
            System.Windows.Application.Current.Shutdown();
        }

        private void LeftScroll_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                TabSv.ScrollToHorizontalOffset(TabSv.HorizontalOffset - scrollSize);
            }
            catch (Exception)
            {
            }
        }

        private void RightScroll_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                TabSv.ScrollToHorizontalOffset(TabSv.HorizontalOffset + scrollSize);
            }
            catch (Exception)
            {
            }
        }

        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (0 < e.Delta)
            {
                TabSv.ScrollToHorizontalOffset(TabSv.HorizontalOffset - scrollSize);
            }
            else
            {
                TabSv.ScrollToHorizontalOffset(TabSv.HorizontalOffset + scrollSize);
            }
            e.Handled = true;
            return;
        }

        private void SettingBt_Click(object sender, RoutedEventArgs e)
        {

        }

        private void MenuStartupBt_Click(object sender, RoutedEventArgs e)
        {
            if (CheckStartup())
            {
                InstallMeOffStartUp();
                MenuStartupBt.Icon = null;
                MenuStartupBt2.Icon = null;
            }
            else
            {
                InstallMeOnStartUp();
                MenuStartupBt.Icon = "✔";
                MenuStartupBt2.Icon = "✔";
            }
        }

        public void InstallMeOnStartUp()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
                {
                    if (key == null)
                    {
                        return;
                    }
                    Assembly curAssembly = Assembly.GetExecutingAssembly();
                    key.SetValue(curAssembly.GetName().Name, curAssembly.Location);
                }
                Alert.Info("スタートアップに登録しました。");
            }
            catch { }
        }

        public void InstallMeOffStartUp()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
                {
                    if (key == null)
                    {
                        return;
                    }
                    Assembly curAssembly = Assembly.GetExecutingAssembly();
                    key.DeleteValue(curAssembly.GetName().Name, false);
                }
                Alert.Info("スタートアップから削除しました。");
            }
            catch { }
        }

        public bool CheckStartup()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false))
                {
                    if (key == null)
                    {
                        return false;
                    }
                    Assembly curAssembly = Assembly.GetExecutingAssembly();
                    return key.GetValue(curAssembly.GetName().Name) != null;
                }
            }
            catch { }
            return false;
        }

        private void ExitBt_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        public static void SetActive(int index)
        {
            try
            {
                foreach (var tab in TabList)
                {
                    if (tab.Index == index)
                    {
                        tab.SetActive();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        private void UpdateBt_Click(object sender, RoutedEventArgs e)
        {
            UpdateManager.UserUpdate("ExcelTab", Assembly.GetExecutingAssembly());
        }

        private void MoveBottomBt_Click(object sender, RoutedEventArgs e)
        {
            MoveBottom();
        }

        private void TestBt_Click(object sender, RoutedEventArgs e)
        {
            //Common.TestMode = !Common.TestMode;
            //if (Common.TestMode)
            //{
            //    TestBt.Icon = "✔";
            //}
            //else
            //{
            //    TestBt.Icon = null;
            //}
        }

        private void FavoriteWindowOpen_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (App.Setting.FavoriteCellDic.Count == 0)
            {
                Alert.Warn("お気に入りに登録されている情報がありません。\n先に登録してください。\n\n１．登録するセルを選択\n２．Ctrl + Alt + Numpad1 ~ 9");
                return;
            }

            TopmostFlg = false;
            try
            {
                FavoriteWindow fw = new FavoriteWindow();
                fw.ShowDialog();
            }
            catch (Exception)
            {
            }
            TopmostFlg = true;
        }

        private void ExcelKillBt_Click(object sender, RoutedEventArgs e)
        {
            if (Alert.Confirm("すべてのEXCELを強制終了しますか？", "確認") == MessageBoxResult.Yes)
            {
                try
                {
                    Process process = new Process();
                    ProcessStartInfo startInfo = new ProcessStartInfo();
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    startInfo.FileName = "cmd.exe";
                    startInfo.Arguments = "/C taskkill /f /im Excel.exe";
                    process.StartInfo = startInfo;
                    process.Start();
                }
                catch { }
            }
        }
    }
}
