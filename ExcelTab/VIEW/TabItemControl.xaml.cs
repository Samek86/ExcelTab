using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ExcelTab;
using ExcelTab.CLASSES;
using WPF.ColorPicker;
using Enums = NetOffice.ExcelApi.Enums;
using Excel = NetOffice.ExcelApi;

namespace ExcelTab.VIEW
{
    /// <summary>
    /// Interaction logic for TabItemControl.xaml
    /// </summary>
    public partial class TabItemControl : System.Windows.Controls.UserControl
    {
        public Excel.Workbook Book { get; set; }
        public string FullName { get; set; }
        public string FileName { get; set; }
        public string ViewName { get; set; }

        public System.Drawing.Bitmap LastImage { get; set; } = default;
        public Brush Color { get; set; }

        public bool IsClick { get; set; } = false;
        public int Index { get; set; } = -1;
        public int Order { get; set; } = -1;

        private Point _pressPoint;
        private bool _isPressed;
        private bool _dragStarted;
        private const double DragThreshold = 8;
        public TabItemControl()
        {
            InitializeComponent();
        }

        public void SetWorkBook(Excel.Workbook book)
        {
            Book = book;
            SetName(book.FullName, book.Name);
        }

        public void SetBackgroundColor(Brush b, bool saveFlg = false)
        {
            try
            {
                Color = b;
                TabBd.Background = Color;
                ToolTipSp.Background = Color;
                FileNameTb.Foreground = Common.CheckBright(Color) ? Brushes.White : Brushes.Black;
                if (Common.CheckBright(Color))
                {
                    FileNameTb.Foreground = Brushes.White;
                    ToolTipName.Foreground = Brushes.White;
                }
                else
                {
                    FileNameTb.Foreground = Brushes.Black;
                    ToolTipName.Foreground = Brushes.Black;
                }
                if (saveFlg)
                {
                    if (App.Setting.TabColorDic.ContainsKey(Book.FullName))
                    {
                        App.Setting.TabColorDic[Book.FullName] = Common.ToColor(Color);
                    }
                    else
                    {
                        App.Setting.TabColorDic.Add(Book.FullName, Common.ToColor(Color));
                    }
                    App.Setting.Save();
                }
            }
            catch (Exception)
            {
            }
        }
        public void SetBackgroundColor(int index)
        {
            try
            {
                if (App.Setting.TabColorDic.ContainsKey(Book.FullName))
                {
                    SetBackgroundColor(App.Setting.TabColorDic[Book.FullName]);
                }
                else
                {
                    SetBackgroundColor(Common.BasicTabColor[index]);
                }
            }
            catch (Exception)
            {
            }
        }
        public void SetBackgroundColor(Color c)
        {
            SetBackgroundColor(new SolidColorBrush(c), true);
        }

        private void SetName(string fullName, string fileName)
        {
            FullName = fullName;
            FileName = fileName;
            ToolTipName.Text = fileName;
            IndexTb.Text = Index.ToString();
            if (App.Setting.TabNameDic.ContainsKey(Book.FullName))
            {
                ViewName = App.Setting.TabNameDic[Book.FullName];
                FileNameTb.Text = ViewName;
            }
            else
            {
                ViewName = fileName;
                if (ViewName.Contains(".xls"))
                {
                    ViewName = ViewName.Substring(0, ViewName.IndexOf(".xls"));
                }
                FileNameTb.Text = ViewName;
            }
        }

        public void SetIndex(int index)
        {
            Index = index;
            IndexTb.Text = Index.ToString();
        }

        public void ReleaseResources()
        {
            try
            {
                ToolTipImg.Source = null;
                LastImage?.Dispose();
                LastImage = null;
            }
            catch (Exception)
            {
            }
            ComHelper.ReleaseCom(Book);
            Book = null;
        }

        private void ReplaceLastImage(System.Drawing.Bitmap next)
        {
            var prev = LastImage;
            LastImage = next;
            if (prev != null && !ReferenceEquals(prev, next))
            {
                try
                {
                    prev.Dispose();
                }
                catch (Exception)
                {
                }
            }
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            ToolTipImg.Width = Common.TooltipWidth;
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        public bool ActiveCheck()
        {
            try
            {
                if(Book == null)
                {
                    return false;
                }
                var windows = Book.Windows;
                for (int i = 1; i <= windows.Count; i++)
                {
                    Excel.Window window = windows[i];
                    IntPtr handler = (IntPtr)window.Hwnd;

                    if (Common.IsActive(handler))
                    {
                        Opacity = 1;
                        Common.CurrentTabIndex = Index;
                        return true;
                    }
                    else
                    {
                        Opacity = 0.5;
                    }
                }
                return false;
            }
            catch (Exception)
            {
            }
            return false;
        }

        private void Border_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            bool wasDrag = _dragStarted;
            _isPressed = false;
            _dragStarted = false;
            if (IsMouseCaptured)
            {
                ReleaseMouseCapture();
            }
            if (!wasDrag)
            {
                SetActive();
            }
        }


        [MethodImpl(MethodImplOptions.Synchronized)]
        public void SetActive(System.Action afterAct = null)
        {
            Task.Run(() =>
            {
                Dispatcher.Invoke(() =>
                {
                    ToolTipService.SetIsEnabled(TabBd, false);
                    try
                    {
                        ((MainWindow)App.Current.MainWindow).ActiveClear();
                        Opacity = 1;
                        var windows = Book.Windows;
                        for (int i = 1; i <= windows.Count; i++)
                        {
                            Excel.Window window = windows[i];
                            //if (Convert.ToString(window.Type) == "xlWorkbook")
                            {
                                IntPtr handler = (IntPtr)window.Hwnd;
                                if (window.WindowState == Enums.XlWindowState.xlMinimized)
                                {
                                    window.WindowState = Enums.XlWindowState.xlNormal;
                                }
                                window.Activate();
                                SetForegroundWindow(handler);
                            }
                        }
                        Common.CurrentTabIndex = Index;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                    afterAct?.Invoke();
                });
            });
        }

        private void TabBd_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            ToolTipService.SetIsEnabled(TabBd, false);
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        private void UserControl_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            try
            {
                ToolTipService.SetIsEnabled(TabBd, true);
                ToolTipService.SetHorizontalOffset(TabBd, (Common.TooltipWidth / 2 - ActualWidth / 2) * -1);

                //var img = GetCaptureImg();
                //string path = System.IO.Path.Combine(Common.TempFolderPath, $"{Index}.png");
                //img.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                //ToolTipImg.Source = null;
                //ToolTipImg.Source = Common.BitmapFromUri(new Uri(path));
                //img?.Dispose();

                ToolTipImg.Source = null;
                if (LastImage != null && LastImage.Height > 50)
                {
                    ToolTipImg.Source = CaptureHelper.ImageSourceFromBitmap(LastImage);
                }
                ToolTipImg.Source = CaptureHelper.ImageSourceFromBitmap(GetCaptureImg());

            }
            catch (Exception)
            {
            }
        }

        public System.Drawing.Bitmap GetCaptureImg()
        {
            try
            {
                var windows = Book.Windows;
                for (int i = 1; i <= windows.Count; i++)
                {
                    Excel.Window window = windows[i];
                    //if (Convert.ToString(window.Type) == "xlWorkbook")
                    {
                        IntPtr handler = (IntPtr)window.Hwnd;

                        if (window.WindowState == Enums.XlWindowState.xlMinimized)
                        {
                            if (LastImage == null || LastImage.Height <= 50)
                            {
                                window.WindowState = Enums.XlWindowState.xlNormal;
                                ReplaceLastImage(CaptureHelper.CaptureWindow(handler));
                                window.WindowState = Enums.XlWindowState.xlMinimized;
                            }
                        }
                        else
                        {
                            ReplaceLastImage(CaptureHelper.CaptureWindow(handler));
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            return LastImage;
        }

        private void ExcelCloseBt_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SetActive(() => {
                    var book = Book;
                    Book = null;
                    if (book != null && !book.IsDisposed)
                    {
                        book.Close();
                        book.Dispose();
                    }
                    MainWindow.ReleaseExcelApp();
                });
            }
            catch (Exception)
            {
            }
        }

        private void SelectColorBt_Click(object sender, RoutedEventArgs e)
        {
            if (ColorPickerWindow.ShowDialog(out Color c))
            {
                SetBackgroundColor(c);
            }
        }

        private void ViewNameChangeBt_Click(object sender, RoutedEventArgs e)
        {
            string viewName = PromptDialog.Prompt("変更する名前を入力してください。", "表示名変更", ViewName);
            if (string.IsNullOrEmpty(viewName))
            {
                return;
            }

            if (App.Setting.TabNameDic.ContainsKey(Book.FullName))
            {
                App.Setting.TabNameDic[Book.FullName] = viewName;
            }
            else
            {
                App.Setting.TabNameDic.Add(Book.FullName, viewName);
            }
            App.Setting.Save();
            ViewName = viewName;
            FileNameTb.Text = ViewName;
            ((MainWindow)App.Current.MainWindow).SetTabSp();
        }

        private void DoDragDrop(DependencyObject dragSource, object data, bool flg = true)
        {
            try
            {
                if (data != null)
                {
                    DragDrop.DoDragDrop(dragSource, data, flg ? System.Windows.DragDropEffects.Copy : DragDropEffects.None);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        private void UserControl_Drop(object sender, System.Windows.DragEventArgs e)
        {
            try
            {
                if (e.Data.GetData(typeof(TabItemControl)) is TabItemControl ctrl)
                {
                    if (Order > ctrl.Order)
                    {
                        Common.SortList.Remove(ctrl.FullName);
                        Common.SortList.Insert(Order, ctrl.FullName);
                    }
                    else if (Order < ctrl.Order)
                    {
                        Common.SortList.Remove(ctrl.FullName);
                        Common.SortList.Insert(Order, ctrl.FullName);
                    }
                    else
                    {
                        //DoDragDrop(this, this, false);
                        //Console.WriteLine("NotDrop");
                        SetActive();
                        return;
                    }
                    ((MainWindow)App.Current.MainWindow).SetTabSp();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        private void UserControl_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left)
            {
                return;
            }
            _isPressed = true;
            _dragStarted = false;
            _pressPoint = e.GetPosition(this);
            CaptureMouse();
        }

        private void UserControl_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isPressed || _dragStarted || e.LeftButton != MouseButtonState.Pressed)
            {
                return;
            }
            Point current = e.GetPosition(this);
            if (Math.Abs(current.X - _pressPoint.X) <= DragThreshold && Math.Abs(current.Y - _pressPoint.Y) <= DragThreshold)
            {
                return;
            }
            _dragStarted = true;
            try
            {
                DoDragDrop(this, this);
            }
            finally
            {
                _isPressed = false;
                if (IsMouseCaptured)
                {
                    ReleaseMouseCapture();
                }
            }
        }

        private void UserControl_DragEnter(object sender, System.Windows.DragEventArgs e)
        {
            try
            {
                if (e.Effects.HasFlag(DragDropEffects.Copy) && e.Data.GetData(typeof(TabItemControl)) is TabItemControl ctrl)
                {
                    if (Order > ctrl.Order)
                    {
                        DropIcon.Kind = MahApps.Metro.IconPacks.PackIconBoxIconsKind.SolidArrowFromLeft;
                        DropBd.Visibility = Visibility.Visible;
                    }
                    else if (Order < ctrl.Order)
                    {
                        DropIcon.Kind = MahApps.Metro.IconPacks.PackIconBoxIconsKind.SolidArrowFromRight;
                        DropBd.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        private void UserControl_DragLeave(object sender, System.Windows.DragEventArgs e)
        {
            DropBd.Visibility = Visibility.Collapsed;
        }

        private void OpenFolderBt_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string dir = System.IO.Path.GetDirectoryName(FullName);
                Process.Start(dir);
            }
            catch (Exception)
            {
            }
        }

        private void UserControl_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            //Console.WriteLine("MouseUp");
            //DoDragDrop(this, this, false);
        }

        private void UserControl_MouseLeave(object sender, MouseEventArgs e)
        {
            DropBd.Visibility = Visibility.Collapsed;
        }
    }
}
