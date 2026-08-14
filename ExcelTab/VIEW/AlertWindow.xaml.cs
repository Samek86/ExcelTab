using ExcelTab.CLASSES;
using System;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace ExcelTab.VIEW
{
    public enum AlertStateEnum
    {
        INFO = 0, //情報
        OK = 1, //正常
        WARN = 2, //注意/警告
        ERROR = 4,//エラー
        EXCEPTION = 8, //エラー
    }


    /// <summary>
    /// Interaction logic for AlertWindow.xaml
    /// </summary>
    public partial class Alert : Window
    {
        //info : FFC1C1C1
        //ok : FF00E434
        //warn : FFE4E400
        //error : FFE41500

        private readonly SolidColorBrush InfoColor = new BrushConverter().ConvertFrom("#FFC1C1C1") as SolidColorBrush;
        private readonly SolidColorBrush OkColor = new BrushConverter().ConvertFrom("#FF00E434") as SolidColorBrush;
        private readonly SolidColorBrush WarnColor = new BrushConverter().ConvertFrom("#FFE4E400") as SolidColorBrush;
        private readonly SolidColorBrush ErrorColor = new BrushConverter().ConvertFrom("#FFE41500") as SolidColorBrush;

        MessageBoxResult result = MessageBoxResult.None;
        readonly MessageBoxButton buttons = MessageBoxButton.OK;
        public bool IsShortkey { get; set; } = true;
        public Alert(string messageBoxText, string caption, MessageBoxButton buttons, AlertStateEnum state, int autoCloseTime = -1)
        {
            InitializeComponent();
            //Log.Info($"[Alert]({state}) {messageBoxText}");
            TitleTb.Text = caption;
            MessageTb.Text = messageBoxText;
            this.buttons = buttons;
            this.Opacity = 0;
            switch (buttons)
            {
                case MessageBoxButton.OK:
                    CancelBt.Visibility = Visibility.Collapsed;
                    OkBt.Content = "OK";
                    break;
                case MessageBoxButton.OKCancel:
                    CancelBt.Content = "Cancel";
                    OkBt.Content = "OK";
                    break;
                case MessageBoxButton.YesNoCancel:
                case MessageBoxButton.YesNo:
                    CancelBt.Content = "No";
                    OkBt.Content = "Yes";
                    break;
                default:
                    break;
            }
            switch (state)
            {
                case AlertStateEnum.INFO:
                    StateBd.Background = InfoColor;
                    break;
                case AlertStateEnum.OK:
                    StateBd.Background = OkColor;
                    break;
                case AlertStateEnum.WARN:
                    StateBd.Background = WarnColor;
                    break;
                case AlertStateEnum.ERROR:
                    StateBd.Background = ErrorColor;
                    break;
                case AlertStateEnum.EXCEPTION:
                    StateBd.Background = ErrorColor;
                    break;
                default:
                    break;
            }

            //if (autoCloseTime > 0)
            //{
            //    int time = autoCloseTime;
            //    AutoCloseTimeTb.Text = time.ToString();
            //    Task.Factory.StartNew(() =>
            //    {
            //        try
            //        {
            //            while (true)
            //            {
            //                Thread.Sleep(1000);
            //                --time;
            //                if (result == MessageBoxResult.None)
            //                {
            //                    if (time > 0)
            //                    {
            //                        ViewCommon.Invoke(() => { AutoCloseTimeTb.Text = time.ToString(); });
            //                    }
            //                    else
            //                    {
            //                        ViewCommon.SetOverlayBd(false);
            //                        ViewCommon.Invoke(() => Close());
            //                        break;
            //                    }
            //                }
            //                else
            //                {
            //                    break;
            //                }
            //            }
            //        }
            //        catch (Exception)
            //        {
            //        }
            //    });
            //}
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //ViewCommon.SetOverlayBd(true);
            var msgList = MessageTb.Text.Split('\n');
            double maxLength = 0;
            foreach (string m in msgList)
            {
                double halfCount = numFullWidth(m) * 2d + numHalfWidth(m) * 1.1d;
                if (halfCount > maxLength)
                {
                    maxLength = halfCount;
                }
            }
            Width = (double)maxLength * 7.65d + 70d;
            Height = MainPanel.ActualHeight + 62;

            void SetPosition(Rect r)
            {
                Left = r.Left + ((r.Width - ActualWidth) / 2);
                Top = r.Top + ((r.Height - ActualHeight) / 2);
            }

            // 各画面の真ん中に表示する
            //Window win;
            //win = Application.Current.Windows.OfType<Window>().SingleOrDefault(x => x.ToString().Contains("ObjectDetailWindow"));
            //if (win != null && win.WindowState == WindowState.Normal)
            //{
            //    SetPosition(win.RestoreBounds);
            //    goto PositionSettingEnd;
            //}
            //win = Application.Current.Windows.OfType<Window>().SingleOrDefault(x => x.ToString().Contains("ObjectWindow"));
            //if (win != null && win.WindowState == WindowState.Normal)
            //{
            //    SetPosition(win.RestoreBounds);
            //    goto PositionSettingEnd;
            //}
            //win = Application.Current.Windows.OfType<Window>().SingleOrDefault(x => x.ToString().Contains("SimpleWindow"));
            //if (win != null && win.WindowState == WindowState.Normal)
            //{
            //    SetPosition(win.RestoreBounds);
            //    goto PositionSettingEnd;
            //}

            Rect ScreenRect = ScreenManager.GetScreenFrom(this).WorkingArea;
            SetPosition(ScreenRect);

            //PositionSettingEnd:;

            this.Opacity = 1;
            MessageTb.Focus();

            int numFullWidth(string chkStr)
            {
                int chrByteNum = shiftjisEnc.GetByteCount(chkStr);
                return chrByteNum - chkStr.Length;
            }

            int numHalfWidth(string chkStr)
            {
                int chrByteNum = shiftjisEnc.GetByteCount(chkStr);
                return chkStr.Length * 2 - chrByteNum;
            }

        }

        readonly Encoding shiftjisEnc = Encoding.GetEncoding("Shift_JIS");

        public static MessageBoxResult Info(
            object messageBoxText,
            string caption = "お知らせ"
            )
        {
            return Show(messageBoxText, caption, MessageBoxButton.OK, AlertStateEnum.INFO);
        }

        public static MessageBoxResult Confirm(
            object messageBoxText,
            string caption = "確認"
            )
        {
            return Show(messageBoxText, caption, MessageBoxButton.YesNo, AlertStateEnum.OK);
        }

        public static MessageBoxResult OK(
            object messageBoxText,
            string caption = "お知らせ"
            )
        {
            return Show(messageBoxText, caption, MessageBoxButton.OK, AlertStateEnum.OK);
        }

        public static MessageBoxResult Warn(
            object messageBoxText,
            string caption = "警告"
            )
        {
            return Show(messageBoxText, caption, MessageBoxButton.OK, AlertStateEnum.WARN);
        }

        public static MessageBoxResult Error(
            object messageBoxText,
            string caption = "エラー"
            )
        {
            return Show(messageBoxText, caption, MessageBoxButton.OK, AlertStateEnum.ERROR);
        }

        public static MessageBoxResult Exception(
            object messageBoxText,
            string caption = "Exception"
            )
        {
            return Show(messageBoxText, caption, MessageBoxButton.OK, AlertStateEnum.EXCEPTION);
        }

        public static MessageBoxResult AutoClose(
            object messageBoxText,
            string caption = "お知らせ",
            int time = 5
            )
        {
            return Show(messageBoxText, caption, MessageBoxButton.OK, AlertStateEnum.INFO, time);
        }

        public static MessageBoxResult Show(
            object messageBoxText,
            string caption = "お知らせ",
            MessageBoxButton button = MessageBoxButton.OK,
            AlertStateEnum state = AlertStateEnum.INFO,
            int autoCloseTime = -1,
            bool isShortkey = true
            )
        {
            MessageBoxResult result = MessageBoxResult.None;
            Common.Invoke(() =>
            {
                Alert at = new Alert(messageBoxText.ToString(), caption, button, state, autoCloseTime);
                at.IsShortkey = isShortkey;
                if (at.ShowDialog() == true)
                {
                    result = at.result;
                }
                else
                {
                    result = at.result;
                    //result = MessageBoxResult.Cancel;
                }
            });
            return result;
        }

        private void OKButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            //ViewCommon.SetOverlayBd(false);
            DialogResult = true;
            switch (buttons)
            {
                case MessageBoxButton.YesNoCancel:
                case MessageBoxButton.YesNo:
                    result = MessageBoxResult.Yes;
                    break;
                default:
                    result = MessageBoxResult.OK;
                    break;
            }
        }

        private void CancelButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            //ViewCommon.SetOverlayBd(false);
            DialogResult = false;
            switch (buttons)
            {
                case MessageBoxButton.YesNoCancel:
                case MessageBoxButton.YesNo:
                    result = MessageBoxResult.No;
                    break;
                default:
                    result = MessageBoxResult.Cancel;
                    break;
            }
        }

        private void Close_MouseDown(object sender, MouseButtonEventArgs e)
        {
            //ViewCommon.SetOverlayBd(false);
            DialogResult = false;
            result = MessageBoxResult.Cancel;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                try
                {
                    this.DragMove();
                }
                catch (Exception)
                {
                }
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if (IsShortkey)
            {
                if (e.Key == Key.Enter)
                {
                    OKButton_Click(null, null);
                    return;
                }
                if (e.Key == Key.Escape)
                {
                    Close_MouseDown(null, null);
                    return;
                }
            }
        }

        private void StateBd_MouseDown(object sender, MouseButtonEventArgs e)
        {
            //Log.Info(TestTb.ActualWidth.ToString());
        }
    }
}
