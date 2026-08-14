using System;
using System.IO;
using System.Windows;
using System.Windows.Input;
using ExcelTab.CLASSES;

namespace ExcelTab.VIEW
{
    /// <summary>
    /// Interaction logic for FavoriteWindow.xaml
    /// </summary>
    public partial class FavoriteWindow : Window
    {
        public FavoriteWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Normal;
            System.Drawing.Rectangle bounds = Common.GetBounds();
            Top = bounds.Top;
            Left = bounds.Left;
            Width = bounds.Width;
            Height = bounds.Height;

            for (int i = 1; i <= 9; i++)
            {
                string index = i.ToString();
                var item = FindName($"Ctrl{index}") as FavoriteItemControl;
                if (App.Setting.FavoriteCellDic.ContainsKey(index))
                {
                    var fe = App.Setting.FavoriteCellDic[index];
                    item.IndexTb.Text = index;
                    item.FileNameTb.Text = System.IO.Path.GetFileName(fe.FullName);
                    item.FileNameTb.ToolTip = fe.FullName;
                    item.SheetNameTb.Text = fe.SheetName;
                    item.SheetNameTb.ToolTip = fe.SheetName;
                    item.CellTb.Text = $"R:{fe.Row}, C:{fe.Column}";
                    string imgPath = System.IO.Path.Combine(Common.TempFolderPath, $"Favorite_{index}.png");
                    if (File.Exists(imgPath))
                    {
                        item.PreviewImg.Source = CaptureHelper.BitmapFromUri(new Uri(imgPath));
                    }
                    else
                    {
                        // TODO
                    }
                    //item.PreviewImg.Source = Common.GetSettingImage($"Favorite_{index}.png");

                    if (fe.IsCell)
                    {
                        item.CellRb.IsChecked = true;
                    }
                    else
                    {
                        item.SheetRb.IsChecked = true;
                    }
                    item.Item = fe;
                    item.Index = index;
                    item.Visibility = Visibility.Visible;
                }
                else
                {
                    item.Visibility = Visibility.Collapsed;
                }
            }
            Topmost = false;
            Topmost = true;

            Activate();
        }

        public void RemoveCtrl(string index)
        {
            var item = FindName($"Ctrl{index}") as FavoriteItemControl;
            item.Visibility = Visibility.Collapsed;
        }
        private void Window_MouseUp(object sender, MouseButtonEventArgs e)
        {
            DialogResult = false;
            MainWindow.TopmostFlg = true;
            ((MainWindow)App.Current.MainWindow).Topmost = true;
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Escape:
                    DialogResult = false;
                    MainWindow.TopmostFlg = true;
                    ((MainWindow)App.Current.MainWindow).Topmost = true;
                    break;
                default:
                    if ((e.Key >= Key.D1 && e.Key <= Key.D9) || (e.Key >= Key.NumPad1 && e.Key <= Key.NumPad9))
                    {
                        int index = e.Key >= Key.NumPad1 ? (int)(e.Key - Key.NumPad0) : (int)(e.Key - Key.D0);
                        Common.SetTimeout(() => FavoriteHelper.MoveFavoriteCell(index.ToString()), 200);
                        DialogResult = true;
                    }
                    return;
            }
        }
    }
}
