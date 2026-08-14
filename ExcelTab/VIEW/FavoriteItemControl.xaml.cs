using ExcelTab.CLASSES;
using ExcelTab.ITEM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelTab.VIEW
{
    /// <summary>
    /// Interaction logic for FavoriteItemControl.xaml
    /// </summary>
    public partial class FavoriteItemControl : UserControl
    {
        public string Index { get; set; } = null;
        public FavoriteExcel Item { get; set; } = null;
        public FavoriteItemControl()
        {
            InitializeComponent();
        }

        private void Grid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if(Index != null)
            {
                Common.SetTimeout(() => FavoriteHelper.MoveFavoriteCell(Index), 200);

                Window parentWindow = Window.GetWindow(this);
                if (parentWindow is FavoriteWindow esw)
                {
                    esw.DialogResult = true;
                }
            }

        }

        private void CellRb_Checked(object sender, RoutedEventArgs e)
        {
            if(Item != null)
            {
                Item.IsCell = true;
                App.Setting.Save();
            }
        }

        private void SheetRb_Checked(object sender, RoutedEventArgs e)
        {
            if(Item != null)
            {
                Item.IsCell = false;
                App.Setting.Save();
            }
        }

        private void DeleteBt_Click(object sender, RoutedEventArgs e)
        {
            App.Setting.FavoriteCellDic.Remove(Index);
            App.Setting.Save();
            Window parentWindow = Window.GetWindow(this);
            if (parentWindow is FavoriteWindow esw)
            {
                esw.RemoveCtrl(Index);
            }
        }
    }
}
