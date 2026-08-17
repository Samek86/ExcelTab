using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using ExcelTab.ITEM;
using ExcelTab.VIEW;
using Excel = NetOffice.ExcelApi;
using Enums = NetOffice.ExcelApi.Enums;

namespace ExcelTab.CLASSES
{
    public static class FavoriteHelper
    {
        public static void AddFavoriteCell(string key)
        {
            try
            {
                var book = ExcelHelper.GetActiveBook();
                if (book == null)
                {
                    Common.Wink(WinkEnum.Fail);
                    return;
                }

                FavoriteExcel fc = new FavoriteExcel()
                {
                    FullName = book.FullName,
                    SheetName = GetSheetName(book.ActiveSheet),
                    Row = book.Application.ActiveCell.Row,
                    Column = book.Application.ActiveCell.Column,
                    IsCell = true,
                };

                try
                {
                    var img = CaptureHelper.GetWorkbookCapture(book);
                    if (img != null)
                    {
                        img.Save(Path.Combine(Common.TempFolderPath, "Favorite_" + key + ".png"), System.Drawing.Imaging.ImageFormat.Png);
                        img.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    AppLog.Warn("お気に入り画像の保存に失敗しました。", ex);
                }

                App.Setting.FavoriteCellDic[key] = fc;
                App.Setting.Save();
                Common.Wink(WinkEnum.Success);
            }
            catch (Exception ex)
            {
                AppLog.Error("お気に入り登録に失敗しました。", ex);
                Common.Wink(WinkEnum.Fail);
            }
        }

        public static void MoveFavoriteCell(string key)
        {
            MoveFavoriteCell(key, true);
        }

        public static void MoveFavoriteCell(string key, bool allowOpen)
        {
            MainWindow.TopmostFlg = true;
            Common.Invoke(() =>
            {
                ((MainWindow)App.Current.MainWindow).Topmost = true;
            });

            try
            {
                if (!App.Setting.FavoriteCellDic.ContainsKey(key))
                {
                    return;
                }

                var fc = App.Setting.FavoriteCellDic[key];
                if (!File.Exists(fc.FullName))
                {
                    Common.Wink(WinkEnum.Fail);
                    Alert.Error("ファイルが存在しません。");
                    return;
                }

                foreach (var ctrl in MainWindow.TabList)
                {
                    if (ctrl.FullName != fc.FullName)
                    {
                        continue;
                    }

                    ctrl.SetActive();
                    for (int i = 1; i <= ctrl.Book.Sheets.Count; i++)
                    {
                        Excel.Worksheet sheet = AsWorksheet(ctrl.Book.Sheets[i]);
                        if (sheet.Name == fc.SheetName)
                        {
                            sheet.Activate();
                            if (fc.IsCell)
                            {
                                sheet.Cells[fc.Row, fc.Column].Select();
                            }
                            Common.Wink(WinkEnum.Success);
                            break;
                        }
                    }
                    ImageSave(ctrl.Book, key);
                    return;
                }

                if (!allowOpen)
                {
                    Common.Wink(WinkEnum.Fail);
                    return;
                }

                Excel.Application excelApp = ExcelAppHelper.GetActiveApplication();
                try
                {
                    Excel.Workbook book;
                    if (excelApp == null)
                    {
                        Process.Start(fc.FullName);
                        excelApp = WaitForExcelApp();
                        if (excelApp == null)
                        {
                            Common.Wink(WinkEnum.Fail);
                            Alert.Error("Excel を起動できませんでした。");
                            return;
                        }
                        book = excelApp.ActiveWorkbook;
                    }
                    else
                    {
                        book = excelApp.Workbooks.Open(fc.FullName);
                    }

                    var windows = book.Windows;
                    for (int i = 1; i <= windows.Count; i++)
                    {
                        Excel.Window window = windows[i];
                        IntPtr handler = (IntPtr)window.Hwnd;
                        if (window.WindowState == Enums.XlWindowState.xlMinimized)
                        {
                            window.WindowState = Enums.XlWindowState.xlNormal;
                        }
                        window.Activate();
                        Win32Helper.SetForegroundWindow(handler);
                    }

                    if (excelApp.Workbooks.Count > 0)
                    {
                        excelApp.Visible = true;
                    }

                    Thread.Sleep(1000);
                    MoveFavoriteCell(key, false);
                }
                finally
                {
                    ExcelAppHelper.Release(ref excelApp);
                    MainWindow.ReleaseExcelApp();
                }
            }
            catch (Exception ex)
            {
                AppLog.Error("お気に入り移動に失敗しました。", ex);
                Common.Wink(WinkEnum.Fail);
            }
        }

        public static void ImageSave(Excel.Workbook book, string index)
        {
            Task.Run(() =>
            {
                Thread.Sleep(1000);
                Common.Invoke(() =>
                {
                    var img = CaptureHelper.GetWorkbookCapture(book);
                    img?.Save(Path.Combine(Common.TempFolderPath, "Favorite_" + index + ".png"), System.Drawing.Imaging.ImageFormat.Png);
                    img?.Dispose();
                }, priority: DispatcherPriority.Input);
            });
        }

        private static string GetSheetName(object sheet)
        {
            Excel.Worksheet worksheet = AsWorksheet(sheet);
            return worksheet != null ? worksheet.Name : "";
        }

        private static Excel.Worksheet AsWorksheet(object sheet)
        {
            if (sheet is Excel.Worksheet worksheet)
            {
                return worksheet;
            }
            return null;
        }

        private static Excel.Application WaitForExcelApp()
        {
            for (int i = 0; i < 30; i++)
            {
                try
                {
                    var excelApp = ExcelAppHelper.GetActiveApplication();
                    if (excelApp != null && excelApp.Workbooks.Count > 0)
                    {
                        return excelApp;
                    }
                    excelApp?.Dispose();
                }
                catch (Exception)
                {
                }
                Thread.Sleep(1000);
            }
            return null;
        }
    }
}
