using ExcelTab.Extension;
using ExcelTab.VIEW;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTab.CLASSES
{
    public class KeyHookManager
    {
        public static KeyboardHook _hook;

        public KeyHookManager()
        {
            _hook = new KeyboardHook
            {
                OnKeyDown = OnHookKeyDown,
                OnKeyUp = OnHookKeyUp
            };
            _hook.Init();
        }

        public void Close()
        {
            _hook.Close();
        }

        public static bool IsLeftDown = false;
        public static string CurrentKeyDown = "";

        /// <summary>
        /// ショートカットキー「+」の文字列に変換
        /// </summary>
        /// <param name="keys"></param>
        /// <returns></returns>
        public static string GetKeyToString(string keys)
        {
            try
            {
                string[] keyArray = keys.Split(',').Trim();
                SortedDictionary<int, string> resultKeyList = new SortedDictionary<int, string>();
                foreach (string key in keyArray)
                {
                    switch (key)
                    {
                        case "LControlKey":
                        case "RControlKey":
                        case "Control":
                            if (!resultKeyList.ContainsKey(0))
                            {
                                resultKeyList.Add(0, "Ctrl");
                            }
                            break;
                        case "LMenu":
                        case "RMenu":
                        case "Alt":
                            if (!resultKeyList.ContainsKey(1))
                            {
                                resultKeyList.Add(1, "Alt");
                            }
                            break;
                        case "LShiftKey":
                        case "RShiftKey":
                        case "Shift":
                            if (!resultKeyList.ContainsKey(2))
                            {
                                resultKeyList.Add(2, "Shift");
                            }
                            break;
                        case "Tab":
                            if (!resultKeyList.ContainsKey(3))
                            {
                                resultKeyList.Add(3, "Tab");
                            }
                            break;
                        default:
                            if (!resultKeyList.ContainsKey(4))
                            {
                                resultKeyList.Add(4, key);
                            }
                            break;
                    }
                }
                return string.Join(" + ", resultKeyList.Values);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "";
            }
            
        }

        public void OnHookKeyDown(object sender, KeyEventArgs e)
        {
            //Log.Info(e.Key.ToString());
            //Debug.WriteLine(e.KeyCode.ToString());
            //Debug.WriteLine(e.KeyData.ToString());
            try
            {
                string keyData = GetKeyToString(e.KeyData.ToString());

                //Debug.WriteLine(keyData);

                if (string.IsNullOrEmpty(keyData) || keyData == "Ctrl + Shift" || keyData == "Ctrl" || keyData == "Shift" || keyData == "D2")
                {
                    return;
                }

                if (Common.IsGetKeyData)
                {
                    GetKeyToString(e.KeyData.ToString());
                    if (keyData.Split('+').Length >= CurrentKeyDown.Split('+').Length)
                    {
                        CurrentKeyDown = keyData;
                    }
                    return;
                }
            }
            catch (Exception)
            {
                return;
            }
           

            //if (IsLeftDown)
            //{
            //    //Log.Info("IsLeftDown : " + IsLeftDown.ToString());
            //    return;
            //}

            try
            {
                //if (Common.IsRunning)
                {
                    string keyData = GetKeyToString(e.KeyData.ToString());

                    if (Common.IsIgnoreCtrlAltV && keyData == "Ctrl + Alt + V")
                    {
                        Common.IsIgnoreCtrlAltV = false;
                        e.Handled = true;
                        return;
                    }

                    if (keyData.StartsWith("Alt + Num") || keyData.StartsWith("Alt + Shift + "))
                    {
                        string[] tabShortcuts =
                        {
                            App.Setting.Tab1Shortcut, App.Setting.Tab2Shortcut, App.Setting.Tab3Shortcut,
                            App.Setting.Tab4Shortcut, App.Setting.Tab5Shortcut, App.Setting.Tab6Shortcut,
                            App.Setting.Tab7Shortcut, App.Setting.Tab8Shortcut, App.Setting.Tab9Shortcut
                        };
                        for (int i = 0; i < tabShortcuts.Length; i++)
                        {
                            if (keyData == tabShortcuts[i] || keyData == "Alt + Shift + D" + (i + 1))
                            {
                                MainWindow.SetActive(i + 1);
                                e.Handled = true;
                                return;
                            }
                        }
                    }
                    if (keyData.StartsWith("Alt +") || keyData.StartsWith("Alt + Shift + "))
                    {
                        if (keyData == App.Setting.LeftTabShortcut || keyData == "Alt + Shift + Left")
                        {
                            int activeOrder = Common.CurrentTabIndex - 1;
                            if (activeOrder <= 0)
                            {
                                activeOrder = MainWindow.ExcelCount;
                            }
                            MainWindow.SetActive(activeOrder);
                            e.Handled = true;
                            return;
                        }
                        else if (keyData == App.Setting.RightTabShortcut || keyData == "Alt + Shift + Right")
                        {
                            int activeOrder = Common.CurrentTabIndex + 1;
                            if (activeOrder > MainWindow.ExcelCount)
                            {
                                activeOrder = 1;
                            }
                            MainWindow.SetActive(activeOrder);
                            e.Handled = true;
                            return;
                        }
                    }

                    if (Common.TestMode && Common.IsExcelActive && keyData.StartsWith("Alt +"))
                    {
                        //if (keyData == "Alt + C")
                        //{
                        //    Common.SetTimeout(() => Common.ExcelGetText(), 100);
                        //    e.Handled = true;
                        //    return;
                        //}
                        //else if (keyData == "Alt + V")
                        //{
                        //    Common.SetTimeout(() => Common.ExcelSetText(Common.GetNextText()), 100);
                        //    e.Handled = true;
                        //    return;
                        //}
                        //else 
                        if (keyData == "Alt + B")
                        {
                            Common.SetTimeout(() => ExcelHelper.SetBackground(System.Drawing.Color.FromArgb(217, 217, 217)), 100);
                            e.Handled = true;
                            return;
                        }
                        else if (keyData == "Alt + H")
                        {
                            Common.SetTimeout(() => ExcelHelper.AddBracket(), 100);
                            e.Handled = true;
                            return;
                        }
                        else if (keyData == "Alt + T")
                        {
                            Common.SetTimeout(() => ExcelHelper.SetBackground(), 100);
                            e.Handled = true;
                            return;
                        }
                        else if (keyData == "Alt + Y")
                        {
                            Common.SetTimeout(() => ExcelHelper.SetBackground(System.Drawing.Color.FromArgb(255, 255, 0)), 100);
                            e.Handled = true;
                            return;
                        }
                        else if (keyData == "Alt + OemMinus")
                        {
                            Common.SetTimeout(() => ExcelHelper.PlusText(-1), 200);
                            e.Handled = true;
                            return;
                        }
                        else if (keyData == "Alt + Oem7")
                        {
                            Common.SetTimeout(() => ExcelHelper.PlusText(), 200);
                            e.Handled = true;
                            return;
                        }
                        else if (keyData == "Alt + Oem5")
                        {
                            Common.SetTimeout(() => ExcelHelper.AutoNum(), 200);
                            e.Handled = true;
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppLog.Warn("ショートカット処理に失敗しました。", ex);
            }

            try
            {
                if (!e.Control && e.Alt && !e.Shift && Common.IsExcelActive)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            Common.SetTimeout(() => ExcelHelper.Copy(), 100);
                            e.Handled = true;
                            return;
                        case Keys.V:
                            Common.IsIgnoreCtrlAltV = true;
                            e.Handled = true;
                            Common.SetTimeout(() => ExcelHelper.Paste(), 200);
                            return;
                    }
                }

                if (e.Control && e.Alt && !e.Shift && TryFavoriteDigit(e.KeyCode, out string addKey))
                {
                    Common.SetTimeout(() => FavoriteHelper.AddFavoriteCell(addKey), 200);
                    e.Handled = true;
                    return;
                }

                if (e.Control && !e.Alt && !e.Shift)
                {
                    if (e.KeyCode == Keys.NumPad0)
                    {
                        ShowFavoriteWindow();
                        e.Handled = true;
                        return;
                    }
                    if (TryFavoriteDigit(e.KeyCode, out string moveKey) && e.KeyCode >= Keys.NumPad1 && e.KeyCode <= Keys.NumPad9)
                    {
                        Common.SetTimeout(() => FavoriteHelper.MoveFavoriteCell(moveKey), 200);
                        e.Handled = true;
                        return;
                    }
                }

                if (e.Control && !e.Alt && e.Shift)
                {
                    if (e.KeyCode == Keys.D0)
                    {
                        ShowFavoriteWindow();
                        e.Handled = true;
                        return;
                    }
                    if (TryFavoriteDigit(e.KeyCode, out string moveKey) && e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D9)
                    {
                        Common.SetTimeout(() => FavoriteHelper.MoveFavoriteCell(moveKey), 200);
                        e.Handled = true;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                AppLog.Warn("お気に入りショートカット処理に失敗しました。", ex);
            }

        }

        private static bool TryFavoriteDigit(Keys key, out string digit)
        {
            if (key >= Keys.NumPad1 && key <= Keys.NumPad9)
            {
                digit = ((int)(key - Keys.NumPad0)).ToString();
                return true;
            }
            if (key >= Keys.D1 && key <= Keys.D9)
            {
                digit = ((int)(key - Keys.D0)).ToString();
                return true;
            }
            digit = null;
            return false;
        }

        private static void ShowFavoriteWindow()
        {
            var app = System.Windows.Application.Current;
            if (app == null)
            {
                return;
            }
            app.Dispatcher.Invoke(() =>
            {
                MainWindow.TopmostFlg = false;
                try
                {
                    FavoriteWindow fw = new FavoriteWindow();
                    fw.ShowDialog();
                }
                finally
                {
                    MainWindow.TopmostFlg = true;
                }
            });
        }

        public void OnHookKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                string keyData = GetKeyToString(e.KeyData.ToString());
                //Debug.WriteLine("KeyUp : " + keyData);
                if (keyData == "Alt + V")
                {
                    Common.IsIgnoreCtrlAltV = false;
                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                AppLog.Warn("キー入力の解析に失敗しました。", ex);
            }

        }
    }
}
