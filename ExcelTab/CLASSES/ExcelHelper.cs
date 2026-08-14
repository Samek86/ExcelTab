using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using ExcelTab.VIEW;
using static ExcelTab.CLASSES.Win32Helper;
using Excel = NetOffice.ExcelApi;

namespace ExcelTab.CLASSES
{
    public static class ExcelHelper
    {
        public static string PasteText { get; set; } = "";
        public static List<List<string>> AltCopyData { get; } = new List<List<string>>();

        public static string CellText(object value2)
        {
            if (value2 == null)
            {
                return "";
            }
            return Convert.ToString(value2) ?? "";
        }

        public static Excel.Workbook GetActiveBook()
        {
            var title = ActiveWindowTitle();
            if (string.IsNullOrEmpty(title))
            {
                return null;
            }

            foreach (TabItemControl item in MainWindow.TabList)
            {
                if (string.IsNullOrEmpty(item.FileName))
                {
                    continue;
                }
                if (title.IndexOf(item.FileName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return item.Book;
                }

                string stem = item.FileName;
                int ext = stem.LastIndexOf(".xls", StringComparison.OrdinalIgnoreCase);
                if (ext > 0)
                {
                    stem = stem.Substring(0, ext);
                }
                if (!string.IsNullOrEmpty(stem) && title.IndexOf(stem, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return item.Book;
                }
            }
            return null;
        }

        private static Excel.Worksheet ActiveSheet(Excel.Workbook book)
        {
            object sheet = book.Application.ActiveSheet;
            if (sheet is Excel.Worksheet worksheet)
            {
                return worksheet;
            }
            if (sheet != null && Marshal.IsComObject(sheet))
            {
                return new Excel.Worksheet(book.Application, sheet);
            }
            return null;
        }

        private static Excel.Range SelectionRange(Excel.Workbook book)
        {
            object sel = book.Application.Selection;
            if (sel is Excel.Range range)
            {
                return range;
            }
            if (sel != null && Marshal.IsComObject(sel))
            {
                return new Excel.Range(book.Application, sel);
            }
            return null;
        }

        public static void Copy()
        {
            var book = GetActiveBook();
            if (book == null)
            {
                return;
            }
            Excel.Worksheet sheet = ActiveSheet(book);
            Excel.Range ran = SelectionRange(book);
            if (sheet == null || ran == null)
            {
                return;
            }

            Excel.Areas areas = ran.Areas;
            StringBuilder sb = new StringBuilder();
            if (areas.Count <= 0)
            {
                return;
            }

            AltCopyData.Clear();
            for (int i = 1, cnt = areas.Count; i <= cnt; i++)
            {
                Excel.Range rng = areas[i];
                int rowsCount = rng.Rows.Count;
                int columnsCount = rng.Columns.Count;
                for (int row = 0; row < rowsCount; row++)
                {
                    List<string> rowList = new List<string>();
                    for (int col = 0; col < columnsCount; col++)
                    {
                        var r = (Excel.Range)sheet.Cells[rng.Row + row, rng.Column + col];
                        string val = CellText(r.Value2);
                        rowList.Add(val);
                        if (col != 0)
                        {
                            sb.Append("\t");
                        }
                        sb.Append(val);
                    }
                    if (row < rowsCount - 1)
                    {
                        sb.Append("\n");
                    }
                    AltCopyData.Add(rowList);
                }
                if (i < cnt)
                {
                    sb.Append("\n");
                }
            }

            App.Current.Dispatcher.Invoke(() =>
            {
                Clipboard.Clear();
                Clipboard.SetText(sb.ToString());
            });
        }

        public static void Paste()
        {
            System.Windows.Forms.SendKeys.SendWait("^v");
        }

        public static string GetText()
        {
            var book = GetActiveBook();
            if (book == null)
            {
                return PasteText;
            }
            PasteText = CellText(book.Application.ActiveCell.Value2);
            return PasteText;
        }

        public static void SetText(string text)
        {
            var book = GetActiveBook();
            if (book == null)
            {
                return;
            }
            book.Application.ActiveCell.Value2 = text;
        }

        public static void PlusText(int plus = 1)
        {
            var book = GetActiveBook();
            if (book == null)
            {
                return;
            }
            Excel.Worksheet sheet = ActiveSheet(book);
            Excel.Range ran = SelectionRange(book);
            if (sheet == null || ran == null)
            {
                return;
            }

            foreach (Excel.Range rng in ran.Areas)
            {
                int countRows = rng.Rows.Count;
                for (int i = 0; i < countRows; i++)
                {
                    var r = (Excel.Range)sheet.Cells[rng.Row + i, rng.Column];
                    string val = CellText(r.Value2);
                    if (val.Contains("-"))
                    {
                        int hyphenIndex = val.LastIndexOf("-");
                        string numText = val.Substring(hyphenIndex + 1);
                        if (int.TryParse(numText, out int num))
                        {
                            val = val.Substring(0, hyphenIndex + 1) + (num + plus);
                        }
                    }
                    r.Value2 = val;
                }
            }
        }

        public static void AutoNum()
        {
            var book = GetActiveBook();
            if (book == null)
            {
                return;
            }
            Excel.Worksheet sheet = ActiveSheet(book);
            Excel.Range ran = SelectionRange(book);
            if (sheet == null || ran == null)
            {
                return;
            }

            int index = 1;
            string prefix = null;
            foreach (Excel.Range rng in ran.Areas)
            {
                int countRows = rng.Rows.Count;
                for (int i = 0; i < countRows; i++)
                {
                    var r = (Excel.Range)sheet.Cells[rng.Row + i, rng.Column];
                    string val = CellText(r.Value2);
                    if (prefix == null && val.Contains("-"))
                    {
                        int hyphenIndex = val.LastIndexOf("-");
                        prefix = val.Substring(0, hyphenIndex + 1);
                        if (int.TryParse(val.Substring(hyphenIndex + 1), out int firstIndex))
                        {
                            index = firstIndex;
                        }
                    }
                    if (!string.IsNullOrEmpty(val))
                    {
                        r.Value2 = prefix + index;
                        index++;
                    }
                }
            }
        }

        public static void SetBackground(System.Drawing.Color? color = null)
        {
            var book = GetActiveBook();
            if (book == null)
            {
                return;
            }
            Excel.Range ran = SelectionRange(book);
            if (ran == null)
            {
                return;
            }

            foreach (Excel.Range rng in ran.Areas)
            {
                if (color == null)
                {
                    rng.Interior.ColorIndex = 0;
                }
                else
                {
                    rng.Interior.Color = (System.Drawing.Color)color;
                }
            }
        }

        public static void AddBracket()
        {
            var book = GetActiveBook();
            if (book == null)
            {
                return;
            }
            Excel.Worksheet sheet = ActiveSheet(book);
            Excel.Range ran = SelectionRange(book);
            if (sheet == null || ran == null)
            {
                return;
            }

            foreach (Excel.Range rng in ran.Areas)
            {
                int countRows = rng.Rows.Count;
                for (int i = 0; i < countRows; i++)
                {
                    var r = (Excel.Range)sheet.Cells[rng.Row + i, rng.Column];
                    r.Value2 = "「" + CellText(r.Value2) + "」";
                }
            }
        }

        public static string ClipboardGetText()
        {
            string result = null;
            try
            {
                result = Clipboard.GetText();
                if (string.IsNullOrEmpty(result))
                {
                    System.Windows.IDataObject data = Clipboard.GetDataObject();
                    if (data != null)
                    {
                        try
                        {
                            result = data.GetData(DataFormats.StringFormat)?.ToString();
                        }
                        catch (Exception ex)
                        {
                            AppLog.Warn("クリップボード文字列の取得に失敗しました。", ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppLog.Warn("クリップボードの取得に失敗しました。", ex);
            }
            return result;
        }

        public static ((int, int), (int, int)) RangeToInt(string range)
        {
            range = range.Replace("$", "");
            if (string.IsNullOrEmpty(range) || !range.Contains(":"))
            {
                var cellNum = StringToCellNum(range, false);
                return (cellNum, cellNum);
            }
            string[] rangeArray = range.Split(':');
            return (StringToCellNum(rangeArray[0], false), StringToCellNum(rangeArray[1], false));
        }

        public static (int, int) StringToCellNum(string cell, bool baseZeroFlg = true)
        {
            var result = StringToCell(cell);
            return (StringToColumnNum(result.Item1, baseZeroFlg), int.Parse(result.Item2) - (baseZeroFlg ? 1 : 0));
        }

        public static int StringToColumnNum(string column, bool baseZeroFlg = true)
        {
            column = column.ToUpper();
            int columnIndex;
            switch (column.Length)
            {
                case 1:
                    columnIndex = column.ElementAt(0) - 65;
                    break;
                case 2:
                    columnIndex = ((column.ElementAt(0) - 64) * 26) + (column.ElementAt(1) - 65);
                    break;
                case 3:
                    columnIndex = ((column.ElementAt(0) - 64) * 676) + ((column.ElementAt(1) - 64) * 26) + (column.ElementAt(2) - 65);
                    break;
                default:
                    return -1;
            }
            if (!baseZeroFlg)
            {
                columnIndex += 1;
            }
            return columnIndex;
        }

        public static (string, string) StringToCell(string cell)
        {
            int boundary = -1;
            for (int i = 0; i < cell.Length; i++)
            {
                if (int.TryParse(cell[i].ToString(), out _))
                {
                    boundary = i;
                    break;
                }
            }
            if (boundary < 0)
            {
                return (cell, "0");
            }
            return (cell.Substring(0, boundary), cell.Substring(boundary));
        }
    }
}
