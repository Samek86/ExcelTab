using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ExcelTab.CLASSES
{
    public static class ComHelper
    {
        public static void ReleaseCom(object com)
        {
            if (com == null)
            {
                return;
            }
            try
            {
                if (com is IDisposable disposable)
                {
                    disposable.Dispose();
                    return;
                }
                if (Marshal.IsComObject(com))
                {
                    Marshal.ReleaseComObject(com);
                }
            }
            catch (Exception ex)
            {
                AppLog.Warn("COM 解放に失敗しました。", ex);
            }
        }

        public static void ReleaseComList<T>(IEnumerable<T> items)
        {
            if (items == null)
            {
                return;
            }
            foreach (var item in items)
            {
                ReleaseCom(item);
            }
        }
    }
}
