using System;
using Excel = NetOffice.ExcelApi;

namespace ExcelTab.CLASSES
{
    public static class ExcelAppHelper
    {
        public static Excel.Application GetActiveApplication()
        {
            try
            {
                return Excel.Application.GetActiveInstance(false);
            }
            catch (Exception ex)
            {
                AppLog.Warn("Excel インスタンス取得に失敗しました。", ex);
                return null;
            }
        }

        public static Excel.Application EnsureAlive(Excel.Application app)
        {
            if (app != null && !app.IsDisposed)
            {
                try
                {
                    _ = app.Workbooks.Count;
                    return app;
                }
                catch (Exception)
                {
                    Release(ref app);
                }
            }
            return GetActiveApplication();
        }

        public static void Release(ref Excel.Application app)
        {
            if (app == null)
            {
                return;
            }
            try
            {
                if (!app.IsDisposed)
                {
                    app.Dispose();
                }
            }
            catch (Exception ex)
            {
                AppLog.Warn("Excel Application の解放に失敗しました。", ex);
            }
            app = null;
        }
    }
}
