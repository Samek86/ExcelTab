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
                    try
                    {
                        app.Dispose();
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            return GetActiveApplication();
        }
    }
}
