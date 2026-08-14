using System;
using System.Threading;
using System.Windows;

namespace ExcelTab
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static ApplicationSetting Setting { get; set; }

        private static Mutex _instanceMutex;

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            bool createdNew;
            _instanceMutex = new Mutex(true, @"Local\ExcelTab.SingleInstance", out createdNew);
            if (!createdNew)
            {
                Shutdown();
                return;
            }

            Setting = ApplicationSetting.Load();
            Setting.Save();

            Common.CreateDirectory(Common.TempFolderPath);
            new MainWindow().Show();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            try
            {
                if (_instanceMutex != null)
                {
                    _instanceMutex.ReleaseMutex();
                    _instanceMutex.Dispose();
                    _instanceMutex = null;
                }
            }
            catch (Exception)
            {
            }
            base.OnExit(e);
        }
    }
}
