using System;
using System.Diagnostics;
using System.IO;

namespace ExcelTab.CLASSES
{
    public static class AppLog
    {
        private static readonly object Sync = new object();
        private static readonly string LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelTab.log");

        public static void Info(string message)
        {
            Write("INFO", message, null);
        }

        public static void Warn(string message, Exception ex = null)
        {
            Write("WARN", message, ex);
        }

        public static void Error(string message, Exception ex = null)
        {
            Write("ERROR", message, ex);
        }

        private static void Write(string level, string message, Exception ex)
        {
            string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " [" + level + "] " + message;
            if (ex != null)
            {
                line += " " + ex;
            }
            Debug.WriteLine(line);
            try
            {
                lock (Sync)
                {
                    File.AppendAllText(LogPath, line + Environment.NewLine);
                }
            }
            catch
            {
            }
        }
    }
}
