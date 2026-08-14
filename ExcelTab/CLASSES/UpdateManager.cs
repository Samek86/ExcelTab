using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;

namespace ExcelTab.CLASSES
{
    public class UpdateManager
    {
        public static string appUrl = @"\\kagami\Share_000661\Users\handa\APP\";

        public static void Update(string fileNameStartWith, string NewFilePath)
        {
            string exeFileName = Assembly.GetExecutingAssembly().CodeBase.Replace("file:///", string.Empty).Replace("/", "\\");
            string exeFolder = System.IO.Path.GetDirectoryName(exeFileName);
            //string newFileName = Path.GetFileName(NewFilePath);
            string newFileName = $"{fileNameStartWith}.exe";
            string batchName = $"Update{fileNameStartWith}.bat";

            string batchCommands = "";
            batchCommands += "@ECHO OFF\n";
            //batchCommands += "echo Update.... \n";
            batchCommands += "ping 127.0.0.1 > nul\n";
            batchCommands += "echo j | del /F \"" + exeFileName + "\"\n";
            batchCommands += "echo j | copy \"" + NewFilePath + "\" \"" + Path.Combine(exeFolder, newFileName) + "\"\n";
            //batchCommands += "mshta javascript:alert(\"Update Completed.\");close();\n";
            //batchCommands += "echo j | start \"" + Path.Combine(exeFolder, newFileName) + "  -v runAs\"\n";
            batchCommands += "echo j | start \"\" \"" + Path.Combine(exeFolder, newFileName) + "\"\n";
            batchCommands += $"echo j | del {batchName}";
            File.WriteAllText(batchName, batchCommands);
            Process.Start(batchName);
            App.Current.Shutdown();
        }

        public static bool StartUpdateCheck(string fileNameStartWith, Assembly asm)
        {
            try
            {
                string checkPath = appUrl;
                string[] files = Directory.GetFiles(checkPath);
                bool isUpdate = false;
                string NewFilePath = "";
                foreach (string file in files)
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    if (fileName.StartsWith(fileNameStartWith))
                    {
                        //Version newVer = new Version(fileName.Substring(fileName.IndexOf("v") + 1));
                        var newVerInfo = FileVersionInfo.GetVersionInfo(file);
                        Version newVer = new Version(newVerInfo.FileVersion);
                        Version thisVer = asm.GetName().Version;
                        int result = newVer.CompareTo(thisVer);
                        if (result > 0)
                        {
                            NewFilePath = file;
                            isUpdate = true;
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                if (isUpdate)
                {
                    var result = MessageBox.Show("新しいバージョンがあります。アップデートしますか？", "Update", MessageBoxButton.YesNo);
                    if (result == MessageBoxResult.Yes)
                    {
                        Update(fileNameStartWith, NewFilePath);
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
            }
            return false;
        }

        public static void UserUpdate(string fileNameStartWith, Assembly asm)
        {
            try
            {
                string checkPath = appUrl;
                string[] files = Directory.GetFiles(checkPath);
                bool isUpdate = false;
                string NewFilePath = "";
                foreach (string file in files)
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    if (fileName.StartsWith(fileNameStartWith))
                    {
                        //Version newVer = new Version(fileName.Substring(fileName.IndexOf("v") + 1));
                        var newVerInfo = FileVersionInfo.GetVersionInfo(file);
                        Version newVer = new Version(newVerInfo.FileVersion);
                        Version thisVer = asm.GetName().Version;
                        int result = newVer.CompareTo(thisVer);
                        if (result > 0)
                        {
                            NewFilePath = file;
                            isUpdate = true;
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                if (isUpdate)
                {
                    Update(fileNameStartWith, NewFilePath);
                }
                else
                {
                    MessageBox.Show("最新バージョンです。");
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
