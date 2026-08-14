using ExcelTab.ITEM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Media;
using YamlDotNet.Serialization;

namespace ExcelTab
{
    [Serializable]
    public class ApplicationSetting
    {
        public static readonly string Location = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelTabSetting.yaml");

        public Dictionary<string, Color> TabColorDic = new Dictionary<string, Color>();
        public Dictionary<string, string> TabNameDic = new Dictionary<string, string>();
        public string Tab1Shortcut { get; set; } = "Alt + NumPad1";
        public string Tab2Shortcut { get; set; } = "Alt + NumPad2";
        public string Tab3Shortcut { get; set; } = "Alt + NumPad3";
        public string Tab4Shortcut { get; set; } = "Alt + NumPad4";
        public string Tab5Shortcut { get; set; } = "Alt + NumPad5";
        public string Tab6Shortcut { get; set; } = "Alt + NumPad6";
        public string Tab7Shortcut { get; set; } = "Alt + NumPad7";
        public string Tab8Shortcut { get; set; } = "Alt + NumPad8";
        public string Tab9Shortcut { get; set; } = "Alt + NumPad9";
        public string LeftTabShortcut { get; set; } = "Alt + Left";
        public string RightTabShortcut { get; set; } = "Alt + Right";

        public Dictionary<string, FavoriteExcel> FavoriteCellDic = new Dictionary<string, FavoriteExcel>();

        public List<string> IgnoreFiles = new List<string> {
            "Relaxtools.xlam",
            "RelaxTools.xlam",
            "ATPVBAEN.XLAM",
            "FUNCRES.XLAM",
            "EUROTOOL.XLAM",
        };

        //public Dictionary<string, byte[]> ImageDic = new Dictionary<string, byte[]>();


        public static ApplicationSetting Load()
        {
            if (!File.Exists(Location))
                using (var fs = File.Create(Location)) { }

            ApplicationSetting setting = null;
            try
            {
                setting = Deserialize(Location);
            }
            catch (Exception)
            {
            }

            if (setting == null)
            {
                setting = new ApplicationSetting();
                setting.Save(false);
            }

            return setting;
        }

        /// <summary>
        /// 저장
        /// </summary>
        /// <param name="tabDataClearFlg"></param>
        public void Save(bool tabDataClearFlg = true)
        {
            Serialize(this, Location);
        }

        public bool Ensure(bool isChanged = false)
        {
            //if (Level == 0)
            //{
            //    Level = 1;
            //    isChanged |= true;
            //}
            return isChanged;
        }

        /// <summary>
        /// 파일로 출력
        /// </summary>
        /// <param name="setting"></param>
        /// <param name="path"></param>
        private static void Serialize(ApplicationSetting setting, string path)
        {
            var serializer = new SerializerBuilder().Build();
            var yml = serializer.Serialize(setting);
            using (var sr = new StreamWriter(path))
            {
                sr.Write(yml);
            }
        }

        /// <summary>
        /// 파일에서 불러오기
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private static ApplicationSetting Deserialize(string path)
        {
            using (var sr = new StreamReader(path))
            {
                using (var input = new StringReader(sr.ReadToEnd()))
                {
                    var deserializer = new DeserializerBuilder().IgnoreUnmatchedProperties().Build();
                    return deserializer.Deserialize<ApplicationSetting>(input);
                }
            }
        }


    }
}
