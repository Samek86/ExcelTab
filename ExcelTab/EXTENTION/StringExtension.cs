using System.Linq;

namespace ExcelTab.Extension
{
    public static class StringExtension
    {
        public static string[] Trim(this string[] value)
        {
            return value.Select(v => v.Trim()).ToArray();
        }
    }
}
