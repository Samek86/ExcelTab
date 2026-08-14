using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTab.ITEM
{
    public class FavoriteExcel
    {
        public string FullName { get; set; }
        public string SheetName { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public bool IsCell { get; set; } = true;
    }
}
