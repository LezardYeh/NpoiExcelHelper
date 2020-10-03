using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Y41.OfficeTool.NpoiExcelHelper.Extensions
{
    public static class SheetExtension
    {
        public static SheetHelper Helper(this ISheet _sheet)
        {
            return new SheetHelper(_sheet);
        }
    }
}
