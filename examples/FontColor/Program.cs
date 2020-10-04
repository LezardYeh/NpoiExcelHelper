using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using Y41.OfficeTool.NpoiExcelHelper.Extensions;

namespace FontColor
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet("Test").Helper();

            //sheetHelper[rowNumber, colNumber]
            sheet[0][1].SetValue("Column01");
            sheet[0][2].SetValue("Column02");
            sheet[0][3].SetValue("Column03");

            sheet[1][0].SetValue("Row1");
            sheet[2][0].SetValue("Row2");
            sheet[3][0].SetValue("Row3").Style().FontColor(IndexedColors.Red.Index).Bold().Render();

            var file = new FileStream(@"d:\npoi_helper_font_color_example.xls", FileMode.Create);
            workbook.Write(file);
            file.Close();
            System.Diagnostics.Process.Start(@"d:\npoi_helper_font_color_example.xls");
        }
    }
}
