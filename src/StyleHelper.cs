using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Y41.OfficeTool.NpoiExcelHelper
{
    public class StyleHelper
    {
        private readonly ICell _cell;
        private IWorkbook _workbook;
        private IDataFormat _format;
        private ICellStyle _cellStyle;
        private IFont _font;

        public IWorkbook Workbook
        {
            get { return _workbook; }
            set { _workbook = value; }
        }
        public IDataFormat DataFormat
        {
            get { return _format; }
            set { _format = value; }
        }
        public IFont Font
        {
            get { return _font; }
            set { _font = value; }
        }
        public ICellStyle CellStyle
        {
            get
            {
                if (_cellStyle == null)
                    _cellStyle = _workbook.CreateCellStyle();

                return _cellStyle;
            }
        }

        public StyleHelper(IWorkbook workbook)
        {
            _cellStyle = workbook.CreateCellStyle();
            _workbook = workbook;
            _format = _workbook.CreateDataFormat();
            _font = _workbook.CreateFont();
        }

        public StyleHelper Clone()
        {
            var newStyleHelper = new StyleHelper(_workbook);
            newStyleHelper.Workbook = _workbook;
            newStyleHelper.CellStyle.CloneStyleFrom(_cellStyle);
            //var newDataFormat = _workbook.CreateDataFormat();
            newStyleHelper.CellStyle.DataFormat = newStyleHelper.DataFormat.GetFormat(_cellStyle.GetDataFormatString());

            //newStyleHelper.DataFormat = newDataFormat;
            newStyleHelper.Font = _workbook.CreateFont();//newStyleHelper.CellStyle.GetFont(_workbook);
            newStyleHelper.Font.CloneStyleFrom(_font);
            return newStyleHelper;
        }

        public StyleHelper(ICell cell, StyleHelper styleHelper)
        {
            _cell = cell;
            _cellStyle = styleHelper.CellStyle;
            _workbook = styleHelper.Workbook;
            _format = styleHelper.DataFormat;
            _font = styleHelper.Font;
        }

        public StyleHelper(ICell cell)
        {
            _cell = cell;
            _workbook = cell.Sheet.Workbook;
            _cellStyle = _workbook.CreateCellStyle();
            _format = _workbook.CreateDataFormat();
            _font = _workbook.CreateFont();
        }
        /// <summary>
        /// 輸出格式
        /// </summary>
        /// <param name="formatString"></param>
        /// <returns></returns>
        public StyleHelper Format(string formatString)
        {
            _cellStyle.DataFormat = _format.GetFormat(formatString);
            return this;
        }
        public StyleHelper 百分比格式()
        {
            return Format("0.0%");
        }
        public StyleHelper 無小數千分位格式()
        {
            return Format("#,##0");
        }
        /// <summary>
        /// 字體大小
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public StyleHelper FontSize(short size)
        {
            _font.FontHeightInPoints = size;
            return this;
        }
        /// <summary>
        /// 文字顏色<param/> IndexedColors.BrightGreen.Index
        /// </summary>
        /// <param name="colorIndex"></param>
        /// <returns></returns>
        public StyleHelper FontColor(short colorIndex)
        {
            _font.Color = colorIndex;
            return this;
        }
        /// <summary>
        /// 字體
        /// </summary>
        /// <param name="fontName"></param>
        /// <returns></returns>
        public StyleHelper FontName(string fontName)
        {
            _font.FontName = fontName;
            return this;
        }
        /// <summary>
        /// 置中
        /// </summary>
        /// <returns></returns>
        public StyleHelper AlignCenter()
        {
            _cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            _cellStyle.Alignment = HorizontalAlignment.Center;
            return this;
        }


        /// <summary>
        /// 靠左
        /// </summary>
        /// <returns></returns>
        public StyleHelper AlignLeft()
        {
            _cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            _cellStyle.Alignment = HorizontalAlignment.Left;
            return this;
        }
        /// <summary>
        /// 靠右
        /// </summary>
        /// <returns></returns>
        public StyleHelper AlignRight()
        {
            _cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            _cellStyle.Alignment = HorizontalAlignment.Right;
            return this;
        }
        /// <summary>
        /// 取得Style
        /// </summary>
        /// <returns></returns>
        public ICellStyle GetCellStyle()
        {
            return _cellStyle;
        }
        /// <summary>
        /// 取得字型
        /// </summary>
        /// <returns></returns>
        public IFont GetFont()
        {
            return _font;
        }
        /// <summary>
        /// 粗體
        /// </summary>
        /// <returns></returns>
        public StyleHelper Bold()
        {
            _font.IsBold = true;
            //_font.Boldweight = (short)FontBoldWeight.Bold;
            return this;
        }
        /// <summary>
        /// 框線
        /// </summary>
        /// <returns></returns>
        public StyleHelper Border()
        {
            _cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;//粗
            _cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;//細實線
            _cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;//虛線
            _cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            return this;
        }
        /// <summary>
        /// 上框線
        /// </summary>
        /// <returns></returns>
        public StyleHelper BorderTop()
        {
            _cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            return this;
        }

        /// <summary>
        /// 下框線
        /// </summary>
        /// <returns></returns>
        public StyleHelper BorderBottom()
        {
            _cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            return this;
        }

        /// <summary>
        /// 左框線
        /// </summary>
        /// <returns></returns>
        public StyleHelper BorderLeft()
        {
            _cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            return this;
        }

        /// <summary>
        /// 右框線
        /// </summary>
        /// <returns></returns>
        public StyleHelper BorderRight()
        {
            _cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            return this;
        }

        /// <summary>
        /// 下框線 雙底
        /// </summary>
        /// <returns></returns>
        public StyleHelper BorderBottomDouble()
        {
            _cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            return this;
        }

        /// <summary>
        /// 背景顏色<param/> IndexedColors.BrightGreen.Index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public StyleHelper BackgroudColor(short index)
        {
            _cellStyle.FillForegroundColor = index;
            _cellStyle.FillPattern = FillPattern.SolidForeground;
            return this;
        }
        /// <summary>
        /// 完成設定，繪製儲存格
        /// </summary>
        public void Ok()
        {
            _cellStyle.SetFont(_font);
            _cell.CellStyle = _cellStyle;
        }
    }
}
