using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace Y41.OfficeTool.NpoiExcelHelper
{
    public class CellHelper : ICell
    {
        public delegate MergeDel MergeDel(int cellEnd);
        /// <summary> 
        /// 合併儲存格 
        /// </summary> 
        /// <param name="cellEnd"></param> 

        public MergeDel Merge(int cellEnd)
        {
            var sheet = _cell.Sheet;
            sheet.AddMergedRegion(new CellRangeAddress(_cell.RowIndex, _cell.RowIndex, _cell.ColumnIndex, cellEnd));
            var nextCell = new RowHelper(_cell.Row)[_cell.ColumnIndex + 1];
            return nextCell.Merge;
        }

        public CellHelper SetValue<T>(T value, string toStringContent = null)
        {
            if (value != null)
            {
                switch (Type.GetTypeCode(value.GetType()))//switch (Type.GetTypeCode(typeof(T))) 
                {
                    case TypeCode.Boolean:
                        var booleanValue = Convert.ToBoolean(value);
                        _cell.SetCellValue(booleanValue);
                        break;
                    case TypeCode.Int32:
                    case TypeCode.Int64:
                    case TypeCode.Double:
                    case TypeCode.Decimal:
                        var doubleValue = Convert.ToDouble(value);
                        _cell.SetCellValue(doubleValue);
                        break;
                    case TypeCode.String:
                        var stringValue = Convert.ToString(value);
                        _cell.SetCellValue(stringValue);
                        break;
                    case TypeCode.DateTime:
                        var dtValue = Convert.ToDateTime(value);
                        if (string.IsNullOrWhiteSpace(toStringContent))
                        {
                            _cell.SetCellValue(dtValue);
                        }
                        else
                        {
                            _cell.SetCellValue(dtValue.ToString(toStringContent));
                        }
                        break;
                }
            }
            return this;
        }
        private readonly ICell _cell;

        public CellHelper(ICell cell)
        {
            _cell = cell;
        }

        public void SetCellType(CellType cellType)
        {
            _cell.SetCellType(cellType);
        }

        public void SetCellValue(double value)
        {
            _cell.SetCellValue(value);
        }

        public void SetCellErrorValue(byte value)
        {
            _cell.SetCellErrorValue(value);
        }

        public void SetCellValue(DateTime value)
        {
            _cell.SetCellValue(value);
        }

        public void SetCellValue(IRichTextString value)
        {
            _cell.SetCellValue(value);
        }

        public void SetCellValue(string value)
        {
            _cell.SetCellValue(value);
        }

        public ICell CopyCellTo(int targetIndex)
        {
            return _cell.CopyCellTo(targetIndex);
        }

        public void SetCellFormula(string formula)
        {
            _cell.SetCellFormula(formula);
        }

        public void SetCellValue(bool value)
        {
            _cell.SetCellValue(value);
        }

        public void SetAsActiveCell()
        {
            _cell.SetAsActiveCell();
        }

        public void RemoveCellComment()
        {
            _cell.RemoveCellComment();
        }

        public int ColumnIndex
        {
            get { return _cell.ColumnIndex; }
            private set { }
        }

        public int RowIndex
        {
            get { return _cell.RowIndex; }
            private set { }
        }

        public ISheet Sheet
        {
            get { return _cell.Sheet; }
            private set { }
        }

        public IRow Row { get; private set; }

        public CellType CellType
        {
            get { return _cell.CellType; }
            private set { }
        }

        public CellType CachedFormulaResultType { get { return _cell.CachedFormulaResultType; } private set { } }

        public string CellFormula { get { return _cell.CellFormula; } set { _cell.CellFormula = value; } }

        public double NumericCellValue { get { return _cell.NumericCellValue; } private set { } }

        public DateTime DateCellValue { get { return _cell.DateCellValue; } private set { } }

        public IRichTextString RichStringCellValue { get { return _cell.RichStringCellValue; } private set { } }

        public byte ErrorCellValue { get; private set; }

        public string StringCellValue { get { return _cell.StringCellValue; } private set { } }

        public bool BooleanCellValue { get { return _cell.BooleanCellValue; } private set { } }

        public ICellStyle CellStyle
        {
            get { return _cell.CellStyle; }
            set { _cell.CellStyle = value; }
        }

        public IComment CellComment { get { return _cell.CellComment; } set { _cell.CellComment = value; } }

        public IHyperlink Hyperlink { get { return _cell.Hyperlink; } set { _cell.Hyperlink = value; } }

        public CellRangeAddress ArrayFormulaRange { get { return _cell.ArrayFormulaRange; } private set { } }

        public bool IsPartOfArrayFormulaGroup { get { return _cell.IsPartOfArrayFormulaGroup; } private set { } }

        public bool IsMergedCell { get { return _cell.IsMergedCell; } private set { } }

        public StyleHelper Style()
        {
            return new StyleHelper(_cell);
        }

        public StyleHelper Style(Action<StyleHelper> act)
        {
            var styleHelper = new StyleHelper(_cell);
            act(styleHelper);
            return styleHelper;
        }

        public StyleHelper Style(StyleHelper sh)
        {
            new StyleHelper(_cell, sh);//.Ok();
            var styleHelper = new StyleHelper(_cell, sh);
            return styleHelper;
        }

        public void RemoveHyperlink()
        {
            _cell.RemoveHyperlink();
        }

        public CellType GetCachedFormulaResultTypeEnum()
        {
            return _cell.GetCachedFormulaResultTypeEnum();
        }
    }
}
