using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections;
using System.Collections.Generic;

namespace Y41.OfficeTool.NpoiExcelHelper
{
    public class RowHelper : IRow
    {

        public delegate MergeDel MergeDel(int cellStart, int cellEnd);

        /// <summary> 
        /// 合併儲存格 
        /// </summary> 
        /// <param name="cellStart"></param> 
        /// <param name="cellEnd"></param> 
        public MergeDel Merge(int cellStart, int cellEnd)
        {

            var sheet = _row.Sheet;

            sheet.AddMergedRegion(new CellRangeAddress(_row.RowNum, _row.RowNum, cellStart, cellEnd));

            return Merge;
        }

        private readonly IRow _row;



        public CellHelper this[int cellIndex]
        {
            get
            {
                var cell = _row.GetCell(cellIndex);
                if (cell == null)
                {
                    cell = _row.CreateCell(cellIndex);
                }
                return new CellHelper(cell);
            }
        }

        public RowHelper(IRow row)
        {
            _row = row;
        }

        public ICell CreateCell(int column)
        {
            return _row.CreateCell(column);
        }

        public ICell CreateCell(int column, CellType type)
        {
            return _row.CreateCell(column, type);
        }

        public void RemoveCell(ICell cell)
        {
            _row.RemoveCell(cell);
        }

        public ICell GetCell(int cellnum)
        {
            return _row.GetCell(cellnum);
        }

        public ICell GetCell(int cellnum, MissingCellPolicy policy)
        {
            return _row.GetCell(cellnum, policy);
        }

        public IEnumerator GetEnumerator()
        {
            return _row.GetEnumerator();
        }

        public void MoveCell(ICell cell, int newColumn)
        {
            _row.MoveCell(cell, newColumn);
        }

        public IRow CopyRowTo(int targetIndex)
        {
            return _row.CopyRowTo(targetIndex);
        }

        public ICell CopyCell(int sourceIndex, int targetIndex)
        {
            return _row.CopyCell(sourceIndex, targetIndex);
        }

        public bool HasCustomHeight()
        {
            return _row.HasCustomHeight();
        }

        IEnumerator<ICell> IEnumerable<ICell>.GetEnumerator()
        {
            return _row.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return ((System.Collections.IEnumerable)_row).GetEnumerator();
        }

        public int RowNum { get { return _row.RowNum; } set { _row.RowNum = value; } }

        public short FirstCellNum { get { return _row.FirstCellNum; } private set { } }

        public short LastCellNum { get { return _row.LastCellNum; } private set { } }

        public int PhysicalNumberOfCells { get { return _row.PhysicalNumberOfCells; } private set { } }

        public bool ZeroHeight { get { return _row.ZeroHeight; } set { _row.ZeroHeight = value; } }

        public short Height { get { return _row.Height; } set { _row.Height = value; }}

        public float HeightInPoints { get { return _row.HeightInPoints; } set { _row.HeightInPoints = value; } }

        public bool IsFormatted { get { return _row.IsFormatted; } private set { } }

        public ISheet Sheet { get { return _row.Sheet; } private set { } }

        public ICellStyle RowStyle { get { return _row.RowStyle; } set { _row.RowStyle = value; } }

        public List<ICell> Cells { get { return _row.Cells; } private set { }}

        public int OutlineLevel => _row.OutlineLevel;

        public bool? Hidden { get => _row.Hidden; set => _row.Hidden = value; }

        public bool? Collapsed { get => _row.Collapsed; set => _row.Collapsed = value; }
    }
}
