using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections;
using System.Collections.Generic;

namespace Y41.OfficeTool.NpoiExcelHelper
{
    public class SheetHelper : ISheet
    {
        public delegate MergeDelegate MergeDelegate(int rowStart, int rowEnd, int cellStart, int cellEnd);

        /// <summary> 
        /// 合併儲存格 
        /// </summary> 
        /// <param name="rowStart"></param> 
        /// <param name="rowEnd"></param> 
        /// <param name="cellStart"></param> 
        /// <param name="cellEnd"></param> 
        public MergeDelegate Merge(int rowStart, int rowEnd, int cellStart, int cellEnd)
        {
            _sheet.AddMergedRegion(new CellRangeAddress(rowStart, rowEnd, cellStart, cellEnd));
            return Merge;
        }

        public RowHelper this[int index]
        {
            get
            {
                var row = _sheet.GetRow(index);
                if (row == null)
                {
                    row = _sheet.CreateRow(index);
                }
                return new RowHelper(row);
            }
        }

        private readonly ISheet _sheet;
        public SheetHelper(ISheet sheet)
        {
            _sheet = sheet;
        }

        public IRow CreateRow(int rownum)
        {
            return _sheet.CreateRow(rownum);
        }

        public void RemoveRow(IRow row)
        {
            _sheet.RemoveRow(row);
        }

        public IRow GetRow(int rownum)
        {
            return _sheet.GetRow(rownum);
        }

        public void SetColumnHidden(int columnIndex, bool hidden)
        {
            _sheet.SetColumnHidden(columnIndex, hidden);
        }

        public bool IsColumnHidden(int columnIndex)
        {
            return _sheet.IsColumnHidden(columnIndex);
        }

        public IRow CopyRow(int sourceIndex, int targetIndex)
        {
            return CopyRow(sourceIndex, targetIndex);
        }

        public void SetColumnWidth(int columnIndex, int width)
        {
            _sheet.SetColumnWidth(columnIndex, width);
        }

        public int GetColumnWidth(int columnIndex)
        {
            return _sheet.GetColumnWidth(columnIndex);
        }

        public ICellStyle GetColumnStyle(int column)
        {
            return _sheet.GetColumnStyle(column);
        }

        public int AddMergedRegion(CellRangeAddress region)
        {
            return _sheet.AddMergedRegion(region);
        }

        public void RemoveMergedRegion(int index)
        {
            _sheet.RemoveMergedRegion(index);
        }

        public CellRangeAddress GetMergedRegion(int index)
        {
            return _sheet.GetMergedRegion(index);
        }

        public IEnumerator GetRowEnumerator()
        {
            return _sheet.GetRowEnumerator();
        }

        public IEnumerator GetEnumerator()
        {
            return _sheet.GetEnumerator();
        }

        public double GetMargin(MarginType margin)
        {
            return _sheet.GetMargin(margin);
        }

        public void SetMargin(MarginType margin, double size)
        {
            _sheet.SetMargin(margin, size);
        }

        public void ProtectSheet(string password)
        {
            _sheet.ProtectSheet(password);
        }

        public void ShowInPane(short toprow, short leftcol)
        {
            _sheet.ShowInPane(toprow, leftcol);
        }

        public void ShiftRows(int startRow, int endRow, int n)
        {
            _sheet.ShiftRows(startRow, endRow, n);
        }

        public void ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight)
        {
            _sheet.ShiftRows(startRow, endRow, n, copyRowHeight, resetOriginalRowHeight);
        }

        public void CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow)
        {
            _sheet.CreateFreezePane(colSplit, rowSplit, leftmostColumn, topRow);
        }

        public void CreateFreezePane(int colSplit, int rowSplit)
        {
            _sheet.CreateFreezePane(colSplit, rowSplit);
        }

        public void CreateSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, PanePosition activePane)
        {
            _sheet.CreateSplitPane(xSplitPos, ySplitPos, leftmostColumn, topRow, activePane);
        }

        public bool IsRowBroken(int row)
        {
            return _sheet.IsRowBroken(row);
        }

        public void RemoveRowBreak(int row)
        {
            _sheet.RemoveRowBreak(row);
        }

        public void SetActiveCellRange(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            _sheet.SetActiveCellRange(firstRow, lastRow, firstColumn, lastColumn);
        }

        public void SetActiveCellRange(List<CellRangeAddress8Bit> cellranges, int activeRange, int activeRow, int activeColumn)
        {
            _sheet.SetActiveCellRange(cellranges, activeRange, activeRow, activeColumn);
        }

        public void SetColumnBreak(int column)
        {
            _sheet.SetColumnBreak(column);
        }

        public void SetRowBreak(int row)
        {
            _sheet.SetRowBreak(row);
        }

        public bool IsColumnBroken(int column)
        {
            return _sheet.IsColumnBroken(column);
        }

        public void RemoveColumnBreak(int column)
        {
            _sheet.RemoveColumnBreak(column);
        }

        public void SetColumnGroupCollapsed(int columnNumber, bool collapsed)
        {
            _sheet.SetColumnGroupCollapsed(columnNumber, collapsed);
        }

        public void GroupColumn(int fromColumn, int toColumn)
        {
            _sheet.GroupColumn(fromColumn, toColumn);
        }

        public void UngroupColumn(int fromColumn, int toColumn)
        {
            _sheet.UngroupColumn(fromColumn, toColumn);
        }

        public void GroupRow(int fromRow, int toRow)
        {
            _sheet.GroupRow(fromRow, toRow);
        }

        public void UngroupRow(int fromRow, int toRow)
        {
            _sheet.UngroupRow(fromRow, toRow);
        }

        public void SetRowGroupCollapsed(int row, bool collapse)
        {
            _sheet.SetRowGroupCollapsed(row, collapse);
        }

        public void SetDefaultColumnStyle(int column, ICellStyle style)
        {
            _sheet.SetDefaultColumnStyle(column, style);
        }

        public void AutoSizeColumn(int column)
        {
            _sheet.AutoSizeColumn(column);
        }

        public void AutoSizeColumn(int column, bool useMergedCells)
        {
            _sheet.AutoSizeColumn(column, useMergedCells);
        }

        public IDrawing CreateDrawingPatriarch()
        {
            return _sheet.CreateDrawingPatriarch();
        }

        public void SetActive(bool value)
        {
            _sheet.SetActive(value);
        }

        public ICellRange<ICell> SetArrayFormula(string formula, CellRangeAddress range)
        {
            return _sheet.SetArrayFormula(formula, range);
        }

        public ICellRange<ICell> RemoveArrayFormula(ICell cell)
        {
            return _sheet.RemoveArrayFormula(cell);
        }

        public bool IsMergedRegion(CellRangeAddress mergedRegion)
        {
            return _sheet.IsMergedRegion(mergedRegion);
        }

        public IDataValidationHelper GetDataValidationHelper()
        {
            return _sheet.GetDataValidationHelper();
        }

        public void AddValidationData(IDataValidation dataValidation)
        {
            _sheet.AddValidationData(dataValidation);
        }

        public IAutoFilter SetAutoFilter(CellRangeAddress range)
        {
            return _sheet.SetAutoFilter(range);
        }

        public float GetColumnWidthInPixels(int columnIndex)
        {
            return _sheet.GetColumnWidthInPixels(columnIndex);
        }

        System.Collections.IEnumerator ISheet.GetRowEnumerator()
        {
            return _sheet.GetRowEnumerator();
        }

        System.Collections.IEnumerator ISheet.GetEnumerator()
        {
            return _sheet.GetEnumerator();
        }

        public void ShowInPane(int toprow, int leftcol)
        {
            _sheet.ShowInPane(toprow, leftcol);
        }

        public List<IDataValidation> GetDataValidations()
        {
            return _sheet.GetDataValidations();
        }

        public ISheet CopySheet(string Name)
        {
            return _sheet.CopySheet(Name);
        }

        public ISheet CopySheet(string Name, bool copyStyle)
        {
            return _sheet.CopySheet(Name, copyStyle);
        }

        public int GetColumnOutlineLevel(int columnIndex)
        {
            return _sheet.GetColumnOutlineLevel(columnIndex);
        }

        public bool IsDate1904()
        {
            return _sheet.IsDate1904();
        }

        public void SetZoom(int numerator, int denominator)
        {
            _sheet.SetZoom(numerator, denominator);
        }

        public IComment GetCellComment(int row, int column)
        {
            return _sheet.GetCellComment(row, column);
        }

        public void SetActiveCell(int row, int column)
        {
            _sheet.SetActiveCell(row, column);
        }

        public int PhysicalNumberOfRows
        {

            get { return _sheet.PhysicalNumberOfRows; }

            private set { }
        }

        public int FirstRowNum { get { return _sheet.FirstRowNum; } private set { } }

        public int LastRowNum { get { return _sheet.LastRowNum; } private set { } }

        public bool ForceFormulaRecalculation { get { return _sheet.ForceFormulaRecalculation; } set { _sheet.ForceFormulaRecalculation = value; } }

        public int DefaultColumnWidth { get { return _sheet.DefaultColumnWidth; } set { _sheet.DefaultColumnWidth = value; } }

        public short DefaultRowHeight { get { return _sheet.DefaultRowHeight; } set { _sheet.DefaultRowHeight = value; } }

        public float DefaultRowHeightInPoints { get { return _sheet.DefaultRowHeightInPoints; } set { _sheet.DefaultRowHeightInPoints = value; } }

        public bool HorizontallyCenter { get { return _sheet.HorizontallyCenter; } set { _sheet.HorizontallyCenter = value; } }

        public bool VerticallyCenter { get { return _sheet.VerticallyCenter; } set { _sheet.VerticallyCenter = value; } }

        public int NumMergedRegions { get { return _sheet.NumMergedRegions; } private set { } }

        public bool DisplayZeros { get { return _sheet.DisplayZeros; } set { _sheet.DisplayZeros = value; } }

        public bool Autobreaks { get { return _sheet.Autobreaks; } set { _sheet.Autobreaks = value; } }

        public bool DisplayGuts { get { return _sheet.DisplayGuts; } set { _sheet.DisplayGuts = value; } }

        public bool FitToPage { get { return _sheet.FitToPage; } set { _sheet.FitToPage = value; } }

        public bool RowSumsBelow { get { return _sheet.RowSumsBelow; } set { _sheet.RowSumsBelow = value; } }

        public bool RowSumsRight { get { return _sheet.RowSumsRight; } set { _sheet.RowSumsRight = value; } }

        public bool IsPrintGridlines { get { return _sheet.IsPrintGridlines; } set { _sheet.IsPrintGridlines = value; } }

        public IPrintSetup PrintSetup { get { return _sheet.PrintSetup; } private set { } }

        public IHeader Header { get { return _sheet.Header; } private set { } }

        public IFooter Footer { get { return _sheet.Footer; } private set { } }

        public bool Protect { get { return _sheet.Protect; } private set { } }

        public bool ScenarioProtect { get { return _sheet.ScenarioProtect; } private set { } }

        public short TabColorIndex { get { return _sheet.TabColorIndex; } set { _sheet.TabColorIndex = value; } }

        public IDrawing DrawingPatriarch { get { return _sheet.DrawingPatriarch; } private set { } }

        public short TopRow { get { return _sheet.TopRow; } set { _sheet.TopRow = value; } }

        public short LeftCol { get { return _sheet.LeftCol; } set { _sheet.LeftCol = value; } }

        public PaneInformation PaneInformation { get { return _sheet.PaneInformation; } private set { } }

        public bool DisplayGridlines { get { return _sheet.DisplayGridlines; } set { _sheet.DisplayGridlines = value; } }

        public bool DisplayFormulas { get { return _sheet.DisplayFormulas; } set { _sheet.DisplayFormulas = value; } }

        public bool DisplayRowColHeadings { get { return _sheet.DisplayRowColHeadings; } set { _sheet.DisplayRowColHeadings = value; } }

        public bool IsActive { get { return _sheet.IsActive; } set { _sheet.IsActive = value; } }

        public int[] RowBreaks { get { return _sheet.RowBreaks; } private set { } }

        public int[] ColumnBreaks { get { return _sheet.ColumnBreaks; } private set { } }

        public IWorkbook Workbook { get { return _sheet.Workbook; } private set { } }

        public string SheetName { get { return _sheet.SheetName; } private set { } }

        public bool IsSelected { get { return _sheet.IsSelected; } set { _sheet.IsSelected = value; } }

        public ISheetConditionalFormatting SheetConditionalFormatting { get { return _sheet.SheetConditionalFormatting; } private set { } }

        public bool IsRightToLeft { get { return _sheet.IsRightToLeft; } set { _sheet.IsRightToLeft = value; } }

        public CellRangeAddress RepeatingRows { get { return _sheet.RepeatingRows; } set { _sheet.RepeatingRows = value; } }

        public CellRangeAddress RepeatingColumns { get { return _sheet.RepeatingColumns; } set { _sheet.RepeatingColumns = value; } }


    }
}
