

NpoiExcelHelper
===================

此專案以 [NPOI](https://www.nuget.org/packages/NPOI/) 為基礎進行擴充，讓 NPOI 使用上可更加簡單、易懂、更好維護；
NpoiExcelHelper 讓你能直接透過索引取得資料列、儲存格，並大幅縮短「編輯資料、設定樣式、...」等動作的程式碼，用更簡潔的方式來體驗 NPOI 強大的功能。


範例
===================

當你想對 Excel 進行「在第 0 列、第 1 欄的位置寫入 "Text" 文字，並將文字樣式設定成粗體」，原始程式碼如下：
```csharp
var sheet = workbook.CreateSheet();
var row = sheet.GetRow(0);
var cell = row.GetCell(1);
cell.SetCellValue("Text");
var font = workbook.CreateFont();
font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.BOLD; 
cell.CellStyle = workbook.CreateCellStyle();
cell.CellStyle.SetFont(font);

```
使用 NpoiExcelHelper 精簡後的程式：
```csharp
var sheet = workbook.CreateSheet().Helper();
sheet[0][1].SetValue("Text").Style().Bold().Render();
```
設計特色
=================
* 運用 Indexer 取得資料列、儲存格
* 採用 Fluent Api 設計，讓程式碼更加直覺
* 簡化樣式設定方法，並參考CSS屬性設計，網頁開發者可更快上手

模組概述
=================

1. SheetHelper: 具 Sheet 物件的各種功能，並加入Indexer以產生/取得資料列(RowHelper)。
2. RowHelper: 具 Row 物件的各種功能，並加入Indexer以產生/取得儲存格(CellHelper)。
3. CellHelper: 具 Cell 物件的各種功能，並加入簡化版的資料讀取/寫入、樣式設定方法。
4. StyleHelper: 樣式管理物件，可用類似撰寫CSS方式設定外觀，並將樣式套用於CellHelper。


系統需求
=============
NPOI 2.4.1.0  
.Net Framework 4.5 以上版本


