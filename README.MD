

NpoiExcelHelper
===================

���M�ץH [NPOI](https://www.nuget.org/packages/NPOI/) ����¦�i���X�R�A�� NPOI �ϥΤW�i��[²��B�����B��n���@�F
NpoiExcelHelper ���A�ઽ���z�L���ި��o��ƦC�B�x�s��A�äj�T�Y�u�u�s���ơB�]�w�˦��B...�v���ʧ@���{���X�A�Χ�²�䪺�覡������ NPOI �j�j���\��C


�d��
===================

���A�Q�� Excel �i��u�b�� 0 �C�B�� 1 �檺��m�g�J "Text" ��r�A�ñN��r�˦��]�w������v�A��l�{���X�p�U�G
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
�ϥ� NpoiExcelHelper ��²�᪺�{���G
```csharp
var sheet = workbook.CreateSheet().Helper();
sheet[0][1].SetValue("Text").Style().Bold().Render();
```
�]�p�S��
=================
* �B�� Indexer ���o��ƦC�B�x�s��
* �ĥ� Fluent Api �]�p�A���{���X��[��ı
* ²�Ƽ˦��]�w��k�A�ðѦ�CSS�ݩʳ]�p�A�����}�o�̥i��֤W��

�Ҳշ��z
=================

1. SheetHelper: �� Sheet ���󪺦U�إ\��A�å[�JIndexer�H����/���o��ƦC(RowHelper)�C
2. RowHelper: �� Row ���󪺦U�إ\��A�å[�JIndexer�H����/���o�x�s��(CellHelper)�C
3. CellHelper: �� Cell ���󪺦U�إ\��A�å[�J²�ƪ������Ū��/�g�J�B�˦��]�w��k�C
4. StyleHelper: �˦��޲z����A�i���������gCSS�覡�]�w�~�[�A�ñN�˦��M�Ω�CellHelper�C


�t�λݨD
=============
NPOI 2.4.1.0  
.Net Framework 4.5 �H�W����

