using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;



namespace MlabToExcelExport
{
    public static class ExportToExcel
    {
        private static Excel.Application CreateExcelObj()
        {
            object obj;
            obj = null;
            try
            {
                //Создаём приложение.
                Excel.Application objExcel = new Excel.Application();
                objExcel.Workbooks.Add();
                obj = objExcel;

            }
            catch (Exception ex)
            {
                throw new Exception("Ошибка создания экземпляра MS Excel");
            }
            return (obj as Excel.Application);
        }

        private static void FormatSheetForSet(Excel.Worksheet sheet, SetItem obj)
        {

            // formatting All sheet
            sheet.PageSetup.PrintGridlines = false;
            sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
            sheet.PageSetup.RightFooter = "Дата: &DD Стр &PP из &NN";
            sheet.PageSetup.RightHeader = "Исследование " + obj.Project + ", сет № " + obj.Set + " - " + obj.TestMethod +
                                          " - " + obj.AB;
            sheet.PageSetup.Zoom = false;
            sheet.PageSetup.LeftHeader = "НИИ Антимикробной химиотерапии";
            sheet.PageSetup.TopMargin = 50;
            sheet.PageSetup.BottomMargin = 50;
            sheet.PageSetup.HeaderMargin = 20;
            sheet.PageSetup.FooterMargin = 20;
            sheet.PageSetup.RightMargin = 10;
            sheet.PageSetup.LeftMargin = 50;
            sheet.PageSetup.Order = Excel.XlOrder.xlOverThenDown;

            // Foramatting test method
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, obj.MICList.Count + 5]].Merge();
            FormatHeaderText(sheet.Range[sheet.Cells[1, 1], sheet.Cells[3, 1]]);

            // Formatting Set Number
            sheet.Range[sheet.Cells[3, 1], sheet.Cells[3, 4]].Merge();
            FormatHeaderText(sheet.Range[sheet.Cells[3, 1], sheet.Cells[3, 1]]);

            // Formatting Set Number
            sheet.Range[sheet.Cells[3, 5], sheet.Cells[3, obj.MICList.Count + 5]].Merge();
            FormatHeaderText(sheet.Range[sheet.Cells[3, 5], sheet.Cells[3, obj.MICList.Count + 5]]);

            //Formatting table with MO
            FormatTableCells(sheet.Range[sheet.Cells[5, 1], sheet.Cells[5 + obj.MOList.Count, obj.MICList.Count + 5]]);

            //Formatting Control MO Header
            sheet.Range[sheet.Cells[6 + obj.MOList.Count, 1], sheet.Cells[6 + obj.MOList.Count, obj.MICList.Count + 5]].Merge();
            sheet.Range[sheet.Cells[6 + obj.MOList.Count, 1], sheet.Cells[6 + obj.MOList.Count, obj.MICList.Count + 5]].RowHeight = 15;
            FormatHeaderControlMOText(
                sheet.Range[
                    sheet.Cells[6 + obj.MOList.Count, 1], sheet.Cells[6 + obj.MOList.Count, obj.MICList.Count + 5]]);

            // Formatting table with control MO
            FormatTableCells(sheet.Range[sheet.Cells[5 + obj.MOList.Count + 1, 1], sheet.Cells[5 + obj.MOList.Count + 1 + obj.ControlMOList.Count, obj.MICList.Count + 5]]);
            FormatHeaderControlMOText(sheet.Range[
                    sheet.Cells[7 + obj.MOList.Count, 2], sheet.Cells[6 + obj.MOList.Count + obj.ControlMOList.Count, 4]]);
            //Formatting Top Row
            sheet.Range[sheet.Cells[5, 1], sheet.Cells[5, obj.MICList.Count + 5]].ColumnWidth = 6;
            //Formatting Left Columns
            sheet.Range[sheet.Cells[5, 1], sheet.Cells[5 + obj.MOList.Count, 1]].ColumnWidth = 6;
            sheet.Range[sheet.Cells[5, 2], sheet.Cells[5 + obj.MOList.Count, 2]].ColumnWidth = 8;
            sheet.Range[sheet.Cells[5, 3], sheet.Cells[5 + obj.MOList.Count, 3]].ColumnWidth = 8;
            sheet.Range[sheet.Cells[5, 4], sheet.Cells[5 + obj.MOList.Count, 4]].ColumnWidth = 14;
            //Formatting Right Columns
            sheet.Range[sheet.Cells[5, obj.MICList.Count + 5], sheet.Cells[5 + obj.MOList.Count, obj.MICList.Count + 5]].ColumnWidth = 8;


            sheet.Cells[obj.MOList.Count + obj.ControlMOList.Count + 10] = "Проверил:";
        }

        private static void FormatHeaderText(Excel.Range range)
        {
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.RowHeight = 18;
            range.Font.Size = 10;
            range.Font.Bold = true;
        }

        private static void FormatHeaderControlMOText(Excel.Range range)
        {
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            range.Font.Size = 10;
            range.Font.Bold = true;
            range.Font.Italic = true;
            range.Interior.ColorIndex = 34;
        }

        private static void FormatTableCells(Excel.Range range)
        {
            range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Font.Size = 10;
            range.Font.Bold = true;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.WrapText = true;
            range.RowHeight = 25;
            range.ColumnWidth = 6;
        }

        public static void GetExcelDocumentSet(SetViewModel obj)
        {

            Excel.Application ExcelApp = CreateExcelObj();
            ExcelApp.ScreenUpdating = false;
            Excel.Workbook ExcelWorkbook = ExcelApp.ActiveWorkbook;
           
            try
            {
                foreach (var itemSet in obj.Set)
                {
                    Excel.Worksheet ExcelSheet = ExcelWorkbook.Sheets.Add();

                    var rowsCount = itemSet.MOList.Count + 8 + itemSet.ControlMOList.Count + 1;
                    var columnsCount = itemSet.MICList.Count + 5;

                    Excel.Range ExcelRange =
                        ExcelSheet.Range[ExcelSheet.Cells[1, 1], ExcelSheet.Cells[rowsCount, columnsCount]];

                    ExcelSheet.Name = itemSet.AB;


                    var data = PrepareListForSet(itemSet);


                    ExcelRange.Value = data;
                    FormatSheetForSet(ExcelSheet, itemSet);

                }

                ExcelWorkbook.Sheets[obj.Set.Count + 1].Delete();
                ExcelWorkbook.Sheets[obj.Set.Count + 1].Delete();
                ExcelWorkbook.Sheets[obj.Set.Count + 1].Delete();

                ExcelApp.ScreenUpdating = true;
                ExcelApp.Visible = true;
            }
            catch (Exception ex)
            {

            }
            finally
            {
                ExcelWorkbook = null;
                ExcelApp = null;
                GC.Collect();
            }
        }

        private static object[,] PrepareListForSet(SetItem obj)
        {
            // Create data array (using array for data export optimization)
            var rowsCount = obj.MOList.Count + 8 + obj.ControlMOList.Count + 1;
            var columnsCount = obj.MICList.Count + 5;
            object[,] data = new object[rowsCount, columnsCount];

            data[0, 0] = "Метод тестирования: " + obj.TestMethod;
            data[2, 0] = "Сет № " + obj.Set;
            data[2, 4] = obj.AB;

            data[4, 0] = "Ячейка";
            data[4, 1] = "№";
            data[4, 2] = "Муз. №.";
            data[4, 3] = "МО";

            for (int i = 0; i < obj.MICList.Count; i++)
            {
                data[4, 4 + i] = obj.MICList[i];
            }

            data[4, 4 + obj.MICList.Count] = "МПК";

            for (int i = 0; i < obj.MOList.Count; i++)
            {
                data[5 + i, 0] = obj.MOList[i].Cell;
                data[5 + i, 1] = obj.MOList[i].Number;
                data[5 + i, 2] = obj.MOList[i].MuseumNumber;
                data[5 + i, 3] = obj.MOList[i].MO;

            }

            data[5 + obj.MOList.Count, 0] = "Контрольн.МО";

            for (int i = 0; i < obj.ControlMOList.Count; i++)
            {
                data[6 + obj.MOList.Count + i, 0] = obj.ControlMOList[i].Cell;
                data[6 + obj.MOList.Count + i, 1] = obj.ControlMOList[i].Number;
                data[6 + obj.MOList.Count + i, 2] = obj.ControlMOList[i].MuseumNumber;
                data[6 + obj.MOList.Count + i, 3] = obj.ControlMOList[i].MO;
            }

            return data;
        }
    }
}
