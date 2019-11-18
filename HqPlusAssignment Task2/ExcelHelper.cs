using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace HqPlusAssignment_Task2
{
    public class ExcelHelper
    {
        public static void CreateFormattedHotelRatesExcelFile(string filePath,List<HotelRate> hotelRates)
        {
            Application xlApp = new Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not installed in the system...");
                return;
            }

            object misValue = System.Reflection.Missing.Value;
            Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
            Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);


            // Filling the data in the Excel Worksheet
            xlWorkSheet.Cells[1, 1] = "ARRIVAL_DATE";
            xlWorkSheet.Cells[1, 2] = "DEPARTURE_DATE";
            xlWorkSheet.Cells[1, 3] = "PRICE";
            xlWorkSheet.Cells[1, 4] = "CURRENCY";
            xlWorkSheet.Cells[1, 5] = "RATENAME";
            xlWorkSheet.Cells[1, 6] = "ADULTS";
            xlWorkSheet.Cells[1, 7] = "BREAKFAST_INCLUDED";


            for (int i = 0; i < hotelRates.Count; i++)
            {
                xlWorkSheet.Cells[i+2, 1] = hotelRates[i].ArrivalDate;
                xlWorkSheet.Cells[i+2, 2] = hotelRates[i].DepartureDate;
                xlWorkSheet.Cells[i+2, 3] = hotelRates[i].Price;
                xlWorkSheet.Cells[i+2, 4] = hotelRates[i].Currency;
                xlWorkSheet.Cells[i+2, 5] = hotelRates[i].RateName;
                xlWorkSheet.Cells[i+2, 6] = hotelRates[i].Adults;
                xlWorkSheet.Cells[i+2, 7] = Convert.ToInt32(hotelRates[i].BreakfastIncluded);
            }

            // Formating Excel File (look and feel)
            xlWorkSheet.Cells.Font.Color = Color.FromArgb(45, 96, 144);
            xlWorkSheet.Cells.Font.Name = "Arial";
            
            //Cell Alignment
            xlWorkSheet.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            xlWorkSheet.get_Range("D:E").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            
            // Header Alignment
            Microsoft.Office.Interop.Excel.Range columnHeadingsRange = xlWorkSheet.Range["A1:G1"];
            columnHeadingsRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            columnHeadingsRange.AutoFilter2(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

            // alternative rows color.
            for (int i = 2; i <= hotelRates.Count+1; i++)
            {
                if (i % 2 == 0)
                {
                    var currentColor = Color.FromArgb(219, 230, 240);
                    xlWorkSheet.Range["A" + i + ":G" + i].Interior.Color = currentColor;
                }
            }

            xlWorkSheet.Columns.AutoFit();

            // columns data formatting
            var columnA = xlWorkSheet.get_Range("A1").EntireColumn;
            columnA.NumberFormat = "dd.MM.yy";

            var columnB = xlWorkSheet.get_Range("B1").EntireColumn;
            columnB.NumberFormat = "dd.MM.yy";

            var columnC = xlWorkSheet.get_Range("C1").EntireColumn;
            columnC.NumberFormat = "0.00";

            // Saving the file
            xlWorkBook.SaveAs(filePath, XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue,
                XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            xlWorkBook.Close(null,null,null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;
            GC.Collect();
        }
    }
}