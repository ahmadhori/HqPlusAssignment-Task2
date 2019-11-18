using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop;

namespace HqPlusAssignment_Task2
{
    class Program
    {
        static void Main(string[] args)
        {
            string text = File.ReadAllText("task 2 - hotelrates.json");

            List<HotelRate> hotelRates = JsonHelper.ParseHotelRatesFile(text).ToList();
            
            File.Delete(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + @"\OutputExcelFile-Task2.xlsx");

            ExcelHelper.CreateFormattedHotelRatesExcelFile(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + @"\OutputExcelFile-Task2.xlsx", hotelRates.ToList());
        }
    }
}