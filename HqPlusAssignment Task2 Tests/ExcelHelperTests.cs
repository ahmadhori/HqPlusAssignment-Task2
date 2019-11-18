using HqPlusAssignment_Task2;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HqPlusAssignment_Task2_Tests
{
    public class ExcelHelperTests
    {
        [SetUp]
        public void Setup()
        {
        }




        [TestCase("task 2 - hotelrates.json")]
        public void Should_CreateFormattedHotelRatesExcelFile_When_JsonFile_Is_Valid(string jsonFileName)
        {
            string text = File.ReadAllText(jsonFileName);

            List<HotelRate> hotelRates = JsonHelper.ParseHotelRatesFile(text).ToList();

            File.Delete(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + @"\OutputExcelFile-Task2.xlsx");

            ExcelHelper.CreateFormattedHotelRatesExcelFile(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + @"\OutputExcelFile-Task2.xlsx", hotelRates.ToList());

            Assert.IsTrue(File.Exists(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + @"\OutputExcelFile-Task2.xlsx"));
        }
    }
}