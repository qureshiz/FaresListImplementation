//using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;
//using Interop.Microfoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using System;
using System.Configuration;
using System.IO;
namespace FaresListImplementation.Tests
{
    class Global
    {
        public static string testFileLocation = ConfigurationManager.AppSettings["TestFileLocation"];
        public static string excelTestFile = ConfigurationManager.AppSettings["TestExcelFile"];
        public static string siteLists_xlxs = ConfigurationManager.AppSettings["SiteLists_xlsx"];
    }
    [TestFixture]
    public class FileTests
        
    {
        [Test]
        public void TestFileExists()
        {
            FileFunctions fileFunction = new FileFunctions();

            Assert.IsTrue(fileFunction.CheckFileExists(Path.Combine(Global.testFileLocation, Global.excelTestFile)));
        }
    [Test]
        public void CanOpenFSITestFileXLS()
        {
           FileFunctions fileFunction = new FileFunctions();
           Workbook workBook = fileFunction.OpenExcelWorkBook(Path.Combine("Global.testFileLocation", "Global.excelTestFile"));

           Assert.IsTrue(workBook.GetType().ToString() == "Microsoft.Office.Interop.Excel.WorkbookClass");

            if (workBook !=null)
            {
                //workBook.Close();
                workBook.Application.Quit();
            }
        }
        
        // Get Worksheet Count
        [Test]
        public void GetWorksheetCount()
        {
            //Arrange
           FileFunctions fileFunction = new FileFunctions();
           string fileName = @"c:\temp\FSITestFileXLS.xlsx";
           int workSheetCount = fileFunction.WorkSheetCount(fileName);

            Assert.Greater(workSheetCount, 0);
        }

        [Test]
        public void CanCallExcelTestMacro()
        {
            //Created Macro Test1 in "C:\temp\FSITestFileXLS.xlsm"
            //See https://social.msdn.microsoft.com/Forums/lync/en-US/2e33b8e5-c9fd-42a1-8d67-3d61d2cedc1c/how-to-call-excel-macros-programmatically-in-c?forum=exceldev
            FileFunctions fileFunctions   = new FileFunctions();
            string macroName = "Test1";

            Assert.IsTrue(fileFunctions.RunExcelMacro(Global.testFileLocation, Global.excelTestFile, macroName));
        }
        [Test]
        public void CanSetActiveCellValue()
        {
            FileFunctions filefunctions = new FileFunctions();

            Workbook xlWorkBook = filefunctions.OpenExcelWorkBook(Path.Combine(Global.testFileLocation, Global.excelTestFile));
            Worksheet workSheet = xlWorkBook.Sheets["Sheet1"] as Worksheet;
            filefunctions.SetActiveCell(workSheet, "A", "2");
            string activeCellValue = filefunctions.GetActiveCellValue(workSheet.Application);
            Assert.AreEqual("This is the Active Cell",activeCellValue);
            xlWorkBook.Close();
            xlWorkBook.Application.Quit();

        }
        [Test]
        public void CanPassValueToMacro()
        {
            FileFunctions filefunctions = new FileFunctions();

            string macroName = "Test2MacroWithParameter";
            string passValue = "zishan";
            //Macro "Test2MacroWithParameter" writes to Cell C1.
            filefunctions.PasValueToMacro(Global.testFileLocation,Global.excelTestFile,macroName, passValue);

            string cellValue = filefunctions.GetCellValue(Global.testFileLocation, Global.excelTestFile, "Sheet1", "C1");

            Assert.AreEqual(passValue, cellValue);
        }

        [Test]
        public void CanGetCellValue()
        {
            FileFunctions filefunctions = new FileFunctions();
            string cellValue = filefunctions.GetCellValue(Global.testFileLocation, Global.excelTestFile, "Sheet1","a2");
            Assert.AreEqual("This is the Active Cell", cellValue);
        }

        [Test]
        public void SiteLists_xlsxExists()
        {
            FileFunctions filefunctions = new FileFunctions();
            string pathFileName = Path.Combine(Global.testFileLocation, Global.siteLists_xlxs);

            Assert.IsTrue(filefunctions.CheckFileExists(pathFileName));
        }
        [Test]
        public void CanOpenSiteLists_xlsx()
        {
            FileFunctions filefunctions = new FileFunctions();
            Workbook workBook = filefunctions.OpenExcelWorkBook(Path.Combine(Global.testFileLocation, Global.siteLists_xlxs));
            Assert.IsTrue(workBook.GetType().ToString() == "Microsoft.Office.Interop.Excel.WorkbookClass");
            if (workBook != null)
            {
                //workBook.Close();
                workBook.Application.Quit();
            }
        }
    }

}
