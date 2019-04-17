//using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;
//using Interop.Microfoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using System;
using System.Configuration;
namespace FaresListImplementation.Tests
{
    class Global
    {
        public static string testFileLocation = ConfigurationManager.AppSettings["TestFileLocation"];
        public static string excelTestFile = "FSITestFileXLS.xlsm";
    }
    [TestFixture]
    public class FileTests
        
    {
        [Test]
        public void TestFileExists()
        {
            FileFunctions fileFunction = new FileFunctions();

            //appSettings.
            //https://www.c-sharpcorner.com/article/four-ways-to-read-configuration-setting-in-c-sharp/
            
            Assert.IsTrue(fileFunction.CheckFileExists(string.Concat(Global.testFileLocation,Global.excelTestFile)));
        }
    [Test]
        public void CanOpenFSITestFileXLS()
        {
           FileFunctions fileFunction = new FileFunctions();
            
           Workbook workBook = fileFunction.OpenExcelWorkBook(string.Concat(Global.testFileLocation,Global.excelTestFile));

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
        public void CanGetActiveCellValue()
        {
            FileFunctions filefunctions = new FileFunctions();

            string activeCell = filefunctions.SetActiveCell(Global.testFileLocation, Global.excelTestFile, "sheet1", "A", "2");
            Assert.AreEqual("This is the Active Cell",activeCell);

        }
        [Test]
        public void CanPassValueToMacro()
        {
            FileFunctions filefunctions = new FileFunctions();

            string macroName = "Test2MacroWithParameter";
            string dateValue  = DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss"
            string cellValue = filefunctions.PasValueToMacro(
                            Global.testFileLocation, 
                            Global.excelTestFile, 
                            macroName, 
                            DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss")
                            );

            string testValue = "I'm a parameter";

            Assert.AreEqual(testValue, "actual");

        }
    }
}
