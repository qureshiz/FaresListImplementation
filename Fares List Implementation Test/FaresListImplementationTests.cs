using System;
using NUnit.Framework;
using FaresListImplementation;
//using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;
//using Interop.Microfoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace FaresListImplementation.Tests
{
    [TestFixture]
    public class FileTests
    {
        [Test]
        public void TestFileExists()
        {
           var fileFunction = new FileFunctions();
           string fileName = @"c:\temp\FaresListImplementation.xlsx";
           Assert.IsTrue(fileFunction.CheckFileExists(fileName));
        }
    [Test]
        public void CanOpenFSITestFileXLS()
        {
           FileFunctions fileFunction = new FileFunctions();
           string fileName = @"c:\temp\FSITestFileXLS.xlsm";

           Workbook workBook = fileFunction.OpenExcelWorkBook(fileName);

           Assert.IsTrue(workBook.GetType().ToString() == "Microsoft.Office.Interop.Excel.WorkbookClass");

            if (workBook !=null)
            {
                //workBook.Close();   
                workBook.Application.Quit();
                workBook.Close();   
                
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
            var fileFunctions   = new FileFunctions();
            //Test that a Macro can be run
            string path         = @"c:\temp";
            string fileName     = "FSITestFileXLS.xlsm";
            string macroName = "Test1";

            Assert.IsTrue(fileFunctions.RunExcelMacro(path, fileName, macroName));

        }

    }
}
