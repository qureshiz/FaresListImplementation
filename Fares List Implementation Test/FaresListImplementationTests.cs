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
        public void CanOpenExcelWorkBook()
        {
           var fileFunction = new FileFunctions();
           string fileName = @"c:\temp\FaresListImplementation.xlsx";

            var openedWorkBook = fileFunction.OpenExcelWorkBook(fileName);
            Workbook workBook = fileFunction.OpenExcelWorkBook(fileName);

           Assert.IsTrue(openedWorkBook.GetType().ToString() == "Microsoft.Office.Interop.Excel.WorkbookClass");
    }
    }
}
