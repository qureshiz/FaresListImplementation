using System;
using NUnit.Framework;
using FaresListImplementation;
using Excel = Microsoft.Office.Interop.Excel;
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
    public void OpenExcelApplication()
    {
        var fileFunction = new FileFunctions();
        var excelObject = fileFunction.ExcelApplication();
        //Assert.IsInstanceOf(Excel.Application, excelObject);
    }
    }
}
