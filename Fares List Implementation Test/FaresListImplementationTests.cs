using NUnit.Framework;
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
    public void CanOpenExcelWorkBook()
    {
        var fileFunction = new FileFunctions();
        string fileName = @"c:\temp\FaresListImplementation.xlsx";

        Excel.Workbook wb = fileFunction.OpenExcelWorkBook(fileName);

            Assert.IsTrue(wb.GetType().ToString() == "Microsoft.Office.Interop.Excel.ApplicationClass");
    }
    }
}
