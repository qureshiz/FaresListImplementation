﻿using System;
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
           string fileName = @"c:\temp\FSITestFileXLS.xlsx";

           Workbook workBook = fileFunction.OpenExcelWorkBook(fileName);

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
    }
}
