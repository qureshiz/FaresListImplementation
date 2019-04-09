using System;
using System.IO;
//using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

//Should we also have a using Microsoft.Office.Interop.Excel?

namespace FaresListImplementation
{
    public class FileFunctions
    {
        public bool CheckFileExists(string path)
        {
            //Hard code the name of a file for now.

            return File.Exists(path);
        }

        public Workbook OpenExcelWorkBook(string fileName)
        {

            // Define Excel objects.
            //Excel.Application xlApp = new Excel.Application();
            Application xlApp = new Application();
            //Excel.Workbook xlWorkBook;
            Workbook xlWorkBook;

            //Open the Workbook.
            xlWorkBook = xlApp.Workbooks.Open(fileName);
            
                return xlWorkBook;
 
        }

        public int WorkSheetCount(string fileName)
        {
            // Define Excel objects.
            //Excel.Application xlApp = new Excel.Application();
            Application xlApp = new Application();
            //Excel.Workbook xlWorkBook;
            Workbook xlWorkBook;

            //Open the Workbook.
            xlWorkBook = xlApp.Workbooks.Open(fileName);

            int workSheetCount = xlWorkBook.Sheets.Count;
            xlWorkBook.Close(false);
            xlApp.Quit();

            return workSheetCount;

        }

        public bool RunExcelMacro(string path, string fileName, string macroName)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook;

            //Call the Macro "Test1"
            string pathFileName = Path.Combine(path, fileName);
            
            //Open the Workbook.
            xlWorkBook = xlApp.Workbooks.Open(pathFileName);
            xlApp.DisplayAlerts = false;
            //Call the Macro Test1
            xlApp.Run(macroName);

            xlWorkBook.Save();
            xlApp.Quit();
            return true;
        }

    }

}
