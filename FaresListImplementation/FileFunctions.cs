using System;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

//Should we also have a using Microsoft.Office.Interop.Excel?

namespace FaresListImplementation
{
    public class FileFunctions
    {
        public bool CheckFileExists(string pathFileName)
        {
            //Hard code the name of a file for now.

            return File.Exists(pathFileName);
        }

        public Workbook OpenExcelWorkBook(string pathFileName)
        {

            // Define Excel objects.
            Application xlApp = new Application();
            //Excel.Workbook xlWorkBook;
            Workbook xlWorkBook;

            //Open the Workbook.
            xlWorkBook = xlApp.Workbooks.Open(pathFileName);
            
                return xlWorkBook;
         }

        public int WorkSheetCount(string fileName)
        {
            // Define Excel objects.
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

        public bool RunExcelMacro(string path, string fileName, string macroName, bool save)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook;

            //Call the Macro "Test1"
            string pathFileName = Path.Combine(path, fileName);
            
            //Open the Workbook.
            xlWorkBook = xlApp.Workbooks.Open(pathFileName);
            
            /*
             Suppress Excel Alerts
            https://docs.microsoft.com/en-us/office/vba/api/excel.application.displayalerts
            */
            xlApp.DisplayAlerts = false;
            //Call the Macro Test1
            xlApp.Run(macroName);
            if (save==true)
            {
                xlWorkBook.Save();
            }
              
           
            xlApp.Quit();
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        /// <param name="worksheetName"></param>
        /// <param name="xCordinate"></param>
        /// <param name="yCordinate"></param>
        public void SetActiveCell(string path, string fileName, string worksheetName, string xCordinate, string yCordinate)
        //x and y coordinates
        {

             //Selecting and Activating Cells
             //https://docs.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/selecting-and-activating-cells

            Application xlApp = new Application();
            xlApp.DisplayAlerts = false;
            Workbook xlWorkBook = xlApp.Workbooks.Open(Path.Combine(path, fileName));

            //Open the Workbook.
            //xlWorkBook = xlApp.Workbooks.Open(Path.Combine(path, fileName));

            //https://social.msdn.microsoft.com/Forums/vstudio/en-US/02419ea7-1666-461e-b9f2-445d82e66322/c-with-excel-how-to-select-a-sheet?forum=vsto
            Worksheet workSheet = xlWorkBook.Sheets[worksheetName] as Worksheet;

            //https://www.syncfusion.com/kb/4220/how-to-set-an-active-cell-in-a-worksheet

            workSheet.Range[xCordinate + yCordinate].Activate();

            xlWorkBook.Save();
            xlApp.Quit();
                        
            //return activeCellValue;
        }

        public void SetActiveCell(Worksheet worksheet, string xCordinate, string yCordinate)
        {
            worksheet.Range[xCordinate+yCordinate, xCordinate+yCordinate].Activate();
        }

        public string GetActiveCellValue(Application application)
        {
            return application.ActiveCell.Value.ToString();
        }

        public void PasValueToMacro(string path, string fileName, string macroName, string macroText)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook;

            //Open the Workbook.
            xlWorkBook = xlApp.Workbooks.Open(Path.Combine(path, fileName));

            //Suppress Excel Alerts
            //https://docs.microsoft.com/en-us/office/vba/api/excel.application.displayalerts
            xlApp.DisplayAlerts = false;

            //Call the Macro MacroWithParameter
            xlApp.Run(macroName, macroText);

            xlWorkBook.Save();
            xlApp.Quit();

        }

        public string GetCellValue(string path, string fileName, string worksheetName, string cellCordinate)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook;

            //Open the Workbook.
            xlWorkBook = xlApp.Workbooks.Open(Path.Combine(path, fileName));

            //Suppress Excel Alerts
            //https://docs.microsoft.com/en-us/office/vba/api/excel.application.displayalerts
            xlApp.DisplayAlerts = false;

            Worksheet workSheet = xlWorkBook.Sheets[worksheetName] as Worksheet;
            //
            var cellValue = workSheet.Range[cellCordinate].Value2;
            return cellValue.ToString();
        }

    }
    
    public class cExcel
    {
       
        private static Workbook _Workbook;

        static void main(string path, string fileName)
        {
            Application xlApp = new Application();
            _Workbook = xlApp.Workbooks.Open(Path.Combine(path, fileName));
        }
        public Workbook workBook
            {
            get
            {
                return this._Workbook;
            }

}

    }

}
