﻿using System;
using System.Data;
using System.IO;
using System.Data.SqlClient;
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
            
            /*
             Suppress Excel Alerts
            https://docs.microsoft.com/en-us/office/vba/api/excel.application.displayalerts
            */
            xlApp.DisplayAlerts = false;
            //Call the Macro Test1
            xlApp.Run(macroName);

            xlWorkBook.Save();
            xlApp.Quit();
            return true;
        }

        public string SetActiveCell(string path, string fileName, string worksheetName, string xCordinate, string yCordinate)
        //x and y coordinates
        {

             //Selecting and Activating Cells
             //https://docs.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/selecting-and-activating-cells

            Application xlApp = new Application();
            //Suppress Excel Alerts
            xlApp.DisplayAlerts = false;
            Workbook xlWorkBook;

            string pathFileName = Path.Combine(path, fileName);

            //Open the Workbook.
            xlWorkBook = xlApp.Workbooks.Open(pathFileName);

            //https://social.msdn.microsoft.com/Forums/vstudio/en-US/02419ea7-1666-461e-b9f2-445d82e66322/c-with-excel-how-to-select-a-sheet?forum=vsto
            Worksheet workSheet = xlWorkBook.Sheets[worksheetName] as Worksheet;

            //https://www.syncfusion.com/kb/4220/how-to-set-an-active-cell-in-a-worksheet

            workSheet.Range[xCordinate + yCordinate].Activate();
            string activeCellValue = xlApp.ActiveCell.Value.ToString();

            xlWorkBook.Save();
            xlApp.Quit();
                        
            return activeCellValue;
        }

        public void PasValueToMacro(string path, string fileName, string macroName, string macroText)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook;

            string pathFileName = Path.Combine(path, fileName);

            //Open the Workbook.
            xlWorkBook = xlApp.Workbooks.Open(pathFileName);

            //Suppress Excel Alerts
            //https://docs.microsoft.com/en-us/office/vba/api/excel.application.displayalerts
            xlApp.DisplayAlerts = false;

            //Call the Macro MacroWithParameter
            var x = xlApp.Run(macroName, macroText);

            xlWorkBook.Save();
            xlApp.Quit();

        }

        public string GetCallValue(string path, string fileName, string worksheetName, string cellCordinate)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook;

            string pathFileName = Path.Combine(path, fileName);

            //Open the Workbook.
            xlWorkBook = xlApp.Workbooks.Open(pathFileName);

            //Suppress Excel Alerts
            //https://docs.microsoft.com/en-us/office/vba/api/excel.application.displayalerts
            xlApp.DisplayAlerts = false;

            Worksheet workSheet = xlWorkBook.Sheets[worksheetName] as Worksheet;
            //
            var cellValue = workSheet.Range[cellCordinate].Value2;
            return cellValue.ToString();
        }

        //public static DataSet SelectSQLRows(string connectionString, string queryString, string tableName)
        //{
        //    using (SqlConnection connection = new SqlConnection(connectionString))
        //    {
        //        SqlDataAdapter adapter  = new SqlDataAdapter();
        //    }
        //}
    }

}
