using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
//using _Excel = Microsoft.Office.Interop.Excel; // Require for Excel.Application instantiation.  This creates an object Excel.
//Should we also have a using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;

namespace FaresListImplementation
{
    public class FileFunctions
    {
        public bool CheckFileExists(string path)
        {
            //Hard code the name of a file for now.

            return File.Exists(path);
        }

        public object ExcelApplication()
        {
            Excel.Application xlApp = new Excel.Application();
            return xlApp;
        }
    }

}
