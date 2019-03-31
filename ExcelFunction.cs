using System;
using _Excel = Microsoft.Office.Interop.Excel; // Require for Excel.Application instantiation.  This creates an object Excel.
//Should we also have a using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace FaresListImplementation
{
    public class FileFunctions
    {

        public CreateExcelApplication()
        {
            var excelApp = new Excel.Application();

            // Quit Excel Application.
            excelApp.Quit();

            // Clean Up.
            releaseObject(xlApp);
            releaseObject(xlWorkBook);
            //test
        }

        public bool isthefilethere()
        {
            return CheckFileExists("");
        }

        class Excel
        {
            string path = ""; //Path to Excel file.
            _Application excel = new _Excel.Application();
            workbook wb;
            worksheet ws;

            public Excel(string path, int sheet)
            {
                this.path = path;
                wb = excel.Workbooks.Open(path);
                ws = wb.Worksheets[sheet]; // Not sure if this is a square or curly bracket.
            }

            public string ReadCell(int i, int j) // row and column reference.
            {
                //Note, excel reference is 1 based.
                i++;
                j++;
                //ensure cell is not null.
                if (ws.Cells[i, j].value2 != null) // value2 is the value inside the cell.
                    return ws.cells[i, j].value2;
                else
                    return "";

            }
        }

    }


}
