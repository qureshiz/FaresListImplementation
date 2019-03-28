using System;
using Excel = Microsoft.Office.Interop.Excel; // Require for Excel.Application instantiation.


public class ExcelFunction
{
	public CreateExcelApplication()
	{
	    var excelApp = new Excel.Application();
	    
	    // Quit Excel Application.
	    excelApp.Quit();

	    // Clean Up.
	    releaseObject(xlApp);
	    releaseObject(xlWorkBook);
    }
}
