using Excel = Microsoft.Office.Interop.Excel;
using FxForwardExtract.Modules;
using FxForwardExtract.Models;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

//string fileName = "C:\\172.28.151.165\\nas03_apps\\UOBAM_MY\\Operation\\Valuation\\" +
//    "06 - DBFA Valuation\\DAILY\\YEAR 2023\\03. Mar 2023\\16032023\\16032023 (T1-4pm)\\" +
//    "MY_UOB_Reports_16032023\\49200444FX_Forward_Holding.xls";

string fileName = "C:\\Users\\nursara\\source\\repos\\fxforward-basics\\49200444FX_Forward_Holding.xls";

Excel.Worksheet theWorksheet = await Module_Extract.OpenFile(fileName);

List<List<object>> listSpreadSheet = await Module_Extract.ExtractColumns(theWorksheet);

Excel.Workbooks workbooks = theWorksheet.Application.Workbooks;
Excel.Workbook workbook = theWorksheet.Application.Workbooks[1];
Excel.Application xlApp = theWorksheet.Application;

xlApp.Visible = true;
workbook.Close(0, null, null);

xlApp.Quit();

Marshal.ReleaseComObject(theWorksheet);
Marshal.ReleaseComObject(workbook);
Marshal.ReleaseComObject(workbooks);
Marshal.ReleaseComObject(xlApp);

List<FundBoundary> listBoundaries = await Module_Extract.FindBoundaries(listSpreadSheet);

List<FXForwardData> listFxForward = await Module_Extract.ExtractForwardData(listBoundaries, listSpreadSheet);

await Module_Write.WriteToExcel(listFxForward);

