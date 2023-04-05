using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Extensions.Rafiq;
using System.Net.Http.Headers;
using FxForwardExtract.Models;
using System.Security.Cryptography;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace FxForwardExtract.Modules
{
    public static class Module_Extract
    {
        public static async Task<Excel.Worksheet> OpenFile(string theFileName)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbooks workbooks = xlApp.Workbooks;
            Excel.Workbook workbook = workbooks.Open(theFileName, ReadOnly: true);
            Excel.Worksheet worksheet = workbook.Sheets[1];

            return worksheet;
        }

        public static async Task<List<List<object>>> ExtractColumns(Excel.Worksheet worksheet)
        {
            List<List<object>> listSpreadSheet = new List<List<object>>();

            for (int i = await "A".GetColumnFromAlphabet(); i <= await "O".GetColumnFromAlphabet(); i++)
            {
                object[,] theColumn = worksheet.Columns[i].Value;

                List<object> theColumnFlat = theColumn
                    .Flatten<object>()
                    .ToList();

                listSpreadSheet.Add(theColumnFlat);

                Console.WriteLine($"{i} of {await "O".GetColumnFromAlphabet()} Extracted\r\n");
            }

            return listSpreadSheet;
        }

        public static async Task<List<FundBoundary>> FindBoundaries(List<List<object>> listSpreadSheet)
        {
            List<FundBoundary> listBoundaries = new List<FundBoundary>();

            int lastLastLine = listSpreadSheet[2]
                .FindIndex(x => x?.ToString() == "Grand Total");

            int startIndex = 0;

            while (startIndex != -1)
            {
                int beginIndex = listSpreadSheet[0]
                    .FindIndex(startIndex, x => x?.ToString() == "FX Forward Holding");


                int endIndex = listSpreadSheet[2]
                    .FindIndex(startIndex, x => x?.ToString() == "Total");

                if (endIndex != -1)
                {
                    FundBoundary theBoundary = new FundBoundary()
                    {
                        StartLine = beginIndex,
                        EndLine = endIndex,
                        PFName = listSpreadSheet[0][beginIndex + 3]?
                        .ToString()
                        .Replace("Portfolio Name:", "")
                        .Trim() ?? "",
                        PFPlanCode = listSpreadSheet[0][beginIndex + 2]?
                        .ToString()
                        .Replace("Portfolio Code:", "")
                        .Trim() ?? "",
                        DateAsOf = DateTime.ParseExact(listSpreadSheet[1][beginIndex + 4]?.ToString(), "yyyy/MM/dd", new CultureInfo("EN-MY"))
                    };

                    startIndex = endIndex + 1;
                    listBoundaries.Add(theBoundary);

                    //Console.WriteLine(theBoundary.StartLine);
                    Console.WriteLine(theBoundary.PFName);
                    //Console.WriteLine(theBoundary.EndLine);
                }
                else
                {
                    startIndex = endIndex;
                }

                Console.WriteLine("");
            }

            return listBoundaries;
        }

        public static async Task<List<FXForwardData>> ExtractForwardData(List<FundBoundary> listBoundaries, List<List<object>> listSpreadSheet)
        {
            List<FXForwardData> listFxForward = new List<FXForwardData>();

            PropertyInfo[] propertyInfos = typeof(FXForwardData).GetProperties();


            foreach (FundBoundary boundary in listBoundaries)
            {
                int startLine = boundary.StartLine + 6;

                while (startLine < boundary.EndLine)
                {
                    if (listSpreadSheet[0][startLine] != null)
                    {
                        FXForwardData theData = new FXForwardData()
                        {
                            PFName = boundary.PFName,
                            PFCode = boundary.PFPlanCode,
                            DateAsOf = boundary.DateAsOf
                        };

                        for (int i = await "A".GetColumnFromAlphabet() - 1; i < await "O".GetColumnFromAlphabet(); i++)
                        {
                            propertyInfos[i + 3].SetValue(theData, listSpreadSheet[i][startLine]);
                        }

                        listFxForward.Add(theData);
                    }

                    startLine++;
                }
            }

            return listFxForward;
        }
    }
}
