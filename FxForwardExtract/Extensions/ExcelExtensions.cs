using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Extensions.Rafiq
{
    public static class ExtensionsExcel
    {
        public static async Task<string> FillExcelColumns<T>(this IEnumerable<T> data, Microsoft.Office.Interop.Excel.Worksheet worksheet, int rowStart, int columnStart)
        {
            string columnRange = await CalculateRange(rowStart, columnStart, data.Count());

            var columnData = data.ColumnisedData();

            worksheet.Range[columnRange].Value = columnData;

            return await CalculateRange(rowStart, columnStart, data.Count());
        }

        public static async Task<string> FillExcelColumns<T>(this IEnumerable<T> data, Microsoft.Office.Interop.Excel.Worksheet worksheet, string columnAlphabet, int rowStart)
        {
            int rowEnd = rowStart + data.Count() - 1;

            string columnRange = $"{columnAlphabet}{rowStart}:{columnAlphabet}{rowEnd}";

            var columnData = data.ColumnisedData();

            worksheet.Range[columnRange].Value = columnData;

            return columnRange;
        }

        public static async Task<string> FillExcelColumnsFormula<T>(this IEnumerable<T> data, Microsoft.Office.Interop.Excel.Worksheet worksheet, int rowStart, int columnStart)
        {
            string columnRange = await CalculateRange(rowStart, columnStart, data.Count());

            worksheet.Range[columnRange].FormulaArray = data.ToArray();

            return await CalculateRange(rowStart, columnStart, data.Count());
        }

        public static async Task<string> FillExcelColumnsFormula<T>(this IEnumerable<T> data, Microsoft.Office.Interop.Excel.Worksheet worksheet, string columnAlphabet, int rowStart)
        {
            int rowEnd = rowStart + data.Count() - 1;

            string columnRange = $"{columnAlphabet}{rowStart}:{columnAlphabet}{rowEnd}";

            var columnData = data.ColumnisedData();

            worksheet.Range[columnRange].FormulaArray = data.ToArray();

            return columnRange;
        }

        private static T[,] ColumnisedData<T>(this IEnumerable<T> data)
        {
            var dataList = data.ToList();

            T[,] returnValue = new T[dataList.Count, 1];

            for (int i = 0; i < data.Count(); i++)
            {
                returnValue[i, 0] = dataList[i];
            }

            return returnValue;
        }

        private static async Task<string> CalculateRange(int startRow, int startColumn, int dataLen)
        {
            int endRow = startRow + dataLen - 1;

            string columnName = "";

            int columnNumber = startColumn;

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return $"{columnName}{startRow}:{columnName}{endRow}";
        }

        public static async Task<int> GetColumnFromAlphabet(this string alphabets)
        {
            alphabets = alphabets.ToUpper();

            int theColumn = 0;

            for (int i = 0; i < alphabets.Length; i++)
            {
                theColumn += (int)Math.Pow(26, i) * (Convert.ToInt32(alphabets[alphabets.Length - i - 1]) - Convert.ToInt32('A') + 1);
            }

            return theColumn;
        }

        public static async Task<List<object>> ToDateTimeList<T>(this List<T> theDateList, string schemaDate = null)
        {
            List<object> result = new List<object>();

            foreach (var item in theDateList)
            {
                if (item == null || item.ToString() == "")
                {
                    result.Add(item);
                }
                else
                {
                    DateTime theDateTime = new DateTime();
                    bool parseResult = false;

                    if (schemaDate == null)
                    {
                        parseResult = DateTime.TryParse(item.ToString(), out theDateTime);
                    }
                    else
                    {
                        parseResult = DateTime.TryParseExact(item.ToString(), schemaDate, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out theDateTime);
                    }

                    if (parseResult)
                    {
                        result.Add(theDateTime);
                    }
                    else
                    {
                        result.Add(item);
                    }
                }
            }

            return result;
        }

        public static async Task<List<object>> ToDoubleList<T>(this List<T> theDoubleList)
        {
            List<object> result = new List<object>();

            foreach (var item in theDoubleList)
            {
                if (item == null || item.ToString() == "")
                {
                    result.Add(item);
                }
                else
                {
                    double theDoubleValue = 0;

                    bool success = double.TryParse(item.ToString(), out theDoubleValue);
                    if (success)
                    {
                        result.Add(theDoubleValue);
                    }
                    else
                    {
                        result.Add(item);
                    }
                }
            }

            return result;
        }

        public static async Task FormatExcelCell(this Microsoft.Office.Interop.Excel.Worksheet worksheet, string excelRange,
            string numberformat, Microsoft.Office.Interop.Excel.XlHAlign alignment,
            Microsoft.Office.Interop.Excel.XlVAlign vAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop)
        {
            worksheet.Range[excelRange].NumberFormat = numberformat;
            worksheet.Range[excelRange].HorizontalAlignment = alignment;
            worksheet.Range[excelRange].VerticalAlignment = vAlignment;
        }

        public static async Task FormatExcelCell(this Microsoft.Office.Interop.Excel.Worksheet worksheet, int row, int column,
            string numberformat, Microsoft.Office.Interop.Excel.XlHAlign alignment,
            Microsoft.Office.Interop.Excel.XlVAlign vAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop)
        {
            worksheet.Cells[row, column].NumberFormat = numberformat;
            worksheet.Cells[row, column].HorizontalAlignment = alignment;
            worksheet.Cells[row, column].VerticalAlignment = vAlignment;
        }

        public static async Task FormatExcelText(this Microsoft.Office.Interop.Excel.Worksheet worksheet, string excelRange, bool bold, int size = 11,
            Microsoft.Office.Interop.Excel.XlUnderlineStyle underlineStyle = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone)
        {
            worksheet.Range[excelRange].Font.Bold = bold;
            worksheet.Range[excelRange].Font.Size = size;
            worksheet.Range[excelRange].Font.Underline = underlineStyle;
        }

        public static async Task FormatExcelText(this Microsoft.Office.Interop.Excel.Worksheet worksheet, int row, int column, bool bold, int size = 11,
            Microsoft.Office.Interop.Excel.XlUnderlineStyle underlineStyle = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone)
        {
            worksheet.Cells[row, column].Font.Bold = bold;
            worksheet.Cells[row, column].Font.Size = size;
            worksheet.Cells[row, column].Font.Underline = underlineStyle;
        }

        public static async Task BorderExcelCell(this Microsoft.Office.Interop.Excel.Worksheet worksheet, string excelRange,
            Microsoft.Office.Interop.Excel.XlLineStyle top = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone,
            Microsoft.Office.Interop.Excel.XlLineStyle bottom = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone,
            Microsoft.Office.Interop.Excel.XlLineStyle left = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone,
            Microsoft.Office.Interop.Excel.XlLineStyle right = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)
        {
            worksheet.Range[excelRange].Borders.Item[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = top;
            worksheet.Range[excelRange].Borders.Item[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = bottom;
            worksheet.Range[excelRange].Borders.Item[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = left;
            worksheet.Range[excelRange].Borders.Item[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = right;
        }

        public static async Task BorderExcelCell(this Microsoft.Office.Interop.Excel.Worksheet worksheet, int row, int column,
            Microsoft.Office.Interop.Excel.XlLineStyle top = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone,
            Microsoft.Office.Interop.Excel.XlLineStyle bottom = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone,
            Microsoft.Office.Interop.Excel.XlLineStyle left = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone,
            Microsoft.Office.Interop.Excel.XlLineStyle right = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone)
        {
            worksheet.Cells[row, column].Borders.Item[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = top;
            worksheet.Cells[row, column].Borders.Item[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = bottom;
            worksheet.Cells[row, column].Borders.Item[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = left;
            worksheet.Cells[row, column].Borders.Item[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = right;
        }
    }

    public static class Flattener
    {
        public static T[] Flatten<T>(this Array data)
        {
            var list = new List<T>();
            var stack = new Stack<IEnumerator>();
            stack.Push(data.GetEnumerator());
            do
            {
                for (var iterator = stack.Pop(); iterator.MoveNext();)
                {
                    if (iterator.Current is Array)
                    {
                        stack.Push(iterator);
                        iterator = (iterator.Current as IEnumerable).GetEnumerator();
                    }
                    else
                        list.Add((T)iterator.Current);
                }
            }
            while (stack.Count > 0);
            return list.ToArray();
        }
    }

    public static class NumberFormats
    {
        public static string Text { get; } = "@";
        public static string Bracket2D { get; } = "_(* #,##0.00_);_(* (#,##0.00);_(* \" - \"??_);_(@_)";
        public static string Bracket4D { get; } = "_(* #,##0.0000_);_(* (#,##0.0000);_(* \" - \"??_);_(@_)";
        public static string Accounting2D { get; } = "_-* #,##0.00_-;-* #,##0.00_-;_-* \"0.00\"_-;_-@_-";
        public static string Accounting3D { get; } = "_-* #,##0.000_-;-* #,##0.000_-;_-* \"0.000\"_-;_-@_-";
        public static string Accounting4D { get; } = "_-* #,##0.0000_-;-* #,##0.0000_-;_-* \"0.0000\"_-;_-@_-";
        public static string Percentage { get; } = "0.0000%";
        public static string PercentageOfPercentage { get; } = "0.00\"%\"";
        public static string Calendar { get; } = "dd/MM/yyyy";
    }
}
