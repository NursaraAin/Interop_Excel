using FxForwardExtract.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace FxForwardExtract.Modules
{
    public static class Module_Write
    {
        public static async Task WriteToExcel(List<FXForwardData> listFxForward)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbooks workbooks = xlApp.Workbooks;
            Excel.Workbook workbook = workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];

            PropertyInfo[] propertyInfos = typeof(FXForwardData).GetProperties();
            
            char letter = 'A';
            //int number = 1;
         
            foreach (PropertyInfo propertyInfo in propertyInfos)
            {
                worksheet.Range[letter + "1"].Value = propertyInfo.Name;
                letter++;
            }
            for(int j = 0;j<listFxForward.Count;j++)
            {
                letter = 'A';
                FXForwardData fxData = listFxForward[j];
                for(int i = 0; i < propertyInfos.Length; i++)
                {
                    string loc = letter + (j + 2).ToString();
                    worksheet.Range[loc].Value = propertyInfos[i].GetValue(fxData);
                    letter++;
                }
            }

            xlApp.Visible = true;

        }
    }
}
