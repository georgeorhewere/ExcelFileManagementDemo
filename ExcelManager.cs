using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using excelProvider = Microsoft.Office.Interop.Excel;
using System.Collections;

namespace ExcelFileManagementDemo
{
    public class ExcelManager
    {
        excelProvider.Application xlApp = null;
        excelProvider.Workbooks workbooks = null;
        excelProvider.Workbook workbook = null;
        Hashtable sheets;
        string workbookFilePath;




        public void OpenExcel(string excelFilePath)
        {
            if (!string.IsNullOrEmpty(excelFilePath))
            {
                workbookFilePath = excelFilePath;
                xlApp = new excelProvider.Application();
                workbooks = xlApp.Workbooks;
                workbook = workbooks.Open(excelFilePath);
                sheets = new Hashtable();
                int count = 1;
                // Storing worksheet names in Hashtable.
                foreach (excelProvider.Worksheet sheet in workbook.Sheets)
                {
                    sheets[count] = sheet.Name;
                    Console.WriteLine($"Sheet Name is  {sheets[count]}");
                    count++;
                }
                Console.WriteLine($"number of sheets is  {workbook.Sheets.Count}");
            }
            else
            {
                Console.WriteLine($"File path is not valid.");
            }
        }

        private void closeWorkbook()
        {
           if(!string.IsNullOrEmpty(workbookFilePath))
                workbook.Close(false, workbookFilePath, null); // Close the connection to workbook
        }
        public void CloseExcel()
        {
            closeWorkbook();
            Marshal.FinalReleaseComObject(workbook); // Release unmanaged object references.
            workbook = null;

            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            workbooks = null;

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            xlApp = null;
        }

        public string GetCellData(string sheetName, string colName, int rowNumber)
        {
            

            string value = string.Empty;
            int sheetValue = 0;
            int colNumber = 0;

            if (sheets.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                excelProvider.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as excelProvider.Worksheet;
                excelProvider.Range range = worksheet.UsedRange;

                for (int i = 1; i <= range.Columns.Count; i++)
                {
                    string colNameValue = Convert.ToString((range.Cells[1, i] as excelProvider.Range).Value2);

                    if (colNameValue.ToLower() == colName.ToLower())
                    {
                        colNumber = i;
                        break;
                    }
                    Console.WriteLine(colNameValue);
                }

                value = Convert.ToString((range.Cells[rowNumber, colNumber] as excelProvider.Range).Value2);
                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            CloseExcel();
            return value;
        }

        public List<string> GetHeaderRow(string sheetName)
        {
            if (sheets.ContainsValue(sheetName))
            {
                
            }

            return new List<string>();
        }
    }
}
