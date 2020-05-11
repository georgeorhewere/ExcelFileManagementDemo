using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using excelProvider = Microsoft.Office.Interop.Excel;
using System.Collections;
using ExcelFileManagementDemo.Interface;
using System.Data;
using ExcelFileManagementDemo.Common;

namespace ExcelFileManagementDemo
{
    public class ExcelUpdateManager : IStudentWriter
    {
        excelProvider.Application xlApp = null;
        excelProvider.Workbooks workbooks = null;
        excelProvider.Workbook workbook = null;
        Hashtable sheets;
        string workbookFilePath;




        public bool OpenExcel(string excelFilePath)
        {
            bool fileState = false;
            try
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
                    
                    fileState = true;
                }
                else
                {
                    Console.WriteLine($"File path is not valid.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return fileState;
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

        public void WriteErrorsToFile(DataSet dataset)
        {
            Console.WriteLine($"");
            Console.WriteLine($"Cache Items");
            foreach (DataRow item in dataset.Tables[0].Rows)
            {
                Console.WriteLine($"SSN : {item[FileHeaderDefinitions.StudentSSN] }, Name : {item[FileHeaderDefinitions.FirstName] } { item[FileHeaderDefinitions.LastName]}");
                Console.WriteLine($"Errors { item["Error"]}  ");
                Console.WriteLine($"");
            }
        }
    }
}
