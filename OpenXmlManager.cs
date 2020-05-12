using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileManagementDemo
{
    public class OpenXmlManager
    {



        public void openWorkBook(string fileName, DataSet studentData)
        {

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                List<string> columnNames = new List<string>();

                // Set columns 
                foreach (Cell cell in rows.ElementAt(0))
                {
                    columnNames.Add(getCellValue(workbookPart, cell));
                    //Console.Write($"{} ");
                }
                Console.WriteLine();

                //Write data to datatable 
                foreach (Row row in rows.Skip(1))
                {
                //    DataRow newRow = dt.NewRow();
                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        if (row.Descendants<Cell>().ElementAt(i).CellValue != null)
                        {
                            //Console.Write(row.Descendants<Cell>().ElementAt(i).CellValue.InnerXml);
                          //  Console.Write($"{getCellValue(workbookPart, row.Descendants<Cell>().ElementAt(i))} ");
                        }
                //        else
                //        {
                //            newRow[i] = DBNull.Value;
                //        }
                    }
                //    dt.Rows.Add(newRow);
                }
            }


        }

        private string getCellValue(WorkbookPart workbookPart, Cell cell)
        {
            var cellValue = cell.CellValue;
            var text = (cellValue == null) ? cell.InnerText : cellValue.Text;
            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            {
                text = workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(
                        Convert.ToInt32(cell.CellValue.Text)).InnerText;
            }
            return (text ?? string.Empty).Trim();
        }
    }
}
