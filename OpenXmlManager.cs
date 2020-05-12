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

                Row headerRow = rows.ElementAt(0);
                int errorColIndex = headerRow.Count();
                WorkbookStylesPart styles = workbookPart.WorkbookStylesPart;
                var errorStyle = SetErrorStyle(styles.Stylesheet);

                // Set columns 
                foreach (Cell cell in headerRow)
                {
                    columnNames.Add(getCellValue(workbookPart, cell));
                    //Console.Write($"{} ");
                }
           

                if(studentData.Tables[0].Columns.Contains("Error"))
                {
                   var newCell = AddRowText(workbookPart, headerRow,"Error");                    
                   headerRow.InsertAt(newCell, errorColIndex);
                }
                Console.WriteLine();

                //Write data to datatable 
                var indexDiff = 2;
                var dataTable = studentData.Tables[0];
                foreach (Row row in rows.Skip(1))
                {
                 
                    var rowIndex = Convert.ToInt32(row.RowIndex.ToString()) - indexDiff;
                    var errorText= dataTable.Rows[rowIndex].Field<string>("Error");

                    var newRowCell = AddRowText(workbookPart, row, errorText);
                    // newRowCell.StyleIndex = errorStyle;
                    if (row.Count() < errorColIndex)
                    {
                        Console.WriteLine($"Add cells between { row.Count() } to {errorColIndex}");
                        for(int x = row.Count(); x <= errorColIndex; x++)
                        {
                            Cell emptyCell = new Cell() { CellReference = row.RowIndex.ToString(), CellValue = new CellValue()
                            };
                            row.InsertAt(emptyCell, x);                            
                        }

                    }

                    if (row.ElementAt(errorColIndex) != null)
                    {
                        var existingCell = row.Descendants<Cell>().ElementAt(errorColIndex);
                        existingCell.CellValue.Text = newRowCell.CellValue.InnerText;
                        existingCell.DataType = newRowCell.DataType;
                       // existingCell.StyleIndex = errorStyle;
                    }
                    else
                    {
                        row.InsertAt(newRowCell, errorColIndex);
                    }

                    Console.WriteLine(row.Count());
                }

              //  workSheet.Save();
            }


        }

        private Cell AddRowText(WorkbookPart workbookPart, Row headerRow,string rowText)
        {
            var headerIndex = InsertSharedStringItem(rowText, workbookPart.GetPartsOfType<SharedStringTablePart>().First());
            Cell newCell = new Cell() { CellReference = headerRow.RowIndex.ToString(), CellValue = new CellValue(headerIndex.ToString()), DataType = new EnumValue<CellValues>(CellValues.SharedString) };
            return newCell;
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

        private int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        private uint SetErrorStyle(Stylesheet stylesheet)
        {
            Fill fill = new Fill()
            {
                PatternFill = new PatternFill()
                {
                    PatternType = PatternValues.Solid,
                    BackgroundColor = new BackgroundColor() { Rgb = "D8D8D8" }
                }
            };
            stylesheet.Fills.AppendChild(fill);
            //Adding the  CellFormat which uses the Fill element 
            CellFormats cellFormats = stylesheet.CellFormats;
            CellFormat cf = new CellFormat();
            cf.FillId = stylesheet.Fills.Count;
            cellFormats.AppendChild(cf);

            stylesheet.Save();

            return stylesheet.CellFormats.Count;
        }
    }
}

