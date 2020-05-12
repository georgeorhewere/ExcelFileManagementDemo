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
                WorksheetPart worksheetPart = InsertErrorWorksheet(workbookPart);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                List<string> columnNames = new List<string>();

                //Row headerRow = rows.ElementAt(0);
              // int errorColIndex = headerRow.Count();
                WorkbookStylesPart styles = workbookPart.WorkbookStylesPart;

                //Write Dataset
                var dataTable = studentData.Tables[0];
                



                //var errorStyle = SetErrorStyle(styles.Stylesheet);

                //// Set columns 
                //foreach (Cell cell in headerRow)
                //{
                //    columnNames.Add(getCellValue(workbookPart, cell));
                //    //Console.Write($"{} ");
                //}


                if(dataTable.Columns.Contains("Error"))
                {
                    var columnList = new List<string> { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L" };
                    //add header column
                    var headerRow = AddRowToErrorSheet(sheetData);
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        var cellReference = columnList[column.Ordinal] + headerRow.RowIndex;
                        // If there is not a cell with the specified column name, insert one.                          

                        // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                        Cell refCell = null;
                        foreach (Cell cell in headerRow.Elements<Cell>())
                        {
                            if (cell.CellReference.Value.Length == cellReference.Length)
                            {
                                if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                                {
                                    refCell = cell;
                                    break;
                                }
                            }
                        }

                        var newRowCell = AddRowText(workbookPart, headerRow, column.ColumnName);
                        newRowCell.CellReference = cellReference;
                        headerRow.InsertBefore(newRowCell, refCell);
                    }

                    foreach (DataRow errorRow in dataTable.Rows)
                    {
                        var row = AddRowToErrorSheet(sheetData);
                        foreach (DataColumn column in dataTable.Columns)
                        {
                            var cellReference = columnList[column.Ordinal] + row.RowIndex;
                            // If there is not a cell with the specified column name, insert one.                          
                            
                                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                                Cell refCell = null;
                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    if (cell.CellReference.Value.Length == cellReference.Length)
                                    {
                                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                                        {
                                            refCell = cell;
                                            break;
                                        }
                                    }
                                }

                               var newRowCell = AddRowText(workbookPart, row, errorRow[column.ColumnName].ToString());
                                newRowCell.CellReference = cellReference;
                                row.InsertBefore(newRowCell, refCell);
                        }

                    }

                }
                Console.WriteLine();

                //Write data to datatable 
                //var indexDiff = 2;
                //var dataTable = studentData.Tables[0];
                //foreach (Row row in rows.Skip(1))
                //{

                //    var rowIndex = Convert.ToInt32(row.RowIndex.ToString()) - indexDiff;
                //    var errorText= dataTable.Rows[rowIndex].Field<string>("Error");

                //    var newRowCell = AddRowText(workbookPart, row, errorText);
                //    // newRowCell.StyleIndex = errorStyle;
                //    if (row.Count() < errorColIndex)
                //    {
                //        Console.WriteLine($"Add cells between { row.Count() } to {errorColIndex}");
                //        for(int x = row.Count(); x <= errorColIndex; x++)
                //        {
                //            Cell emptyCell = new Cell() { CellReference = row.RowIndex.ToString(), CellValue = new CellValue()};
                //            row.InsertAt(emptyCell, x);                            
                //        }

                //    }

                //    if (row.ElementAt(errorColIndex) != null)
                //    {
                //        var existingCell = row.Descendants<Cell>().ElementAt(errorColIndex);
                //        existingCell.CellValue.Text = newRowCell.CellValue.InnerText;
                //        existingCell.DataType = newRowCell.DataType;
                //       // existingCell.StyleIndex = errorStyle;
                //    }
                //    else
                //    {
                //        row.InsertAt(newRowCell, errorColIndex);
                //    }

                //    Console.WriteLine(row.Count());
                //}

                 workSheet.Save();
            }


        }

        private Row AddRowToErrorSheet(SheetData sheetData)
        {
            var rowIndex = sheetData.Elements<Row>().Count() + 1;
            Row row;

            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = Convert.ToUInt32(rowIndex) };
             //   sheetData.Append(row);
            }
            sheetData.Append(row);

            return row;
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

        // Inserts a new worksheet for error.
        private static WorksheetPart InsertErrorWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart errorWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            errorWorksheetPart.Worksheet = new Worksheet(new SheetData());
            errorWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(errorWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = $"StudentImportErrors{DateTime.Now.Date.Day}{DateTime.Now.Date.Month}{DateTime.Now.Date.Year}{DateTime.Now.Hour}{DateTime.Now.Millisecond}";

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return errorWorksheetPart;
        }

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
    }
}

