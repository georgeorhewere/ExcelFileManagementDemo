using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelFileManagementDemo.Common;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileManagementDemo
{
    public class OpenXmlWriterManager
    {
        IMemoryCache memoryCache;

        // open workbook
        //add new worksheet
        // write headers
        //write data
        // ssnupdate cells
        public OpenXmlWriterManager(IMemoryCache _memoryCache)
        {
            memoryCache = _memoryCache;
        }


        public void writeDataSetToSheet(string fileName, DataSet dataSet)
        {
            
            using (SpreadsheetDocument xl = SpreadsheetDocument.Open(fileName, true))
            {
                List<OpenXmlAttribute> oxa;
                OpenXmlWriter oxw;

                WorkbookPart workbookPart =   xl.WorkbookPart;
                WorksheetPart wsp = workbookPart.AddNewPart<WorksheetPart>();

                oxw = OpenXmlWriter.Create(wsp);
                oxw.WriteStartElement(new Worksheet());
                oxw.WriteStartElement(new SheetData());

                //write headers
                oxa = new List<OpenXmlAttribute>();
                oxa.Add(new OpenXmlAttribute("r", null, 1.ToString()));
                oxw.WriteStartElement(new Row(), oxa);
                foreach(var item in FileHeaderDefinitions.ColumnDefinitions())
                {
                   WriteRowCells(oxw, item);
                }


                oxw.WriteEndElement();

                var dataTable = dataSet.Tables[0];
                var count = dataTable.Rows.Count;
                for (int i = 0; i < count; ++i)
                {
                    oxa = new List<OpenXmlAttribute>();
                    var index = i + 2;
                    // this is the row index
                    oxa.Add(new OpenXmlAttribute("r", null, index.ToString()));
                    oxw.WriteStartElement(new Row(), oxa);

                    //schoolcode
                    WriteRowCells(oxw, dataTable.Rows[i].Field<String>(FileHeaderDefinitions.SchoolCode));
                    //schoolName
                    WriteRowCells(oxw, dataTable.Rows[i].Field<String>(FileHeaderDefinitions.SchoolName));
                    //FirstName
                    WriteRowCells(oxw, dataTable.Rows[i].Field<String>(FileHeaderDefinitions.FirstName));
                    //MiddleName
                    WriteRowCells(oxw, dataTable.Rows[i].Field<String>(FileHeaderDefinitions.MiddleNameDesc));
                    //LastName
                    WriteRowCells(oxw, dataTable.Rows[i].Field<String>(FileHeaderDefinitions.LastName));
                    //DOB
                    WriteRowCells(oxw, dataTable.Rows[i].Field<DateTime>(FileHeaderDefinitions.DOB).Date.ToString());

                    //for (int j = 1; j <= 100; ++j)
                    //{
                    //    oxa = new List<OpenXmlAttribute>();
                    //    // this is the data type ("t"), with CellValues.String ("str")
                    //    oxa.Add(new OpenXmlAttribute("t", null, "str"));

                    //    // it's suggested you also have the cell reference, but
                    //    // you'll have to calculate the correct cell reference yourself.
                    //    // Here's an example:
                    //    //oxa.Add(new OpenXmlAttribute("r", null, "A1"));

                    //    oxw.WriteStartElement(new Cell(), oxa);
                    //    oxw.WriteElement(new CellValue(string.Format("R{0}C{1}", i, j)));
                    //    // this is for Cell
                    //    oxw.WriteEndElement();
                    //}

                    // this is for Row
                    oxw.WriteEndElement();
                }

                // this is for SheetData
                oxw.WriteEndElement();
                // this is for Worksheet
                oxw.WriteEndElement();
                oxw.Close();

                oxw = OpenXmlWriter.Create(xl.WorkbookPart);
                oxw.WriteStartElement(new Workbook());
                oxw.WriteStartElement(new Sheets());

                // you can use object initialisers like this only when the properties
                // are actual properties. SDK classes sometimes have property-like properties
                // but are actually classes. For example, the Cell class has the CellValue
                // "property" but is actually a child class internally.
                // If the properties correspond to actual XML attributes, then you're fine.
                oxw.WriteElement(new Sheet()
                {
                    Name = "Sheet1",
                    SheetId = 1,
                    Id = xl.WorkbookPart.GetIdOfPart(wsp)
                });

                // this is for Sheets
                oxw.WriteEndElement();
                // this is for Workbook
                oxw.WriteEndElement();
                oxw.Close();

                xl.Close();



            }

            Console.WriteLine($" Student Data Done Writing");
        }

        private static void WriteRowCells(OpenXmlWriter oxw, string item)
        {
            List<OpenXmlAttribute> oxa = new List<OpenXmlAttribute>();
            oxa.Add(new OpenXmlAttribute("t", null, "str"));
            oxw.WriteStartElement(new Cell(), oxa);
            oxw.WriteElement(new CellValue(item));
            oxw.WriteEndElement();            
        }
    }
}
