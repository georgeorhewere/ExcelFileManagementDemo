using ExcelDataReader;
using ExcelFileManagementDemo.Common;
using ExcelFileManagementDemo.Interface;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileManagementDemo
{
    public class ExcelReaderManager :IStudentReader
    {
        private string filePath;
        private DataSet studentData;

        public ExcelReaderManager()
        {
           
        }

        public ProcessStatus OpenDataFeed(string connection)
        {
            ProcessStatus status;            
            filePath = connection;
            
            status = OpenExcelWorkBook();
            
            return status;
        }

        private ProcessStatus OpenExcelWorkBook()
        {
            ProcessStatus status = new ProcessStatus();
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Use the AsDataSet extension method
                        studentData = reader.AsDataSet(GetExcelDataSetConfig());                       
                      
                        List<string> badSheets = new List<string>();                        

                        if (VerifyColumnHeaders(studentData, badSheets))
                        {
                            status.success = true;
                            status.message = $"File was opened and in the correct format";
                        }
                        else
                        {
                            status.success = false;
                            status.message = $"File was opened but one or more sheets have invalid coumns. { string.Join(", ", badSheets)}";
                            status.data = badSheets;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                status.message = $"There was an error opening the file { filePath }. The error is : { ex.Message }";                
            }
            
            return status;
        }

        private bool VerifyColumnHeaders(DataSet importDataset, List<string> badSheets)
        {
            bool hasValidCoumnSet = true;
            int count = importDataset.Tables.Count;

            for (int sheetIndex = 0; sheetIndex < count; sheetIndex++)
            {
                var excelSheet = importDataset.Tables[sheetIndex];
                //Console.WriteLine($"sheets {excelSheet.TableName}");
                // validate column Headers 
                int columnCount = excelSheet.Columns.Count;
                var sheetColumns = excelSheet.Columns;
                List<string> badColumns = new List<string>();

                for (int col = 0; col < columnCount; col++)
                {
                    if (!FileHeaderDefinitions.ColumnDefinitions().Contains(sheetColumns[col].ColumnName))
                    {
                        badColumns.Add(sheetColumns[col].ColumnName);
                        hasValidCoumnSet = false;
                    }
                }

                if (badColumns.Any())
                {
                    badSheets.Add($"Sheet : {excelSheet.TableName}, Columns : { string.Join(" ,", badColumns) } ");
                }
                var rowCount = excelSheet.Rows.Count;
                Console.WriteLine($"number of rows : {rowCount}");
            }
            return hasValidCoumnSet;
        }

        private ExcelDataSetConfiguration GetExcelDataSetConfig()
        {
            return new ExcelDataSetConfiguration()
            {
                // Gets or sets a value indicating whether to set the DataColumn.DataType 
                // property in a second pass.
                //UseColumnDataType = true,

                // Gets or sets a callback to determine whether to include the current sheet
                // in the DataSet. Called once per sheet before ConfigureDataTable.
                //FilterSheet = (tableReader, sheetIndex) => true,

                // Gets or sets a callback to obtain configuration options for a DataTable. 
                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                {
                    // Gets or sets a value indicating the prefix of generated column names.
                  //  EmptyColumnNamePrefix = "Column",

                    // Gets or sets a value indicating whether to use a row from the 
                    // data as column names.
                    UseHeaderRow = true,

                    // Gets or sets a callback to determine which row is the header row. 
                    // Only called when UseHeaderRow = true.
                    //ReadHeaderRow = (rowReader) =>
                    //{
                    //    // F.ex skip the first row and use the 2nd row as column headers:
                    //    rowReader.Read();
                    //},

                    // Gets or sets a callback to determine whether to include the 
                    // current row in the DataTable.
                    FilterRow = (rowReader) =>
                    {
                        return true;
                    },

                    // Gets or sets a callback to determine whether to include the specific
                    // column in the DataTable. Called once per column after reading the 
                    // headers.
                    FilterColumn = (rowReader, columnIndex) =>
                    {
                        return true;
                    }
                }
            };

        }

        public ProcessStatus VerifyInputData()
        {
            ProcessStatus status = new ProcessStatus();
            //check for duplicate SSN
            var ContainsDuplicates = HasDuplicateStudentSSN();
            // Check for Same First,Last, DOB and Grade in the same school
            return status;
        }


        private bool HasDuplicateStudentSSN()
        {
            if(studentData != null)
            {
                Hashtable studentCache = new Hashtable();
                var sheetCount = studentData.Tables.Count;
                for(var index = 0; index < sheetCount; index++)
                {
                    var currentSheet = studentData.Tables[index];
                    // Display row contents
                    using(var reader = studentData.CreateDataReader())
                    {
                        while (reader.Read())
                        {   
                            //cache body
                            var SSN = reader.GetValue(reader.GetOrdinal(FileHeaderDefinitions.StudentSSN));
                            var firstName = reader.GetValue(reader.GetOrdinal(FileHeaderDefinitions.FirstName));

                            Console.WriteLine($"Student SSN for  {firstName } is { SSN }");
                            if (string.IsNullOrEmpty((SSN as string)))
                            {
                                // generate SSN
                                var random = new Random();
                                var tempSSN = $"{random.Next(1, 999)}-{random.Next(1, 99)}-{random.Next(1, 9999)}";
                                
                                //chceck if in cache                                
                                while (studentCache.ContainsKey(tempSSN))
                                {
                                    tempSSN = $"{random.Next(1, 999)}-{random.Next(1, 99)}-{random.Next(1, 9999)}";
                                }

                                SSN = tempSSN;
                            }

                            

                            //Add SSN, record to cache and flag if record has duplicate
                            studentCache.Add(SSN,  firstName );
                            

                            
                           

                        }
                    }
                }

                Console.WriteLine($"");
                Console.WriteLine($"Cache Items");
                foreach (DictionaryEntry item in studentCache)
                {
                    Console.WriteLine($"{ item.Key } : { item.Value} ");
                }
            }
            return false;
        }


    }
}
