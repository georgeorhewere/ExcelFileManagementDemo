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
        private Hashtable schoolCache;

        public ExcelReaderManager()
        {
            schoolCache = new Hashtable();
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

        public ProcessStatus ValidateInputFile()
        {
            ProcessStatus status = new ProcessStatus();
            //check for duplicate SSN
            var isInvalidInput = HasFlaggedErrors();
            if (isInvalidInput)
            {
                //write to File 
                ExcelUpdateManager writerManager = new ExcelUpdateManager();
              
                if (writerManager.OpenExcel(filePath))
                {
                    writerManager.WriteErrorsToFile(studentData);

                    writerManager.CloseExcel();
                }

            }

            status.success = !(isInvalidInput);
            
            return status;
        }


        private bool HasFlaggedErrors()
        {
            if(studentData != null)
            {
                var sheetCount = studentData.Tables.Count;
                for(var index = 0; index < sheetCount; index++)
                {
                    var currentSheet = studentData.Tables[index];
                    currentSheet.Columns.Add("Error", typeof(string));
                    
                    // Display row contents
                    using(var reader = currentSheet.CreateDataReader())
                    {
                        var rowCount = 0;  
                        while (reader.Read())
                        {
                            //validate fields                                                    
                            List<string> recordErrors = new List<string>();
                            //required fields
                            ValidateRequiredFields(reader, recordErrors);

                            var cacheItem = GetStudentCacheDTO(reader);

                            //update school cache
                            if(cacheItem.SchoolCode != null && cacheItem.SchoolName != null)
                            {
                                if (!schoolCache.ContainsKey(cacheItem.SchoolCode))
                                {
                                    schoolCache.Add(cacheItem.SchoolCode, cacheItem.SchoolName);
                                }
                                else
                                {
                                    // same school code and different school name
                                    if(!(schoolCache[cacheItem.SchoolCode] as string).Equals(cacheItem.SchoolName))
                                    {
                                        recordErrors.Add($"There is a different school name associated with this school code");
                                    }
                                }
                            }

                            string Student_SSN = getStudentSSN(currentSheet, reader,cacheItem,recordErrors);                          

                            ////Add SSN, record to cache and flag if record has duplicate  
                            var duplicateQuery = $"[{FileHeaderDefinitions.StudentSSN}] = '{Student_SSN}'";
                            var hasDuplicateSSN = currentSheet.Select(duplicateQuery).Count() > 1;

                            if (hasDuplicateSSN)
                            {
                                recordErrors.Add($"Another student exists with the same social security number");
                            }

                            var duplicateStudentIdQuery = $"[{FileHeaderDefinitions.StudentID}] = '{cacheItem.StudentID}' AND [{FileHeaderDefinitions.SchoolCode}] = '{cacheItem.SchoolCode}'";
                            var IsDuplicateStudentId = currentSheet.Select(duplicateStudentIdQuery).Count() > 1;

                            if(IsDuplicateStudentId)
                                recordErrors.Add($"This student id is associated with another student in this school : {cacheItem.SchoolName} ");

                            //add errors to new column
                            if (recordErrors.Any())
                            {
                                // Console.WriteLine($"Error :  { string.Join(", ", recordErrors) } ");
                                var row = currentSheet.Rows[rowCount];
                                row["Error"] = string.Join(",", recordErrors);
                            }
                            rowCount++;
                        }
                    }
                    currentSheet.AcceptChanges();                    
                }

                

                var hasErrors = $"Error IS NOT NULL";
                return studentData.Tables[0].Select(hasErrors).Any();
            }
            else
            {
                return false;
            }

            
        }

        private string getStudentSSN(DataTable currentSheet, DataTableReader reader, StudentCacheDTO cacheItem, List<string> recordErrors)
        {
            string Student_SSN;
            if (reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.StudentSSN)))
            {
                // generate SSN
                var random = new Random();
                Student_SSN = $"{random.Next(1, 999)}-{random.Next(1, 99)}-{random.Next(1, 9999)}";
                //chceck if in cache                                
                while (currentSheet.Select($"[{FileHeaderDefinitions.StudentSSN}] = '{Student_SSN}'").Any())
                {
                    Student_SSN = $"{random.Next(1, 999)}-{random.Next(1, 99)}-{random.Next(1, 9999)}";
                }
                // check that same Name,DOB and grade do not exist in the same school
                //Grade condition
                var gradeFilter = cacheItem.Grade != null ? "Grade = " + cacheItem.Grade : "Grade Is Null";
                var duplicateStudentQuery = $"FirstName = '{cacheItem.FirstName}' AND LastName = '{cacheItem.LastName}' AND DOB = '{cacheItem.DOB}' AND {gradeFilter} AND SchoolCode = '{ cacheItem.SchoolCode }'";
                var hasMultipleEntires = currentSheet.Select(duplicateStudentQuery).Count() > 1;
                
                if (hasMultipleEntires)
                    recordErrors.Add($"Another student has the same record details in this school");
            }
            else
            {
                Student_SSN = reader.GetString(reader.GetOrdinal(FileHeaderDefinitions.StudentSSN));
            }

            return Student_SSN;
        }

        private StudentCacheDTO GetStudentCacheDTO(DataTableReader reader)
        {
            var cacheItem = new StudentCacheDTO();

            cacheItem.FirstName = reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.FirstName)) ? null : reader.GetString(reader.GetOrdinal(FileHeaderDefinitions.FirstName));
            cacheItem.LastName = reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.LastName)) ? null : reader.GetString(reader.GetOrdinal(FileHeaderDefinitions.LastName));
            cacheItem.DOB = reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.DOB)) ? (DateTime?) null : reader.GetDateTime(reader.GetOrdinal(FileHeaderDefinitions.DOB));
            cacheItem.SchoolCode = reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.SchoolCode)) ? null : reader.GetString(reader.GetOrdinal(FileHeaderDefinitions.SchoolCode));
            cacheItem.SchoolName = reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.SchoolName)) ? null : reader.GetString(reader.GetOrdinal(FileHeaderDefinitions.SchoolName));
            cacheItem.Grade = reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.Grade)) ? (double?)null : reader.GetDouble(reader.GetOrdinal(FileHeaderDefinitions.Grade));
            cacheItem.StudentID = reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.StudentID)) ? null : Convert.ToString(reader.GetValue(reader.GetOrdinal(FileHeaderDefinitions.StudentID)));

            return cacheItem;
        }

        private void ValidateRequiredFields(DataTableReader reader, List<string> recordErrors)
        {
            if (reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.FirstName)))
            {
                recordErrors.Add($"Missing First Name");
            }
            if (reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.LastName)))
            {
                recordErrors.Add($" Missing Last Name");
            }
            if (reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.SchoolCode)))
            {
                recordErrors.Add($" Missing School Code");
            }
            if (reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.SchoolName)))
            {
                recordErrors.Add($"Missing School Name");
            }
            if (reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.DOB)))
            {
                recordErrors.Add($"Missing Date of Birth");
            }
            if (reader.IsDBNull(reader.GetOrdinal(FileHeaderDefinitions.StudentID)))
            {
                recordErrors.Add($"Missing Student ID");
            }
        }

    }
}
