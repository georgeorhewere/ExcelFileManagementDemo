using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileManagementDemo.Common
{
    public class FileHeaderDefinitions
    {
        private static List<string> columnDefinitions = new List<string> { "SchoolCode",
                                                                            "SchoolName",
                                                                            "FirstName",
                                                                            "LastName",
                                                                            "DOB",
                                                                            "Grade",
                                                                            "Gender",
                                                                            "StudentID",
                                                                            "Student SSN",
                                                                            "Internal_ID",
                                                                            };

        public static List<string> ColumnDefinitions()
        {
            return columnDefinitions;
        }

    }
}
