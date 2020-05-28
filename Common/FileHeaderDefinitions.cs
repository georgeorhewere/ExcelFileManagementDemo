using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileManagementDemo.Common
{
    public class FileHeaderDefinitions
    {

        public static string StudentSSN = "Student SSN";
        public static string FirstName = "FirstName";
        public static string MiddleNameDesc = "MiddleName";
        public static string LastName = "LastName";
        public static string SchoolCode = "SchoolCode";
        public static string SchoolName = "SchoolName";
        public static string Grade = "Grade";
        public static string DOB = "DOB";
        public static string StudentID = "StudentID";
        public static string Gender = "Gender";
        public static string Internal_ID = "Internal_ID";

        private static List<string> columnDefinitions = new List<string> { SchoolCode,
                                                                            SchoolName,
                                                                            FirstName,
                                                                            MiddleNameDesc,
                                                                            LastName,
                                                                            DOB,
                                                                            Grade,
                                                                            Gender,
                                                                            StudentID,
                                                                            StudentSSN,
                                                                            Internal_ID,
                                                                            };

        public static List<string> ColumnDefinitions()
        {
            return columnDefinitions;
        }

     

    }
}
