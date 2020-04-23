using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileManagementDemo.Common
{
    public class StudentCacheDTO
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public DateTime? DOB { get; set; }
        public int? Grade { get; set; }
        public string SchoolCode { get; set; }
        public string SchoolName { get; set; }

    }
}
