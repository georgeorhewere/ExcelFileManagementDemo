using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileManagementDemo.Common
{
    public class ProcessStatus
    {
        public bool success { get; set; }
        public string message { get; set; }

        public dynamic data { get; set; }
    }
}
