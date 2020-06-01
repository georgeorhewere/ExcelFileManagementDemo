using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelManagerLibrary.Common
{
    public class FileTaskStatus
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public dynamic Data { get; set; }
    }
}
