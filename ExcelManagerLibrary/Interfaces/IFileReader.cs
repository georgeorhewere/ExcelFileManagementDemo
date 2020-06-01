using ExcelManagerLibrary.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelManagerLibrary
{
    public interface IFileReader
    {
        FileTaskStatus ReadFile(string path);

        FileTaskStatus ValidateData();
    }
}
