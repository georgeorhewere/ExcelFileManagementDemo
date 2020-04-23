using ExcelFileManagementDemo.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileManagementDemo.Interface
{
    public interface IStudentReader
    {
        ProcessStatus OpenDataFeed(string connection);
        ProcessStatus VerifyInputData(string connection);

    }
}
