using ExcelFileManagementDemo.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileManagementDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFile = "F:/Projects/Test Files/StudentData.xlsx";

            IStudentReader manager = new ExcelReaderManager();
            var status = manager.OpenDataFeed(inputFile);
            //ExcelManager manager = new ExcelManager();
            //manager.OpenExcel(inputFile);
           var fileStateInfo = $"Status: { status.success } Message:{ status.message }";
           Console.WriteLine(fileStateInfo);
            if (status.success)
            {
                manager.VerifyInputData();
            }
            

            Console.ReadLine();
        }
    }
}
