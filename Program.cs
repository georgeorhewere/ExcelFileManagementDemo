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
            //string inputFile = "F:/Projects/Test Files/StudentData.xlsx";
            string CsvInputFile = "F:/Projects/Test Files/StudentCSV.csv";
            CsvManager.AppendErrorsToLine(CsvInputFile);

            // IStudentReader manager = new ExcelReaderManager();
            //var status = manager.OpenDataFeed(inputFile);

            //ExcelReaderManager manager = new ExcelReaderManager();
            //manager.OpenExcel(inputFile);
            // var fileStateInfo = manager.ValidateInputFile();
            //Console.WriteLine(fileStateInfo);
            // if (status.success)
            // {
            //    var validationResult =  manager.ValidateInputFile();
            //     if (validationResult.success)
            //     {
            //         Console.WriteLine("The data is in the right format. ");
            //     }
            //     else
            //     {
            //         Console.WriteLine("We encountered a problem with your CSV or the data is not in the correct form");
            //     }
            // }

            //var xlManager = new OpenXmlManager();
            //xlManager.openWorkBook(inputFile);



            Console.ReadLine();
        }
    }
}
