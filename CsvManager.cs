﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelFileManagementDemo
{
    public class CsvManager
    {

        public static void AppendErrorsToLine(string filePath)
        {
            List<String> lines = new List<String>();

            if (File.Exists(filePath))
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    String line;
                    StringBuilder headerLine = new StringBuilder(reader.ReadLine());
                   // headerLine.Append(",Errors");
                    //lines.Add(headerLine.ToString());

                    while ((line = reader.ReadLine()) != null)
                    {
                      
                        //if (line.Contains(","))
                        //{
                        String[] split = SplitCSV(line).ToArray();

                        //    //if (split[1].Contains("34"))
                        //    //{
                        //    //    split[1] = "100";
                        //    //    line = String.Join(",", split);
                        //    //}
                        //}

                        lines.Add(line);
                    }
                }

                //using (StreamWriter writer = new StreamWriter(filePath, false))
                //{
                //    foreach (String line in lines)
                //        writer.WriteLine(line);
                //}
                //foreach (var item in lines)
                //{
                //    Console.WriteLine(item);
                //}
            }
        }


        public static IEnumerable<string> SplitCSV(string input)
        {
            Regex csvSplit = new Regex("(?:^|,)(\"(?:[^\"]+|\"\")*\"|[^,]*)", RegexOptions.Compiled);

            foreach (Match match in csvSplit.Matches(input))
            {
                yield return match.Value.TrimStart(',');
            }
        }
    }
}
