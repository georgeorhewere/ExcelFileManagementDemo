using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader;
using ExcelManagerLibrary.Common;
using Microsoft.Extensions.Caching.Memory;

namespace ExcelManagerLibrary
{
    public class ExcelReadManager : IFileReader
    {
        private string fileName;

        public FileTaskStatus ReadFile(string path)
        {
            var response = new FileTaskStatus();
            FileInfo fileInfo = new FileInfo(path);
            fileName = fileInfo.Name;
            if (fileInfo.Exists)
            {
                using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    MemoryStream memStream = new MemoryStream();
                    memStream.SetLength(stream.Length);
                    stream.Read(memStream.GetBuffer(), 0, (int)stream.Length);
                    IMemoryCache memoryCache = MemoryCacheManager.MemoryCache;
                    memoryCache.Set(fileInfo.Name, memStream);

                }
                response.Success = true;
                response.Message = $"File saved in cache {fileName}";
            }
            else
            {
                response.Success = false;
                response.Message = $"File does not exist {fileName}";
            }
            
           return response;
        }

        public FileTaskStatus ValidateData()
        {
            var response = new FileTaskStatus();

            var stream = (MemoryStream)MemoryCacheManager.MemoryCache.Get(fileName);
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                //// Use the AsDataSet extension method
                //studentData = reader.AsDataSet(GetExcelDataSetConfig());

                //List<string> badSheets = new List<string>();

                //if (VerifyColumnHeaders(studentData, badSheets))
                //{
                //    response.Success = true;
                //    response.Message = $"File was opened and in the correct format";
                //}
                //else
                //{
                //    response.Success = false;
                //    response.Message = $"File was opened but one or more sheets have invalid coumns. { string.Join(", ", badSheets)}";
                //    response.Data = badSheets;
                //}
            }

            return response;
        }
    }
}
