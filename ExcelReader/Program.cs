using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var path = @"D:\ProgramData\VsProjects\ExcelReader\ExcelReader\Sample.xlsx";

            var (data, csvData) = ReadDataFrom<Person>(path);
            var base64csvData = ConvertToBase64(csvData);
        }

        static (List<T> data, string csvData) ReadDataFrom<T>(string filePath) where T: new()
        {
            var dict = new Dictionary<int, string>();
            var sb = new StringBuilder();
            var workbookFileInfo = new FileInfo(filePath);
            var result = new List<T>();
            using (var excelPackage = new ExcelPackage(workbookFileInfo))
            {
                var worksheet = excelPackage.Workbook.Worksheets[1]; 

                var rowCount = worksheet.Dimension.Rows;
                var columnCount = worksheet.Dimension.Columns;

                for (var columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                {
                    dict[columnIndex] = worksheet.Cells[1, columnIndex].Value.ToString();
                    sb.Append(dict[columnIndex]);
                    if(columnIndex < columnCount)
                    {
                        sb.Append(",");
                    }
                }

                sb.AppendLine();

                for (var rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                {
                    var obj = new T();
                    var props = obj.GetType().GetProperties();
                    for (var columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                    {
                        var value = worksheet.Cells[rowIndex, columnIndex].Value.ToString();
                        props.Where(p => p.Name.ToLower().Equals(dict[columnIndex].ToLower())).FirstOrDefault()?.SetValue(obj, value);
                        sb.Append(value);
                        if(columnIndex < columnCount)
                        {
                            sb.Append(",");
                        }
                    }
                    sb.AppendLine();
                    result.Add(obj);
                }
            }
            return (result, sb.ToString());
        }

        static string ConvertToBase64(string data)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(data));
        }
    }

    public class Person
    {
        public string Name { get; set; }
        public string BirthPlace { get; set; }
    }
}