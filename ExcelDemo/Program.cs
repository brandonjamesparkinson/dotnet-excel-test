using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"C:\Demos\YouTubeDemo.xlsx");

            var people = GetSetupData();

            await SaveExcelFile(people, file);
        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add("MainReport");

            var range = ws.Cells["A1"].LoadFromCollection(people, true);
            range.AutoFitColumns();

            await package.SaveAsync();
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new()
            {
                new() { Id = 1, FirstName = "Brandon", LastName = "Parkinson" },
                new() { Id = 2, FirstName = "Jane", LastName = "Smith" },
                new() { Id = 3, FirstName = "Susan", LastName = "Storm" },
            };

            return output;
        }
    }
}
