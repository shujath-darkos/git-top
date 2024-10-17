using Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmokeTestLogin.Utilities.Random_Data
{
    internal class ExcelSequentialHelper
    {
        private readonly string _filePath;

        public ExcelSequentialHelper(string filePath)
        {
            _filePath = filePath;
        }

        // Generate different types of data based on input type
        public string GenerateRandomData(string dataType, int row, int col)
        {
            string generatedValue = dataType switch
            {
                "sequential" => GetNextSequentialValue(ReadFromExcel(row, col)),
                "email" => $"user{GenerateRandomNumber(1000, 9999)}@example.com",
                "name" => $"User{GenerateRandomNumber(1, 100)}",
                "number" => GenerateRandomNumber(1, 1000).ToString(),
                _ => "default"
            };

            WriteToExcel(row, col, generatedValue);  // Write the new value to Excel
            return generatedValue;
        }

        // Generate the next sequential value (e.g., test1, test2...)
        private string GetNextSequentialValue(string currentValue)
        {
            int nextIndex = string.IsNullOrEmpty(currentValue) ? 1 : int.Parse(currentValue.Substring(4)) + 1;
            return $"test{nextIndex}";
        }

        // Generate a random number within a given range
        private int GenerateRandomNumber(int min, int max) => new Random().Next(min, max);

        // Read value from Excel
        private string ReadFromExcel(int row, int col)
        {
            using var fs = new FileStream(_filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = new XSSFWorkbook(fs);
            ISheet sheet = workbook.GetSheetAt(0);
            return sheet.GetRow(row)?.GetCell(col)?.ToString() ?? string.Empty;
        }

        // Write value to Excel
        private void WriteToExcel(int row, int col, string value)
        {
            IWorkbook workbook;
            using (var fs = new FileStream(_filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fs);
            }

            ISheet sheet = workbook.GetSheetAt(0);
            IRow excelRow = sheet.GetRow(row) ?? sheet.CreateRow(row);
            ICell cell = excelRow.GetCell(col) ?? excelRow.CreateCell(col);

            cell.SetCellValue(value);  // Set the new value

            using var fsOut = new FileStream(_filePath, FileMode.Create, FileAccess.Write);
            workbook.Write(fsOut);  // Save changes to the file
        }
    }
}
