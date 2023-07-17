using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace JSONParser;

class Program
{
    static void Main()
    {

        // Set the EPPlus license context for line 27
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Read the json file
        string json = File.ReadAllText("input.json");

        // Parse the json data into a new 'data object' of the type JObject from Newtonsoft lib
        JObject data = JObject.Parse(json);

        // Create a new Excel package
        using (ExcelPackage package = new ExcelPackage())
        {
            // Create a new worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

            // header row
            int rowIndex = 1;
            int colIndex = 1;
            foreach (var property in data)
            {
                worksheet.Cells[rowIndex, colIndex].Value = property.Key;
                colIndex++;
            }

            // data row doesnt work yet...
            rowIndex = 2;
            foreach (var item in data["items"])
            {
                colIndex = 1;
                foreach (var property in item)
                {
                    worksheet.Cells[rowIndex, colIndex].Value = property.Key;
                    colIndex++;
                }
                rowIndex++;
            }

            // Save the Excel package
            package.SaveAs(new FileInfo("output1.xlsx"));
        }

        Console.WriteLine("saved to output.xlsx");
    }
}
