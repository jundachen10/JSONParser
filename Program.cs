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

        // access "items" array

        JArray itemsArray = (JArray)data["items"];

        // So my target json is structured as an array with objects in it that has keys and values

        // Create a new Excel package
        using (ExcelPackage package = new ExcelPackage())
        {
            // Create a new worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

            // header row I want to take all the objects in the items array and read each key as the header columns
            int rowIndex = 1;
            int colIndex = 1;
            foreach (var property in itemsArray[0].ToObject<JObject>())
            {
                worksheet.Cells[rowIndex, colIndex].Value = property.Key;
                colIndex++;
            }

            // data row I want to take all of the objects in the items array


            // so the below wasn't working because I had to convert each value of each object into a string
            rowIndex = 2;
            foreach (var item in itemsArray)
            {
                colIndex = 1;
                foreach (var property in item.ToObject<JObject>())
                {
                    string valueString = Convert.ToString(property.Value);//convert tostring here of each property's value
                    worksheet.Cells[rowIndex, colIndex].Value = valueString;
                    colIndex++;
                }
                rowIndex++;
            }

            // Save the Excel package
            package.SaveAs(new FileInfo("output3.xlsx"));
        }

        Console.WriteLine("saved to output.xlsx");
    }
}
