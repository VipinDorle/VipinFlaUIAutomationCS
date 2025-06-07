
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VipinFlaUIAutomationCS.Utility
{
    internal class ExcelUtility
    {
        public static Dictionary<string, string> ReadExcelFile(string ExcelFile, string TestSet,  string TestCaseID)
        {
            var data = new Dictionary<string, string>();

            if (!File.Exists(ExcelFile))
                throw new FileNotFoundException($"Excel file not found: {ExcelFile}");

            using (var workbook = new XLWorkbook(ExcelFile))
            {
                if (!workbook.Worksheets.Contains(TestSet))
                    throw new ArgumentException($"Sheet '{TestSet}' not found in the Excel file.");

                var worksheet = workbook.Worksheet(TestSet);
                var rows = worksheet.RangeUsed().RowsUsed();

                var headerRow = worksheet.FirstRowUsed();
                var headers = new List<string>();

                // Read headers starting from column 2 (since column 1 is TestCaseID)
                for (int col = 2; col <= headerRow.CellCount(); col++)
                {
                    headers.Add(headerRow.Cell(col).GetValue<string>().Trim());
                }

                // Iterate through rows to find the TestCaseID
                foreach (var row in rows.Skip(1)) // Skip header row
                {
                    var currentID = row.Cell(1).GetValue<string>().Trim();

                    if (currentID.Equals(TestCaseID, StringComparison.OrdinalIgnoreCase))
                    {
                        for (int col = 2; col <= headers.Count + 1; col++)
                        {
                            var key = headers[col - 2]; // headers index starts from 0
                            var value = row.Cell(col).GetValue<string>().Trim();
                            data[key] = value;
                        }
                        break;
                    }
                }
            }

            return data;
        }
    }
}
