using System.Collections.Generic;
using System.Linq;
using Spire.Pdf.Tables;
using Spire.Xls;
using SpreadsheetImporter;

namespace Services
{
    public class TestTemplate : ISpreadsheetTemplate
    {
        public Workbook GetTemplate()
        {
            var workbook = new Workbook();
            workbook.LoadFromFile("template.xlsx");
            //workbook.Version = ExcelVersion.Version2013;

            return workbook;
        }

        public int FirstDataRow => 3;
        public int HeaderRow => 2;
        public string DataSheetName => "Sheet1";
        public string GuidCell => "AA1";

        public Dictionary<string, int> ExportColumnMap { get; } = new Dictionary<string, int>
        {
            ["State"] = 2,
            ["County"] = 3,
            ["ProductTypeName"] = 4,
            ["Pending"] = 5,
            ["CurrentPrice"] = 6
        };

        public Dictionary<int, string> ImportColumnMap { get; } = new Dictionary<int, string>
        {
            [7] = "State",
            [8] = "County",
            [9] = "ProductTypeName",
            [10] = "CurrentPrice"
        };

        public int FindLastRowOfData(Worksheet worksheet)
        {
            for (int i = FirstDataRow; ; i++)
            {
                if (IsRowEmpty(worksheet, i))
                    return i;
            }
        }

        private bool IsRowEmpty(Worksheet worksheet, int row)
        {
            // do all of the cells in this row's relevant range not have any value? If yes, return true.
            return worksheet.Range[$"B{row}:F{row}"].Cells.All(cell => !cell.HasNumber && !cell.HasBoolean && !cell.HasString);
        }
    }
}