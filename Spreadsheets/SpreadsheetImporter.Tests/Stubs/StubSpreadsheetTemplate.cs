using System.Collections.Generic;
using System.Linq;
using Spire.Xls;

namespace SpreadsheetImporter.Tests.Stubs
{
    public class StubSpreadsheetTemplate : ISpreadsheetTemplate
    {
        public Workbook GetTemplate()
        {
            return _package;
        }

        public int FirstDataRow => 2;
        public int HeaderRow => 1;
        public string DataSheetName => "Sheet1";
        public string GuidCell => "AA1";
        public Dictionary<string, int> ExportColumnMap { get; }
        public Dictionary<int, string> ImportColumnMap { get; }
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

        private Workbook _package;

        public StubSpreadsheetTemplate(Workbook package)
        {
            _package = package;
        }

        public StubSpreadsheetTemplate(Workbook workbook, Dictionary<string, int> exportDict, Dictionary<int, string> importDict)
        {
            _package = workbook;
            ExportColumnMap = exportDict;
            ImportColumnMap = importDict;
        }
    }
}