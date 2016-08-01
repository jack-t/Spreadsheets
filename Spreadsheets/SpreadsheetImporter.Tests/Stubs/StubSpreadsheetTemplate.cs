using Spire.Xls;

namespace SpreadsheetImporter.Tests.Stubs
{
    public class StubSpreadsheetTemplate : ISpreadsheetTemplate
    {
        public Workbook GetTemplateStream()
        {
            return _package;
        }

        public int FirstDataRow => 2;
        public int HeaderRow => 1;
        public string DataSheetName => "Sheet1";
        public string GuidCell => "AA1";

        private Workbook _package;

        public StubSpreadsheetTemplate(Workbook package)
        {
            _package = package;
        }
    }
}