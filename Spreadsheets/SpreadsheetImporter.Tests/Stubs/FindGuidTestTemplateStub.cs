using Spire.Xls;

namespace SpreadsheetImporter.Tests.Stubs
{
    public class FindGuidTestTemplateStub : ISpreadsheetTemplate
    {
        public Workbook GetTemplateStream()
        {
            throw new System.NotImplementedException();
        }

        public int FirstDataRow
        {
            get
            {
                throw new System.NotImplementedException();
            }
        }

        public string DataSheetName => "Sheet1";
        public string GuidCell => "AA1";
    }
}