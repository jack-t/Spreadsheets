using System.Collections.Generic;
using Spire.Xls;

namespace SpreadsheetImporter.Tests.Stubs
{
    public class FindGuidTestTemplateStub : ISpreadsheetTemplate
    {
        public Workbook GetTemplate()
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

        public int HeaderRow
        {
            get { throw new System.NotImplementedException(); }
        }

        public string DataSheetName => "Sheet1";
        public string GuidCell => "AA1";
        public Dictionary<string, int> ExportColumnMap { get; }
        public Dictionary<int, string> ImportColumnMap { get; }
        public int FindLastRowOfData(Worksheet worksheet)
        {
            throw new System.NotImplementedException();
        }
    }
}