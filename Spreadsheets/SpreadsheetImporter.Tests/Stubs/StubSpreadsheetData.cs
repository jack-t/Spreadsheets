using System.Data;

namespace SpreadsheetImporter.Tests.Stubs
{
    public class StubSpreadsheetData : ISpreadsheetData
    {
        public DataTable Table { get; set; }
    }
}