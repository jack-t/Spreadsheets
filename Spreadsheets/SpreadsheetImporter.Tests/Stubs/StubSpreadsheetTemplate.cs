using System.IO;

namespace SpreadsheetImporter.Tests.Stubs
{
    public class StubSpreadsheetTemplate : ISpreadsheetTemplate
    {
        public Stream TemplateStream { get; private set; }
        public int FirstDataRow { get { return 0; } }

        public StubSpreadsheetTemplate(Stream stream)
        {
            TemplateStream = stream;
        }
    }
}