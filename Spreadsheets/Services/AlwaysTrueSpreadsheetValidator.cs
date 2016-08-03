using SpreadsheetImporter;

namespace Services
{
    public class AlwaysTrueSpreadsheetValidator : ISpreadsheetValidator
    {
        public bool IsValidData(ImportData data)
        {
            return true;
        }
    }
}