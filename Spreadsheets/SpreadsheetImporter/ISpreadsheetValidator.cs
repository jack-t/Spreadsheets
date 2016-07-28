using System.Net.Configuration;

namespace SpreadsheetImporter
{
    public interface ISpreadsheetValidator
    {
        /// <summary>
        /// Looks at the GUID, mostly likely, and decides if this data is safe to import
        /// </summary>
        /// <param name="template"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        bool IsValidData(IImportData data);

    }
}