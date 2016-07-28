using System.IO;

namespace SpreadsheetImporter
{
    public class SpreadsheetExporterService
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="template">Something you find in the database, most likely</param>
        /// <param name="data">The current data pulled from the database as requested by the client</param>
        /// <returns>A stream which can be turned into a download-able file to be returned to the user.</returns>
        public Stream CreateSpreadsheet(ISpreadsheetTemplate template, ISpreadsheetData data)
        {
            return null;
        }
    }
}