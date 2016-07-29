namespace SpreadsheetImporter
{
    public interface ISpreadsheetValidator
    {
        /// <summary>
        /// This method may need to validate two things: 
        /// * The GUID is correct.
        /// * The data is in such a format that the importer will run correctly.
        /// </summary>
        /// <param name="template"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        bool IsValidData(ImportData data);
    }
}