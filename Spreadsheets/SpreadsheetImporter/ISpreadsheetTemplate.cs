using Spire.Xls;

namespace SpreadsheetImporter
{
    public interface ISpreadsheetTemplate
    {
        /// <summary>
        /// A stream which provides the template file such that it can be fed into a library that reads excel files.
        /// </summary>
        Workbook GetTemplateStream();
        /// <summary>
        /// The first row in the excel file that should contain data when it is populated.
        /// This is necessary, I think, because it's possible some other sort of data could be kept above the headers or w/e.
        /// </summary>
        int FirstDataRow { get; }

        string DataSheetName { get; }
        string GuidCell { get; }
        
    }
}