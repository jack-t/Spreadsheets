using System.Collections.Generic;
using Spire.Xls;

namespace SpreadsheetImporter
{
    public interface ISpreadsheetTemplate
    {
        /// <summary>
        /// A stream which provides the template file such that it can be fed into a library that reads excel files.
        /// </summary>
        Workbook GetTemplate();
        /// <summary>
        /// The first row in the excel file that should contain data when it is populated.
        /// This is necessary, I think, because it's possible some other sort of data could be kept above the headers or w/e.
        /// </summary>
        int FirstDataRow { get; }
        int HeaderRow { get; }
        string DataSheetName { get; }
        string GuidCell { get; }
        /// <summary>
        /// Maps columns in the data recieved to the columns in the template.
        /// </summary>
        Dictionary<string, int> ExportColumnMap { get; }
        Dictionary<int, string> ImportColumnMap { get; }


    }
}