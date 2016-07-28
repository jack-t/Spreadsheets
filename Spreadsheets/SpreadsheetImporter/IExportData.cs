using System;
using System.Data;

namespace SpreadsheetImporter
{
    /// <summary>
    /// This is a separate interface because I'm virtually certain this will require some extra data to be passed
    /// </summary>
    public interface IExportData
    {
        DataTable Table { get; }
        Guid SheetGuid { get; }
    }
}