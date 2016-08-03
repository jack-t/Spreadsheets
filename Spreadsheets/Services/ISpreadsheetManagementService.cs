using System;
using System.Data;
using System.IO;
using SpreadsheetImporter;

namespace Services
{
    public interface ISpreadsheetManagementService
    {
        void Upload(Stream stream);
        Stream Download(DataTable table);
    }

    public class SpreadsheetManagementService : ISpreadsheetManagementService
    {
        private readonly SpreadsheetService _service;

        public SpreadsheetManagementService(ISqlConnectionProvider connectionProvider)
        {
            _connectionProvider = connectionProvider;
            var exporter = new DefaultSpreadsheetExporter();
            var importer = new DefaultSpreadsheetImporter(connectionProvider, "InsertIntoFeesSP");
            var template = new TestTemplate();

            _service = new SpreadsheetService(exporter, importer, new AlwaysTrueSpreadsheetValidator(), template);
        }

        private readonly ISqlConnectionProvider _connectionProvider;



        public void Upload(Stream stream)
        {
            _service.ImportSpreadsheet(stream);
        }

        public Stream Download(DataTable table)
        {
            return _service.ExportSpreadsheetToStream(new ExportData(table, Guid.NewGuid()));
        }
    }
}