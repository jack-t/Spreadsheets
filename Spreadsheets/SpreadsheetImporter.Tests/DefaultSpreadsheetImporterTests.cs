using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetImporter;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using SpreadsheetImporter.Tests.Stubs;

namespace SpreadsheetImporter.Tests
{
    [TestClass()]
    public class DefaultSpreadsheetImporterTests
    {
        [TestMethod()]
        public void ImportSpreadsheetTest()
        {
            Workbook workbook = new Workbook();
            workbook.Worksheets["Sheet1"]["A1"].Value = "column1";
            workbook.Worksheets["Sheet1"]["A2"].Value = "column2";
            workbook.Worksheets["Sheet1"]["A3"].Value = "column3";
            ISpreadsheetTemplate template = new StubSpreadsheetTemplate(workbook);
            ISqlConnectionProvider connectionProvider = new DefaultSqlConnectionProvider("main");
            DefaultSpreadsheetImporter importer = new DefaultSpreadsheetImporter(connectionProvider, "InsertIntoTestSP");

            DataTable table = new DataTable();
            table.Columns.Add("column1");
            table.Columns.Add("column2");
            table.Columns.Add("column3");

            table.Rows.Add(1, 2, 3);
            table.Rows.Add(1, 2, 3);
            table.Rows.Add(1, 2, 3);

            ImportData data = new ImportData(table, null, template);

            importer.ImportSpreadsheet(data);

        }
    }
}