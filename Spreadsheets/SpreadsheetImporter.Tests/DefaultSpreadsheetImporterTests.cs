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
        [TestClass()]
        public class StripExcelDataTests
        {
            private Workbook _workbook;
            private ISpreadsheetTemplate _template;
            private ISqlConnectionProvider _connectionProvider;
            private DefaultSpreadsheetImporter _importer;
            private DataTable _table;

            [TestInitialize()]
            public void Init()
            {
                _workbook = new Workbook();
                _workbook.Worksheets["Sheet1"]["A1"].Value = "column1";
                _workbook.Worksheets["Sheet1"]["A2"].Value = "column2";
                _workbook.Worksheets["Sheet1"]["A3"].Value = "column3";
                _template = new StubSpreadsheetTemplate(_workbook, null, new Dictionary<int, string>
                {
                    [1] = "column1",
                    [2] = "column2",
                    [3] = "column3"
                });
                _connectionProvider = new DefaultSqlConnectionProvider("main");
                _importer = new DefaultSpreadsheetImporter(_connectionProvider, "dummy");

                _table = _importer.StripExcelData(_workbook, _template);
            }

            [ExpectedException(typeof(IndexOutOfRangeException))]
            [TestMethod()]
            public void StripExcelDataTest()
            {
                _workbook.Worksheets["Sheet1"]["A4"].Value = "column3";

                var columnName = _table.Columns[3].ColumnName; // this should throw an exception
            }

            [TestMethod()]
            public void StripExcelData_CorrectColumnNames()
            {
                Assert.AreEqual("column1", _table.Columns[0].ColumnName);
                Assert.AreEqual("column2", _table.Columns[1].ColumnName);
                Assert.AreEqual("column3", _table.Columns[2].ColumnName);
            }
        }
    }
}