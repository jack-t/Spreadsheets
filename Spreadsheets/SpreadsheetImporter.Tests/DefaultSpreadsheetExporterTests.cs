using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using SpreadsheetImporter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Spire.Xls;
using SpreadsheetImporter.Tests.Stubs;

namespace SpreadsheetImporter.Tests
{
    [TestClass()]
    public class DefaultSpreadsheetExporterTests
    {

        private static DataTable _dataTable;
        private static Workbook _workbook;
        private static readonly string SheetName = "Sheet1";
        private static DefaultSpreadsheetExporter _defaultSpreadsheetExporter;
        private static StubExportData _data;
        private static StubSpreadsheetTemplate _template;
        private static Workbook _wk;

        [TestMethod()]
        public void ExportSpreadsheet_With_ExtraDataColumns()
        {
            var sheet = _wk.Worksheets[SheetName];
            Assert.AreEqual("column1", sheet[1, 1].Text);
            Assert.AreEqual("column2", sheet[1, 2].Text);
            Assert.AreEqual("column3", sheet[1, 4].Text);
            //wk.SaveToFile("./test.xlsx");
        }

        [TestMethod()]
        public void ExportSpreadsheet_Row4()
        {
            var sheet = _wk.Worksheets[SheetName];
            Assert.AreEqual(1d, sheet[4, 1].NumberValue);
            Assert.AreEqual(2d, sheet[4, 2].NumberValue);
            Assert.AreEqual(3d, sheet[4, 4].NumberValue);
            Assert.IsFalse(sheet[4, 3].HasString);
        }
        [TestMethod()]
        public void ExportSpreadsheet_Row3()
        {
            var sheet = _wk.Worksheets[SheetName];
            Assert.AreEqual(1d, sheet[3, 1].NumberValue);
            Assert.AreEqual(2d, sheet[3, 2].NumberValue);
            Assert.IsFalse(sheet[3, 4].HasNumber);
            Assert.IsFalse(sheet[3, 3].HasString);
        }
        [TestMethod()]
        public void ExportSpreadsheet_Row2()
        {
            var sheet = _wk.Worksheets[SheetName];
            Assert.AreEqual(1d, sheet[2, 1].NumberValue);
            Assert.AreEqual(2d, sheet[2, 2].NumberValue);
            Assert.AreEqual(3d, sheet[2, 4].NumberValue);
            Assert.IsFalse(sheet[2, 3].HasString);
        }

        private static void addRow(DataTable data, params object[] p)
        {
            var row = data.NewRow();
            row.ItemArray = p;
            data.Rows.Add(row);
        }

        [ClassInitialize]
        public static void InitTest(TestContext cnt)
        {
            _dataTable = new DataTable();
            _dataTable.Columns.Add("column1");
            _dataTable.Columns.Add("column2");
            _dataTable.Columns.Add("column3");
            _dataTable.Columns.Add("column4");

            addRow(_dataTable, new object[] { 1, 2, 3, 4 });
            addRow(_dataTable, new object[] { 1, 2 });
            for (int i = 0; i < 197; i++)
            {
                addRow(_dataTable, new object[] { 1, 2, 3, 4 });
            }
            _workbook = new Workbook();
            //workbook.Worksheets.Add("Sheet1");
            _workbook.Worksheets[SheetName].SetCellValue(1, 1, "column1");
            _workbook.Worksheets[SheetName][1, 1].Style.Color = Color.DarkRed;
            _workbook.Worksheets[SheetName].SetCellValue(1, 2, "column2");
            _workbook.Worksheets[SheetName].SetCellValue(1, 4, "column3");
            _workbook.Worksheets[SheetName].SetCellValue(1, 5, "column5");
            _defaultSpreadsheetExporter = new DefaultSpreadsheetExporter();
            _data = new StubExportData(_dataTable, null);
            _template = new StubSpreadsheetTemplate(_workbook, new Dictionary<string, int>
            {
                ["column1"] = 1,
                ["column2"] = 2,
                ["column3"] = 4
            }, null);

            using (MemoryStream stream = new MemoryStream())
            {
                _wk = new Workbook();
                _defaultSpreadsheetExporter.ExportSpreadsheet(_data, _template, stream);
                _wk.LoadFromStream(stream);
                _wk.Version = ExcelVersion.Version2013;
            }
        }
    }
}