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

        private DataTable dataTable;
        private Workbook workbook;
        private static readonly string SheetName = "Sheet1";

        [TestMethod()]
        public void ExportSpreadsheet_With_ExtraDataColumns()
        {
            DefaultSpreadsheetExporter ex = new DefaultSpreadsheetExporter();
            StubExportData data = new StubExportData(dataTable, null);
            StubSpreadsheetTemplate template = new StubSpreadsheetTemplate(workbook);
            MemoryStream stream = new MemoryStream();
            ex.ExportSpreadsheet(data, template, stream);
            Workbook wk = new Workbook();
            wk.LoadFromStream(stream);
            wk.Version = ExcelVersion.Version2013;
            var sheet = wk.Worksheets[SheetName];
            Assert.AreEqual(1d, sheet[2, 1].NumberValue);
            Assert.AreEqual(2d, sheet[2, 2].NumberValue);
            Assert.AreEqual(3d, sheet[2, 4].NumberValue);
            Assert.IsFalse(sheet[2, 3].HasString);
            Assert.AreEqual(1d, sheet[3, 1].NumberValue);
            Assert.AreEqual(2d, sheet[3, 2].NumberValue);
            Assert.IsFalse(sheet[3, 4].HasNumber);
            Assert.IsFalse(sheet[3, 3].HasString);
            Assert.AreEqual(1d, sheet[4, 1].NumberValue);
            Assert.AreEqual(2d, sheet[4, 2].NumberValue);
            Assert.AreEqual(3d, sheet[4, 4].NumberValue);
            Assert.IsFalse(sheet[4, 3].HasString);
            Assert.AreEqual("column1", sheet[1, 1].Text);
            Assert.AreEqual("column2", sheet[1, 2].Text);
            Assert.AreEqual("column3", sheet[1, 4].Text);
            //wk.SaveToFile("./test.xlsx");
        }

        private void addRow(DataTable data, params object[] p)
        {
            var row = data.NewRow();
            row.ItemArray = p;
            data.Rows.Add(row);
        }

        [TestInitialize]
        public void InitTest()
        {
            dataTable = new DataTable();
            dataTable.Columns.Add("column1");
            dataTable.Columns.Add("column2");
            dataTable.Columns.Add("column3");
            dataTable.Columns.Add("column4");

            addRow(dataTable, new object[] { 1, 2, 3, 4 });
            addRow(dataTable, new object[] { 1, 2 });
            for (int i = 0; i < 197; i++)
            {
                addRow(dataTable, new object[] {1, 2, 3, 4});
            }
            workbook = new Workbook();
            //workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets[SheetName].SetCellValue(1, 1, "column1");
            workbook.Worksheets[SheetName][1, 1].Style.Color = Color.DarkRed;
            workbook.Worksheets[SheetName].SetCellValue(1, 2, "column2");
            workbook.Worksheets[SheetName].SetCellValue(1, 4, "column3");
            workbook.Worksheets[SheetName].SetCellValue(1, 5, "column5");
        }
    }
}