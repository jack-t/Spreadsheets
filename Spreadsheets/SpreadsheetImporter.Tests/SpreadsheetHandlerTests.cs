using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetImporter;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using SpreadsheetImporter.Tests.Stubs;

namespace SpreadsheetImporter.Tests
{
    [TestClass()]
    public class SpreadsheetHandlerTests
    {
        [TestClass()]
        public class PackageTests
        {
            private ExcelPackage package;
            StubSpreadsheetTemplate stub;


            [TestInitialize()]
            public void TestInit()
            {
                package = new ExcelPackage();
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Demo");

                worksheet.Cells["A1"].Value = 123;
                package.Save();
                stub = new StubSpreadsheetTemplate(package.Stream);
            }


            [TestMethod()]
            public void SpreadsheetHandlerTest()
            {
                SpreadsheetHandler handler = new SpreadsheetHandler(stub);

                Assert.IsNotNull(handler.Package);
                Assert.IsNotNull(handler.Package.Workbook);
                Assert.AreEqual(123.0, handler.Package.Workbook.Worksheets["Demo"].Cells["A1"].Value);

                package.Dispose();
            }

            [TestMethod()]
            public void DisposeTest()
            {
                SpreadsheetHandler handler = new SpreadsheetHandler(stub);
                handler.Dispose();
                Assert.IsNull(handler.Package.Stream);
            }
        }

        [TestClass()]
        public class InsertDataTests
        {
            private ExcelPackage package;
            private StubSpreadsheetTemplate stub;


            [TestInitialize()]
            public void TestInit()
            {
                package = new ExcelPackage();
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Demo");

                worksheet.Cells["A1"].Value = "column1";
                worksheet.Cells["A2"].Value = "column2";
                worksheet.Cells["A3"].Value = "column3";
                package.Save();
                stub = new StubSpreadsheetTemplate(package.Stream);
            }

            [TestMethod()]
            public void InsertDataTest()
            {
                SpreadsheetHandler handler = new SpreadsheetHandler(stub);
                DataTable table = new DataTable();
                table.Columns.Add("column1");
                table.Columns.Add("column2");
                table.Columns.Add("column3");
                DataRow row = table.NewRow();
                row.ItemArray = new object[] { 1, 2, 3 };
                table.Rows.Add(row);

                handler.InsertData(new StubExportData { Table = table }, "Demo", 2);

                Assert.AreEqual(1, package.Workbook.Worksheets["Demo"].Cells[2, 1].Value);
                Assert.AreEqual(2, package.Workbook.Worksheets["Demo"].Cells[2, 2].Value);
                Assert.AreEqual(3, package.Workbook.Worksheets["Demo"].Cells[2, 3].Value);
            }

        }
    }
}