using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetImporter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using SpreadsheetImporter.Tests.Stubs;

namespace SpreadsheetImporter.Tests
{
    [TestClass()]
    public class SpreadsheetServiceTests
    {
        [TestMethod()]
        public void FindGuidTest()
        {
            Guid guid = Guid.NewGuid();
            Workbook workbook = new Workbook();
            ISpreadsheetTemplate template = new FindGuidTestTemplateStub();
            workbook.Worksheets[template.DataSheetName][template.GuidCell].Value = guid.ToString();
            SpreadsheetService service = new SpreadsheetService(null, null, null, template);
            var g = service.FindGuid(workbook);
            Assert.AreEqual(guid, g);
        }
    }
}