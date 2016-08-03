using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Services;
using SpreadsheetImporter;

namespace Site.Controllers
{
    public class HomeController : Controller
    {
        private IFeeRepository _fees;
        private ISpreadsheetManagementService _managementService;

        public HomeController(IFeeRepository fees, ISpreadsheetManagementService spreadsheet)
        {
            _fees = fees;
            _managementService = spreadsheet;
        }

        public IActionResult Index()
        {
            var fees = _fees.GetAll();

            return View("Index", fees);
        }

        public IActionResult CurrentFees()
        {
            var fees = _fees.GetAll();
            return View("_CurrentFees", fees);
        }

        public IActionResult Export()
        {
            using (var mstream = new MemoryStream())
            {
                var stream = _managementService.Download(_fees.GetDataTable());
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(mstream);

                return new FileContentResult(mstream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            }
        }
        
        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
            using (var stream = file.OpenReadStream())
            {
                _managementService.Upload(stream);
            }

            return RedirectToAction("Index");
        }

        public IActionResult Error()
        {
            return View();
        }
    }
}
