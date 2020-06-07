using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using AspNetCore.ExcelProject_UsingEPPlus.Models;
using OfficeOpenXml;

namespace AspNetCore.ExcelProject_UsingEPPlus.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult BringExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            var excelBlank = excelPackage.Workbook.Worksheets.Add("WorkSheet-1");
            //excelBlank.Cells[1, 1].Value = "Name";
            //excelBlank.Cells[1, 2].Value = "Surname";


            //excelBlank.Cells[2, 1].Value = "Kübay";
            //excelBlank.Cells[2, 2].Value = "TELCEKEN";

            excelBlank.Cells["A1"].LoadFromCollection(new List<Customer>
            {
                new Customer{Id=1,Name="Kübay"},
                new Customer{Id=2,Name="Mehmet"},
            },true,OfficeOpenXml.Table.TableStyles.Light15);
            var bytes = excelPackage.GetAsByteArray(); // It provides to handle the data in excel as a btye array.
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Guid.NewGuid() + "" + ".xlsx");

        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }

    class Customer
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
}
