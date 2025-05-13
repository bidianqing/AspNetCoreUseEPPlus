using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace AspNetCoreUseEPPlus.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class HomeController : ControllerBase
    {
        private const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public IActionResult Get()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");

            var persons = new List<Person>
                {
                    new Person("John", "Doe", 176, new DateTime(1978, 3, 15)),
                    new Person("Sven", "Svensson", 183, new DateTime(1995, 11, 3)),
                    new Person("Jane", "Doe", 168, new DateTime(1989, 2, 26))
                };
            sheet.Cells["A1"].LoadFromCollection(persons, options =>
            {
                options.PrintHeaders = true;
            });

            var excelData = p.GetAsByteArray();
            var fileName = "MyWorkbook.xlsx";

            return File(excelData, ContentType, fileName);
        }
    }
}
