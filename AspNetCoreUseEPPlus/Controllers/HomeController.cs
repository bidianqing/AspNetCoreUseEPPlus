using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToCollection;
using System;

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

        [HttpPost]
        public string ReadAsync([FromForm] IFormFile file)
        {
            using Stream newStream = file.OpenReadStream();
            using ExcelPackage package = new ExcelPackage(newStream);
            //Get the first worksheet in the workbook
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            // ��������
            var value = worksheet.Cells[1, 1].Value;

            //for (int row = 1; row < 5; row++)
            //{
            //    for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
            //    {
            //        Console.WriteLine(worksheet.Cells[row, i].Value);
            //    }
            //}

            var list = worksheet.Cells["A1:C2"].ToCollectionWithMappings<ReadModel>(row =>
            {
                var model = new ReadModel();

                row.Automap(model);

                model.Age = row.GetValue<int>("Age");

                return model;
            }, new ToCollectionRangeOptions
            {
                HeaderRow = 0,
                ConversionFailureStrategy = ToCollectionConversionFailureStrategy.SetDefaultValue
            });

            return "ok";
        }
    }
}
