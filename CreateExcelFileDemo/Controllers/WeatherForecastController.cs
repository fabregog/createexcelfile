using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;

namespace CreateExcelFileDemo.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
    };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");
                worksheet.Cell("A1").Value = "Hello World!";
                worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                workbook.SaveAs("HelloWorld.xlsx");
            }
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }
        [HttpGet("downloadexcel")]
        public IActionResult GetExcel()
        {
            var data = System.IO.File.ReadAllText("Data.json");
            var deserializeData = JsonSerializer.Deserialize<List<QuotationSummary>>(data);
            int row = 2;
            using var workbook = new XLWorkbook();
            using var ms = new MemoryStream();
            var worksheet = workbook.Worksheets.Add("Sample Sheet");
            worksheet.Cell("A1").Value = "Id";
            worksheet.Cell("B1").Value = "Customer Name";
            worksheet.Cell("C1").Value = "Car Brand";
            worksheet.Cell("D1").Value = "Car Model";
            worksheet.Cell("E1").Value = "Anual Payment";


            foreach (var item in deserializeData)
            {
                worksheet.Cell($"A{row}").Value = item.Id;
                worksheet.Cell($"B{row}").Value = item.CustomerName;
                worksheet.Cell($"C{row}").Value = item.CarBrand;
                worksheet.Cell($"D{row}").Value = item.CarModel;
                worksheet.Cell($"E{row}").Value = item.AnualPayment;

                row++;
            }
            worksheet.Cell($"E{row}").FormulaA1 = $"=SUM(E2:E{row - 1})";
            worksheet.Rows().AdjustToContents();
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(ms);
            return File(ms.ToArray(), "application/octet-stream", "hola.xlsx");
        }
        [HttpGet("downloadexcelinserttable")]
        public IActionResult GetExcelInsertTable()
        {
            var data = System.IO.File.ReadAllText("Data.json");
            var deserializeData = JsonSerializer.Deserialize<List<QuotationSummary>>(data);

            using var workbook = new XLWorkbook();
            using var ms = new MemoryStream();

            var worksheet = workbook.Worksheets.Add("Sample Sheet");
            worksheet.Cell(2, 2).InsertTable(deserializeData);

            worksheet.Rows().AdjustToContents();
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(ms);
            return File(ms.ToArray(), "application/octet-stream", "hola.xlsx");
        }
        [HttpGet("downloadexcelinsertdata")]
        public IActionResult GetExcelInsertData()
        {
            var data = System.IO.File.ReadAllText("Data.json");
            var deserializeData = JsonSerializer.Deserialize<List<QuotationSummary>>(data);

            using var workbook = new XLWorkbook();
            using var ms = new MemoryStream();

            var worksheet = workbook.Worksheets.Add("Sample Sheet");
            worksheet.Cell(2, 2).InsertData(deserializeData);

            worksheet.Rows().AdjustToContents();
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(ms);
            return File(ms.ToArray(), "application/octet-stream", "hola.xlsx");
        }

    }
}