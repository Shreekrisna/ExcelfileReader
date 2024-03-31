using ExcelDataReader;
using ExcelfileReader.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace ExcelfileReader.Controllers
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

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult ExcelReader()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ExcelReader(IFormFile file)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            //Upload File
            if (file!=null && file.Length>0)
            {
                var uploadDirectory =$"{Directory.GetCurrentDirectory()}\\wwwroot\\uploads"; 
                if(!Directory.Exists(uploadDirectory))
                {
                    Directory.CreateDirectory(uploadDirectory);
                }
                var filename = Path.GetFileName(file.FileName);

                var filepath = Path.Combine(uploadDirectory,filename);
                using(var stream=new FileStream(filepath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }
                //Read File
                using (var stream = System.IO.File.Open(filepath, FileMode.Open, FileAccess.Read))
                {
                    var excelData = new List<List<object>>();
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {   
                        do
                        {
                            while (reader.Read())
                            {
                                var rowData = new List<object>();//list to hold data for current row
                                for(int column=0 ;column<reader.FieldCount;column++)//This loop iterates over each column in the current row and retrieves the data using reader.GetValue(column)
                                {
                                    rowData.Add(reader.GetValue(column));//adds the value of each cell in the row to the rowData list.
                                }
                                excelData.Add(rowData);//Once all columns in the row are read, the rowData list representing the row is added to the excelData list.
                            }
                        } while (reader.NextResult());
                        ViewBag.ExcelData = excelData;
                        // The result of each spreadsheet is in result.Tables
                    }
                }
            }
            return View();

        }
    }
}
