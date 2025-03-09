using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ClosedXML.Excel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models;
using ExcelDataReader;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;
    private readonly WebContext _context;

    public HomeController(ILogger<HomeController> logger, WebContext context)
    {
        _logger = logger;
        _context = context;
    }

    public IActionResult Index()
    {
        return View();
    }

    private readonly List<string> knownManagers = new List<string>
    {
        "ÁÀĞÀÒÈÍÑÜÊÈÉ ÂÀÄÈÌ ÌÈÊÎËÀÉÎÂÈ×", "ÃÓÇÅÍÊÎ ĞÓÑËÀÍ ²ÂÀÍÎÂÈ×", "ÊÀÇÀÊÎÂ ŞĞ²É ÂÀËÅÍÒÈÍÎÂÈ×",
        "ÊÓËÈÊ ÂÎËÎÄÈÌÈĞ ÂÀÑÈËÜÎÂÈ×", "Í²Ê²Ò²ÍÀ ªÂÃÅÍ²ß ÌÈÊÎËÀ¯ÂÍÀ",
        "ĞßÆÅÂÀ ÍÀÒÀË²ß ÀÍÀÒÎË²¯ÂÍÀ", "ÑÎÊÎËÎÂÀ ²ĞÈÍÀ ÎËÅÊÑÀÍÄĞ²ÂÍÀ",
        "ÑÓÄÈÊÀ ÂÀËÅĞ²É Â²ÊÒÎĞÎÂÈ×", "ÒÈÑß×ÍÈÉ Â²ÊÒÎĞ ÌÈÊÎËÀÉÎÂÈ×",
        "ÙÓĞ Â²ÊÒÎĞ ÏÅÒĞÎÂÈ×"
    };

    [HttpPost]
    public async Task<IActionResult> UploadFile(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            ViewBag.Message = "Please select a valid file.";
            return View("Index");
        }

        var uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
        if (!Directory.Exists(uploadPath))
        {
            Directory.CreateDirectory(uploadPath);
        }

        var filePath = Path.Combine(uploadPath, file.FileName);

        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }

        // ×èòàºìî Excel ³ äîäàºìî äàí³ â áàçó
        ReadExcelAndSaveToDb(filePath);

        ViewBag.Message = "File uploaded and processed successfully.";
        return View("Index");
    }

    private void ReadExcelAndSaveToDb(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLower();

        if (extension == ".xls" || extension == ".xlsx")
        {
            using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    var dataTable = result.Tables[0]; // First sheet

                    Manager currentManager = null; // Track current manager

                    for (int row = 0; row < dataTable.Rows.Count; row++)
                    {
                        var cellValue = dataTable.Rows[row][0]?.ToString()?.Trim();

                        if (string.IsNullOrEmpty(cellValue)) continue;

                        // Check if the cell contains a known manager's name
                        if (knownManagers.Any(m => cellValue.StartsWith(m)))
                        {
                            // Found a manager's name, create a new manager object
                            currentManager = new Manager
                            {
                                Name = cellValue
                            };
                        }
                        else if (currentManager != null)
                        {
                            // Check if the row contains a date (text starting with "îò")
                            var dateString = ExtractDateFromText(cellValue);
                            if (!string.IsNullOrEmpty(dateString))
                            {
                                // Try parsing the date
                                if (DateTime.TryParse(dateString, out DateTime date))
                                {
                                    // Add the date for the current manager
                                    _context.Managers.Add(new Manager
                                    {
                                        Name = currentManager.Name,
                                        Date = date
                                    });
                                }
                            }
                        }
                    }
                }
                _context.SaveChanges();
            }
        }
        else
        {
            throw new ArgumentException("Unsupported file format.");
        }
    }

    private string ExtractDateFromText(string text)
    {
        // Look for the word "îò" and parse everything after it
        var dateIndex = text?.IndexOf("îò") ?? -1;
        if (dateIndex >= 0)
        {
            var datePart = text.Substring(dateIndex + 3).Trim(); // Cut off "îò"
            return datePart;
        }
        return null;
    }

}
