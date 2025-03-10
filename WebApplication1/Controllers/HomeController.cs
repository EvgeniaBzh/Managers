using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ClosedXML.Excel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models;
using ExcelDataReader;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
        "������������ ����� �����������", "������� ������ ��������", "������� �в� ������������",
        "����� ��������� ����������", "ͲʲҲ�� ����Ͳ� ����������",
        "������ ����˲� �����˲����", "�������� ����� ��������в���",
        "������ ����в� ²��������", "�������� ²���� �����������",
        "��� ²���� ��������", "�������� �.�"
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

        ReadExcelAndSaveToDb(filePath);

        ViewBag.Message = "File uploaded and processed successfully.";
        return View("Index");
    }

    private readonly Dictionary<string, (string Place, string Purpose, string DetailedInfo)> managerData = new Dictionary<string, (string, string, string)>
{
    { "������������ ����� �����������", ("�� ����� �����", "�������������� � ����.-�������� ������: ���������� �� ������ �� ����������� ����.���������� ������ � ������������� �볺��� �� ����������� ���������� �� ����������; - ���������� ����� �����, ������������� ������������� �������� � ������������� ������ ��� ���������� SHIFT, GPC8, VANOQUAT (��� �������� ���������.", "���  ������������ ����� �����������  \r\n���������������� �����: 2805208433\r\n�/� 26005054308944\r\n�/� 26058054307668\r\n� ���� ��� ��� ������ֲ���� ���� ����������ʻ ���: 328704\r\n����� ������ � ��� 2 544 000 0000 004085 �� 15.07.2013 �. \r\n������� ������� ������� 3 ����� ������ 5%.\r\n������: 67400, ������� ���.,\r\n�������������� �-�, �. ��������, \r\n���. ���������, ���. 85, ��. 1\r\n") },
    { "������� ������ ��������", ("��� ������������, ��� ���������, ��� �³���+�", "��� ��������� ������������� ������: ����� ������� ���� ������������; ���������� ������� �����, ������������� �� ����㳺�, ������ ����, ����, ������ � ������������; �������� �������� ����� �� ��� ������� ��� ������ � ������� ���� ������������; - �������������� � �������-�������� ������: �������������� � ������ ������������� ��������� ��������� ������ �� �����, ������������� �� ���������� ������� �����; �������������� � ������ ����� ��� �������� �� ���������� ������.", "���  ������� ������ ��������\r\n�/� 26006060370261 � �� ����������� \r\n��� 305299, �� 3052306778 \r\n����� ��� �������� ��������� ���  � 2 218 000 0000 001031 �� 02.11.2010 �.\r\n������: 53500, ��������������� ���., ����������� �-�, ��� ��������, ����. ���������, 12\r\n") },
    { "������� �в� ������������", ("��� ��� ����������-ʻ", "�������������� � �������-�������� ������: ���������� �� ������ �� ����������� �������-���������� ������ � ������������� �볺��� �� ����������� ���������� �� ����������; - ���������� ����� �����, ������������� ������������� �������� � ������������� ������ ��� ���������� SHIFT, GPC8, VANOQUAT (��� �������� ���������).", "���  ������� ��� ������������\r\n�/� 26005011000327 � ��� ���������ʻ ��� 325365\r\n���� �� ����� ������ ��� ���������� �������� ���������: 25.09.2007, � 2 471 000 0000 063995 \r\n���������������� ����� 2428501299\r\n������: 62404, �������� ���., ����������  �-�., ���. ��������, ���. �������, ���. 2-�, ��. 67\r\n") },
    { "����� ��������� ����������", ("Գ�� ����������� ���� �����������", "�������������� � �������-�������� ������: ���������� �� ������ �� ����������� �������-���������� ������ � ������������� �볺��� �� ����������� ���������� �� ����������; - ���������� ����� �����, ������������� ������������� �������� � ������������� ������ ��� ���������� SHIFT, GPC8, VANOQUAT (��� �������� ���������).", "���  ����� ��������� ����������\r\n�/� 260581861 � ����������� �� ���������� ����� �������, ������������ ��������, ��� 311528  \r\n���������������� ����� 1821810275\r\n����� ��� �������� ��������� ���-������� ����� � 2 297 000 0000 001485 �� 8 ����� 2008 �. \r\n������: 13502, ����������� ���., ������������� �-�, ���. ��� �������, 36\r\n") },
    { "ͲʲҲ�� ����Ͳ� ����������", ("Գ�� ����������� ���� �����������", "�������������� � �������-�������� ������: �������������� � ������ ������������� ��������� ��������� ������ �� �����, ������������� �� ���������� ������� �����; �������������� � ������ ����� ��� �������� �� ���������� ������; ���������� �� ������ �� ����������� �������-���������� ������ � ������������� �볺��� �� ����������� ���������� �� ����������.", "���  ͳ���� ������ ���������\r\n�/�: 2600847087 � ��� �� ����������� ���� ����ܻ, �. ���\r\n��� 322904\r\n�� 2890109768\r\n������: 07400, ������� ���., �. �������, ���. ��������� ������, 4, ��. 92\r\n") },
    { "������ ����˲� �����˲����", ("��� ����������� ����� ��������, ��� ���������� ���������, ��� ��������", "��� ��������� ������������� ������: ����� ������� ���� ������������; ���������� ������� �����, ������������� �� ����㳺�, ������ ����, ����, ������ � ������������; �������� �������� ����� �� ��� ������� ��� ������ � ������� ���� ������������; - �������������� � �������-�������� ������: �������������� � ������ ������������� ��������� ��������� ������ �� �����, ������������� �� ���������� ������� �����.", "���  ������ ������ �����볿��� \r\n�� 2681213028\r\n��-�� ��� �������� ��������� ���-������� ����� �� 24.10.2007 �., ����� � 2 544 000 0000 002094 \r\n������: 67400, ������� ���., �������������� �-�., �. ��������, ���. �����, ���. 109-�, ��. 18\r\n") },
    { "�������� ����� ��������в���", ("��� ������", "��� ��������� ������������� ������: ����� ������� ���� ������������; ���������� ������� �����, ������������� �� ����㳺�, ������ ����, ����, ������ � ������������; �������� �������� ����� �� ��� ������� ��� ������ � ������� ���� ������������; - �������������� � �������-�������� ������: �������������� � ������ ������������� ��������� ��������� ������ �� �����, ������������� �� ���������� ������� �����; �������������� � ������ ����� ��� �������� �� ���������� ������; ���������� �� ������ �� ����������� �������-���������� ������ � ������������� �볺��� �� ����������� ���������� �� ����������.", "��� �������� ����� ������������ \r\n61174, ��������� ���., ���� �����, ���. �������������, ������� 11, �������� 156\r\n���������������� �����: 3223517141 \r\n�/� 26005052153402\r\n�/�IBAN: UA733515330000026005052153402\r\n���� ������ ��� �������� ���������: 22.12.2017\r\n����� ������: 2 247 000 0000 002313\r\n������� ������� ������� 3 ����� 5%.\r\n") },
    { "������ ����в� ²��������", ("��� ������� ���� ����", "��� ��������� ������������� ������: ����� ������� ���� ������������; ���������� ������� �����, ������������� �� ����㳺�, ������ ����, ����, ������ � ������������; �������������� � ������ ����� ��� �������� �� ���������� ������; - �������������� � �������-�������� ������: �������������� � ������ ������������� ��������� ��������� ������ �� �����, ������������� �� ���������� ������� �����; �������������� � ������ ����� ��� �������� �� ���������� ������.", "���  ������ ������ ³��������  \r\n�� 2709409836\r\n���� ������ ��� �������� ���������: 06.12.2017\r\n����� ������: 2 353 000 0000 038727\r\n������� ������� ������� 3 �����.\r\n������: 09100, ������� ���., �. ���� ������, \r\n���. ��������, ���. 95, ��. 72.\r\n") },
    { "�������� ²���� �����������", ("�� ���������", "��� ��������� ������������� ������: ����� ������� ���� ������������; ���������� ������� �����, ������������� �� ����㳺�, ������ ����, ����, ������ � ������������; �������� �������� ����� �� ��� ������� ��� ������ � ������� ���� ������������; - �������������� � �������-�������� ������: �������������� � ������ ������������� ��������� ��������� ������ �� �����, ������������� �� ���������� ������� �����; �������������� � ������ ����� ��� �������� �� ���������� ������.", "���  �������� ³���� ����������� \r\n�� 2922708036\r\n��-�� ��� �������� ��������� ���-������� ����� ���� �03 � 884993 �� 30.11.2010 �., ����� � 2 170 000 0000 000822\r\n������: 24120, ³������� ���., ����������� �-�, �.����䳿���, ���. ����, 21\r\n") },
    { "��� ²���� ��������", ("��� �³���+�", "�������������� � �������-�������� ������: �������������� � ������ ������������� ��������� ��������� ������ �� �����, ������������� �� ���������� ������� �����; �������������� � ������ ����� ��� �������� �� ���������� ������.", "���  ��� ³���� ��������\r\n��� 2204610657\r\n�/� 2600812615 � ��� ���� �������, \r\n�. �������, ³������� ���., ��� 302247\r\n��-�� ��� ��������� ��� � �.�. � 1696 �� 15 ������� 2004 �.\r\n������: 24400, ³������� ���., ����������� �-�, �. �������, ���. ��� �������, ���. 21\r\n") },
    { "�������� �.�.", ("�� �����-�����", "��� ��������� ������������� ������: ����� ������� ���� ������������; ���������� ������� �����, ������������� �� ����㳺�, ������ ����, ����, ������ � ������������; - �������������� � �������-�������� ������: �������������� � ������ ������������� ��������� ��������� ������ �� �����, ������������� �� ���������� ������� �����; �������������� � ������ ����� ��� �������� �� ���������� ������.", "���  �������� ��������� ³��������\r\n��� 0000000000\r\n�/� UA 863052990000026009031608516 � �� �� �����������\r\n") }
};

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
                    var dataTable = result.Tables[0];
                    Manager currentManager = null;

                    for (int row = 0; row < dataTable.Rows.Count; row++)
                    {
                        var cellValue = dataTable.Rows[row][0]?.ToString()?.Trim();
                        if (string.IsNullOrEmpty(cellValue)) continue;

                        if (managerData.ContainsKey(cellValue))
                        {
                            currentManager = new Manager
                            {
                                Name = cellValue,
                                Place = managerData[cellValue].Place,
                                Purpose = managerData[cellValue].Purpose,
                                DetailedInfo = managerData[cellValue].DetailedInfo
                            };
                        }
                        else if (currentManager != null)
                        {
                            var dateString = ExtractDateFromText(cellValue);
                            if (!string.IsNullOrEmpty(dateString) && DateTime.TryParse(dateString, out DateTime date))
                            {
                                decimal? cost = null;
                                if (decimal.TryParse(dataTable.Rows[row][5]?.ToString()?.Trim(), out decimal parsedCost))
                                {
                                    cost = parsedCost;
                                }

                                var newManager = new Manager
                                {
                                    Name = currentManager.Name,
                                    Date = date,
                                    Cost = cost,
                                    Place = currentManager.Place,
                                    Purpose = currentManager.Purpose,
                                    DetailedInfo = currentManager.DetailedInfo
                                };

                                _context.Managers.Add(newManager);

                                // ��������� GenerateDocFile ��� ������� ������ ���������
                                GenerateDocFile(newManager);
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
        var dateIndex = text?.IndexOf("��") ?? -1;
        if (dateIndex >= 0)
        {
            var datePart = text.Substring(dateIndex + 3).Trim(); // Cut off "��"
            return datePart;
        }
        return null;
    }

    private void GenerateDocFile(Manager manager)
    {
        var docPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "docs");
        if (!Directory.Exists(docPath))
        {
            Directory.CreateDirectory(docPath);
        }

        var fileName = $"{manager.Name}_{manager.Date?.ToString("yyyy-MM-dd")}.docx";
        var filePath = Path.Combine(docPath, fileName);

        using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            // ��������� ���� � ��������� ������
            string formattedDate = manager.Date?.ToString("dd.MM.yyyy �.") ?? "";

            // ������ ��������� � �����
            AddCenteredBoldText(body, "���");
            AddCenteredBoldText(body, "����� - ��������� ���� (������� ������)");
            AddCenteredBoldText(body, $"�� {formattedDate}");

            // ������� ����� ��� 4-�� �����
            string actText = $"��, ������������ ��������� ��� ������, � ���� ��������� ��������� ������, � ������ ����, " +
                             $"�� ����������� ��������� Գ����� ����� � ��������� {manager.Name}, � ������ ����, " +
                             $"������ ��� ��� ��, �� ���������� ���� ������� ������� ������ (����� ��� �������) � ���� " +
                             $"������������ �� �������������: {manager.Place}:";

            // ������ ����� � ������������ �� �������
            AddJustifiedText(body, actText);

            // ������ Purpose, � ��������� ������ �� ������� "-"
            AddPurposeWithLineBreaks(body, manager.Purpose);

            // ������ ������� ����� ��� ����������� Enter
            AddJustifiedText(body, ""); // This creates a blank line before the cost line.

            // ������ ����� ��� �������� ������� ���� (������) ��� ���
            decimal cost = manager.Cost ?? 0;
            string amountInWords = ConvertNumberToWords(cost);
            AddJustifiedText(body, $"\t�������� ������� ���� (������) ������ ��� ��� {manager.Cost},00 ���. ({amountInWords}).");

            // ������ ��� ������� �����
            AddJustifiedText(body, ""); // ������ Enter

            // ������ ����� "������� �������� ���� �� ���� �� �����."
            AddJustifiedText(body, "������� �������� ���� �� ���� �� �����.");

            // ������ ����� "̳��� ���������: ������, �. ���"
            AddJustifiedText(body, "̳��� ���������: ������, �. ���");

            AddJustifiedText(body, ""); // ������ Enter

            // ������ �������
            Table table = new Table();
            TableProperties tblProperties = new TableProperties(new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct });
            table.AppendChild(tblProperties);

            // ������������ ����� "Times New Roman" ��� �񳺿 �������
            TableStyle(table);

            // ������ ������ ����� ��� "³� ���������" �� "³� ���������"
            TableRow row1 = new TableRow();
            AddTableCell(row1, "³� ���������", true);
            AddTableCell(row1, "³� ���������", true);
            table.AppendChild(row1);

            // ������ ������ ����� ��� "�������� �����" �� `manager.Name`, � ������ �������
            TableRow row2 = new TableRow();
            AddTableCell(row2, "�������� �����", true);
            AddTableCell(row2, manager.Name, true);
            table.AppendChild(row2);

            // ������ ����� ����� ��� "��� ������" � ��������� ������� ��� ������� ������
            TableRow row3 = new TableRow();
            AddTableCell(row3, "��� ������", false);
            table.AppendChild(row3);

            // ������ ���� ����� ��� ������� ������
            AddTableRow(table, "������ 25409463");
            AddTableRow(table, "��� 254094626567");
            AddTableRow(table, "����� �������� 100029248");
            AddTableRow(table, "������: 04071, �. ���");
            AddTableRow(table, "���. C������, ���.5, ���� �60");

            // ������ ������� �� ��� ���������
            body.AppendChild(table);

            // ������ DetailedInfo ���������
            AddDetailedInfoRowsSimultaneously(table, manager.DetailedInfo);

            mainPart.Document.Save();
        }
    }

    // ����� ��� ����������� ��������� ����� �� ���� ��������
    private void AddDetailedInfoRowsSimultaneously(Table table, string detailedInfo)
    {
        if (!string.IsNullOrEmpty(detailedInfo))
        {
            // ��������� DetailedInfo �� ����� �� \r\n
            var lines = detailedInfo.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            // ��������� �� ������� ����� �� ������ ��������� � ������ �������
            foreach (var line in lines)
            {
                TableRow row = new TableRow();

                // ��� ������� ������� ������ �������� ������� (�������, ������� ��� � ������)
                AddTableCell(row, "", false); // ������� ������� ��� ������� �������

                // ��� ������� ������� ������ ��������� ����� � DetailedInfo
                AddTableCell(row, line, false); // ����� � DetailedInfo � ������� �������

                table.AppendChild(row);
            }
        }
    }

    // ��������� ����� ��� ������������ ������ "Times New Roman" ��� �񳺿 �������
    private void TableStyle(Table table)
    {
        foreach (var row in table.Elements<TableRow>())
        {
            foreach (var cell in row.Elements<TableCell>())
            {
                var paragraph = cell.Elements<Paragraph>().FirstOrDefault();
                if (paragraph == null) continue;

                var run = paragraph.Elements<Run>().FirstOrDefault();
                if (run == null) continue;

                var runProperties = run.Elements<RunProperties>().FirstOrDefault();
                if (runProperties == null)
                {
                    runProperties = new RunProperties();
                    run.PrependChild(runProperties);
                }

                runProperties.Append(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" });
                runProperties.Append(new FontSize() { Val = "22" });
            }
        }
    }

    // ��������� ����� ��� ��������� ������� � �����
    private void AddTableCell(TableRow row, string text, bool isBold)
    {
        TableCell cell = new TableCell();
        var paragraph = new Paragraph();
        var run = new Run();
        var runProperties = new RunProperties();
        if (isBold)
        {
            runProperties.Append(new Bold());
        }
        runProperties.Append(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" });
        runProperties.Append(new FontSize() { Val = "22" });
        run.Append(runProperties);
        run.Append(new Text(text));
        paragraph.Append(run);
        cell.Append(paragraph);
        row.Append(cell);
    }

    // ��������� ����� ��� ��������� ����� ������� ��� ������� ������
    private void AddTableRow(Table table, string text)
    {
        TableRow row = new TableRow();
        AddTableCell(row, text, false);
        table.AppendChild(row);
    }

    /// <summary>
    /// ���� ����� Purpose, ���������� ���� �� ������ ������� �� �������� ����� ����� �� ������ �����.
    /// </summary>
    private void AddPurposeWithLineBreaks(Body body, string purposeText)
    {
        if (string.IsNullOrEmpty(purposeText)) return;

        var purposeParts = purposeText.Split(new[] { " - " }, StringSplitOptions.None);

        foreach (var part in purposeParts)
        {
            // ������ ����� ������� � ������������ �� �������
            string text = part.Trim(); // �������� ���� ������
            AddJustifiedText(body, $" - {text}");  // ������ ����� ����� ������ �������
        }
    }

    /// <summary>
    /// ���� ������ ����� � ����� ��������� � ������� Times New Roman, ����� 11.
    /// </summary>
    private void AddCenteredBoldText(Body body, string text)
    {
        var paragraph = new Paragraph();
        var run = new Run();
        var runProperties = new RunProperties();

        runProperties.Append(new Bold());
        runProperties.Append(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" });
        runProperties.Append(new FontSize() { Val = "22" });

        run.Append(runProperties);
        run.Append(new Text(text));

        paragraph.Append(run);

        var paragraphProperties = new ParagraphProperties();
        paragraphProperties.Append(new Justification() { Val = JustificationValues.Center });

        paragraph.PrependChild(paragraphProperties);
        body.Append(paragraph);
    }

    /// <summary>
    /// ���� ��������� �����, ��������� �� �������, ����� Times New Roman, ����� 11.
    /// </summary>
    private void AddJustifiedText(Body body, string text)
    {
        var paragraph = new Paragraph();
        var run = new Run();
        var runProperties = new RunProperties();

        runProperties.Append(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" });
        runProperties.Append(new FontSize() { Val = "22" });

        run.Append(runProperties);

        // ������ ������ ��������� ����� �������
        run.Append(new TabChar());
        run.Append(new Text(text));

        paragraph.Append(run);

        var paragraphProperties = new ParagraphProperties();
        paragraphProperties.Append(new Justification() { Val = JustificationValues.Both });

        paragraph.PrependChild(paragraphProperties);
        body.Append(paragraph);
    }

    public static string ConvertNumberToWords(decimal number)
    {
        var ones = new[] { "", "����", "���", "���", "������", "����", "�����", "��", "���", "������", "������", "����������", "����������", "����������", "������������", "����������", "�����������", "���������", "����������", "������������" };
        var tens = new[] { "", "", "��������", "��������", "�����", "��������", "���������", "�������", "��������", "���������" };
        var hundreds = new[] { "", "���", "����", "������", "���������", "������", "�������", "�����", "������", "��������" };
        var thousands = new[] { "", "������", "������", "�����" };

        int wholePart = (int)number;
        int fractionalPart = (int)((number - wholePart) * 100); // ĳ���� ������� �� ������

        var words = new List<string>();

        // ������������ ����� ������
        if (wholePart > 0)
        {
            if (wholePart >= 1000)
            {
                var thousandPart = wholePart / 1000;
                // ��������� ������ (����� ��������: ����, �������, �������)
                if (thousandPart >= 100)
                {
                    words.Add(hundreds[thousandPart / 100]);
                    thousandPart %= 100;
                }
                if (thousandPart >= 20)
                {
                    words.Add(tens[thousandPart / 10]);
                    thousandPart %= 10;
                }
                if (thousandPart > 0)
                {
                    words.Add(ones[thousandPart]);
                }

                // ���������� ���� ������ ��� "������"
                int thousandIndex = (thousandPart % 100 >= 10 && thousandPart % 100 <= 20) ? 3 : (thousandPart % 10 == 1 ? 1 : (thousandPart % 10 >= 2 && thousandPart % 10 <= 4 ? 2 : 3));

                // ����������� ��� "���� ������"
                if (thousandPart == 1)
                {
                    words.Add("����");
                }
                else
                {
                    words.Add(thousands[thousandIndex]);
                }

                wholePart %= 1000;
            }

            // ����� ������, �������, �������
            words.Add(hundreds[wholePart / 100]);
            words.Add(tens[(wholePart % 100) / 10]);
            words.Add(ones[wholePart % 10]);
        }

        // ������������ ������� �������
        string fractionalWords = fractionalPart > 0 ? $"{fractionalPart:D2} ������" : "00 ������";

        // ��'������ ���������
        string result = string.Join(" ", words).Trim() + " ���., " + fractionalWords;
        result = char.ToUpper(result[0]) + result.Substring(1); // ����� ����� ������

        return result;
    }

    private void AddDetailedInfoRows(Table table, string detailedInfo)
    {
        // ��������� DetailedInfo �� ����� �����
        var lines = detailedInfo?.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (lines != null)
        {
            foreach (var line in lines)
            {
                AddTableRow(table, line);
            }
        }
    }

}