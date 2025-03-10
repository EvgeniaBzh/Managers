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
        "БАРАТИНСЬКИЙ ВАДИМ МИКОЛАЙОВИЧ", "ГУЗЕНКО РУСЛАН ІВАНОВИЧ", "КАЗАКОВ ЮРІЙ ВАЛЕНТИНОВИЧ",
        "КУЛИК ВОЛОДИМИР ВАСИЛЬОВИЧ", "НІКІТІНА ЄВГЕНІЯ МИКОЛАЇВНА",
        "РЯЖЕВА НАТАЛІЯ АНАТОЛІЇВНА", "СОКОЛОВА ІРИНА ОЛЕКСАНДРІВНА",
        "СУДИКА ВАЛЕРІЙ ВІКТОРОВИЧ", "ТИСЯЧНИЙ ВІКТОР МИКОЛАЙОВИЧ",
        "ЩУР ВІКТОР ПЕТРОВИЧ", "ШЕВЧЕНКО О.В"
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
    { "БАРАТИНСЬКИЙ ВАДИМ МИКОЛАЙОВИЧ", ("ПП «Абла Центр»", "консультування з наук.-технічних питань: організація та нагляд за проведенням наук.виробничих дослідів в господарствах клієнтів та опрацювання результатів їх проведення; - обстеження стану стада, дезінфікування тваринницьких приміщень з використанням засобів для дезінфекції SHIFT, GPC8, VANOQUAT (для навчання персоналу.", "ФОП  Баратинський Вадим Миколайович  \r\nІдентифікаційний номер: 2805208433\r\nр/р 26005054308944\r\nр/р 26058054307668\r\nв Южне ГРУ ПАТ «КОМЕРЦІЙНИЙ БАНК «ПРИВАТБАНК» МФО: 328704\r\nНомер запису в ЄДР 2 544 000 0000 004085 від 15.07.2013 р. \r\nПлатник єдиного податку 3 група ставка 5%.\r\nАдреса: 67400, Одеська обл.,\r\nРоздільнянський р-н, м. Роздільна, \r\nвул. Димитрова, буд. 85, кв. 1\r\n") },
    { "ГУЗЕНКО РУСЛАН ІВАНОВИЧ", ("ТОВ «Агромайстер», ТДВ «Агронива», ТОВ «Віола+»", "для підвищення продуктивності тварин: аналіз кормової бази господарства; визначення раціону кормів, збалансованих за енергією, вмістом білків, жирів, вітамінів і мікроелементів; розробка рецептур кормів на базі преміксів ТОВ «Текро» і кормової бази господарства; - консультування з науково-технічних питань: консультування з питань удосконалення технологій утримання тварин та птиці, удосконалення та оптимізація раціонів годівлі; консультування з підбору кормів для селекції та відтворення тварин.", "ФОП  Гузенко Руслан Іванович\r\nР/р 26006060370261 в КБ «Приватбанк» \r\nМФО 305299, ІН 3052306778 \r\nЗапис про державну реєстрацію ФОП  № 2 218 000 0000 001031 від 02.11.2010 р.\r\nАдреса: 53500, Дніпропетровська обл., Томаківський р-н, смт Томаківка, пров. Робітничий, 12\r\n") },
    { "КАЗАКОВ ЮРІЙ ВАЛЕНТИНОВИЧ", ("ТОВ «АВ «Агроцентр-К»", "консультування з науково-технічних питань: організація та нагляд за проведенням науково-виробничих дослідів в господарствах клієнтів та опрацювання результатів їх проведення; - обстеження стану стада, дезінфікування тваринницьких приміщень з використанням засобів для дезінфекції SHIFT, GPC8, VANOQUAT (для навчання персоналу).", "ФОП  Казаков Юрій Валентинович\r\nр/р 26005011000327 в ПАТ «КРЕДОБАНК» МФО 325365\r\nДата та номер запису про проведення державної реєстрації: 25.09.2007, № 2 471 000 0000 063995 \r\nІдентифікаційний номер 2428501299\r\nАдреса: 62404, Харківска обл., Харківський  р-н., смт. Кулиничі, вул. Ювілейна, буд. 2-В, кв. 67\r\n") },
    { "КУЛИК ВОЛОДИМИР ВАСИЛЬОВИЧ", ("Філія «Броварська» ПНВК «Інтербізнес»", "консультування з науково-технічних питань: організація та нагляд за проведенням науково-виробничих дослідів в господарствах клієнтів та опрацювання результатів їх проведення; - обстеження стану стада, дезінфікування тваринницьких приміщень з використанням засобів для дезінфекції SHIFT, GPC8, VANOQUAT (для навчання персоналу).", "ФОП  Кулик Володимир Васильович\r\nр/р 260581861 в Житомирській ОД Райффайзен банку «Аваль», Попільнянське відділення, МФО 311528  \r\nІдентифікаційний номер 1821810275\r\nЗапис про державну реєстрацію СПД-фізичної особи № 2 297 000 0000 001485 від 8 липня 2008 р. \r\nАдреса: 13502, Житомирська обл., Попільнянський р-н, вул. Лесі Українки, 36\r\n") },
    { "НІКІТІНА ЄВГЕНІЯ МИКОЛАЇВНА", ("Філія «Броварська» ПНВК «Інтербізнес»", "консультування з науково-технічних питань: консультування з питань удосконалення технологій утримання тварин та птиці, удосконалення та оптимізація раціонів годівлі; консультування з підбору кормів для селекції та відтворення тварин; організація та нагляд за проведенням науково-виробничих дослідів в господарствах клієнтів та опрацювання результатів їх проведення.", "ФОП  Нікітіна Євгенія Миколаївна\r\nр/р: 2600847087 в КРД АТ «РАЙФФАЙЗЕН БАНК АВАЛЬ», м. Київ\r\nМФО 322904\r\nІН 2890109768\r\nАдреса: 07400, Київська обл., м. Бровари, вул. Симоненка Василя, 4, кв. 92\r\n") },
    { "РЯЖЕВА НАТАЛІЯ АНАТОЛІЇВНА", ("ТОВ «Чернігівська м’ясна компанія», ТОВ «Українська вольниця», ПАП «Фортуна»", "для підвищення продуктивності тварин: аналіз кормової бази господарства; визначення раціону кормів, збалансованих за енергією, вмістом білків, жирів, вітамінів і мікроелементів; розробка рецептур кормів на базі преміксів ТОВ «Текро» і кормової бази господарства; - консультування з науково-технічних питань: консультування з питань удосконалення технологій утримання тварин та птиці, удосконалення та оптимізація раціонів годівлі.", "ФОП  Ряжева Наталія Анатоліївна \r\nІН 2681213028\r\nСв-во про державну реєстрацію СПД-фізичної особи від 24.10.2007 р., запис № 2 544 000 0000 002094 \r\nАдреса: 67400, Одеська обл., Роздільнянський р-н., м. Роздільна, вул. Леніна, буд. 109-А, кв. 18\r\n") },
    { "СОКОЛОВА ІРИНА ОЛЕКСАНДРІВНА", ("ТДВ «Колос»", "для підвищення продуктивності тварин: аналіз кормової бази господарства; визначення раціону кормів, збалансованих за енергією, вмістом білків, жирів, вітамінів і мікроелементів; розробка рецептур кормів на базі преміксів ТОВ «Текро» і кормової бази господарства; - консультування з науково-технічних питань: консультування з питань удосконалення технологій утримання тварин та птиці, удосконалення та оптимізація раціонів годівлі; консультування з підбору кормів для селекції та відтворення тварин; організація та нагляд за проведенням науково-виробничих дослідів в господарствах клієнтів та опрацювання результатів їх проведення.", "ФОП Соколова Ірина Олександрівна \r\n61174, Харківська обл., місто Харків, вул. Домобудівельна, будинок 11, квартира 156\r\nІдентифікаційний номер: 3223517141 \r\nр/р 26005052153402\r\nр/рIBAN: UA733515330000026005052153402\r\nДата запису про державну реєстрацію: 22.12.2017\r\nНомер запису: 2 247 000 0000 002313\r\nПлатник єдиного податку 3 група 5%.\r\n") },
    { "СУДИКА ВАЛЕРІЙ ВІКТОРОВИЧ", ("ТОВ «Золота Роса Агро»", "для підвищення продуктивності тварин: аналіз кормової бази господарства; визначення раціону кормів, збалансованих за енергією, вмістом білків, жирів, вітамінів і мікроелементів; консультування з підбору кормів для селекції та відтворення тварин; - консультування з науково-технічних питань: консультування з питань удосконалення технологій утримання тварин та птиці, удосконалення та оптимізація раціонів годівлі; консультування з підбору кормів для селекції та відтворення тварин.", "ФОП  Судика Валерій Вікторович  \r\nІН 2709409836\r\nДата запису про державну реєстрацію: 06.12.2017\r\nНомер запису: 2 353 000 0000 038727\r\nПлатник єдиного податку 3 група.\r\nАдреса: 09100, Київська обл., м. Біла Церква, \r\nвул. Шевченка, буд. 95, кв. 72.\r\n") },
    { "ТИСЯЧНИЙ ВІКТОР МИКОЛАЙОВИЧ", ("ФГ «Плантера»", "для підвищення продуктивності тварин: аналіз кормової бази господарства; визначення раціону кормів, збалансованих за енергією, вмістом білків, жирів, вітамінів і мікроелементів; розробка рецептур кормів на базі преміксів ТОВ «Текро» і кормової бази господарства; - консультування з науково-технічних питань: консультування з питань удосконалення технологій утримання тварин та птиці, удосконалення та оптимізація раціонів годівлі; консультування з підбору кормів для селекції та відтворення тварин.", "ФОП  Тисячний Віктор Миколайович \r\nІН 2922708036\r\nСв-во про державну реєстрацію СПД-фізичної особи Серія В03 № 884993 від 30.11.2010 р., запис № 2 170 000 0000 000822\r\nАдреса: 24120, Вінницька обл., Чернівецький р-н, с.Володіївці, вул. Миру, 21\r\n") },
    { "ЩУР ВІКТОР ПЕТРОВИЧ", ("ТОВ «Віола+»", "консультування з науково-технічних питань: консультування з питань удосконалення технологій утримання тварин та птиці, удосконалення та оптимізація раціонів годівлі; консультування з підбору кормів для селекції та відтворення тварин.", "ФОП  Щур Віктор Петрович\r\nІПН 2204610657\r\nР/р 2600812615 в ВОД АППБ “Аваль”, \r\nм. Бершадь, Вінницька обл., МФО 302247\r\nСв-во про реєстрацію СПД – ф.о. № 1696 від 15 березня 2004 р.\r\nАдреса: 24400, Вінницька обл., Бершадський р-н, м. Бершадь, вул. Лесі Українки, буд. 21\r\n") },
    { "ШЕВЧЕНКО О.В.", ("ПП «Абла-Центр»", "для підвищення продуктивності тварин: аналіз кормової бази господарства; визначення раціону кормів, збалансованих за енергією, вмістом білків, жирів, вітамінів і мікроелементів; - консультування з науково-технічних питань: консультування з питань удосконалення технологій утримання тварин та птиці, удосконалення та оптимізація раціонів годівлі; консультування з підбору кормів для селекції та відтворення тварин.", "ФОП  Шевченко Олександр Вікторович\r\nІПН 0000000000\r\nР/р UA 863052990000026009031608516 в АТ КБ «ПриватБанк»\r\n") }
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

                                // Викликаємо GenerateDocFile для кожного нового менеджера
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
        var dateIndex = text?.IndexOf("от") ?? -1;
        if (dateIndex >= 0)
        {
            var datePart = text.Substring(dateIndex + 3).Trim(); // Cut off "от"
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

            // Форматуємо дату у потрібному вигляді
            string formattedDate = manager.Date?.ToString("dd.MM.yyyy р.") ?? "";

            // Додаємо заголовок у центрі
            AddCenteredBoldText(body, "Акт");
            AddCenteredBoldText(body, "Здачі - прийняття робіт (надання послуг)");
            AddCenteredBoldText(body, $"від {formattedDate}");

            // Формуємо текст для 4-го рядка
            string actText = $"Ми, представники Замовника ТОВ «Текро», в особі директора Вондроуша Йозефа, з одного боку, " +
                             $"та представник Виконавця Фізична особа – підприємець {manager.Name}, з іншого боку, " +
                             $"склали акт про те, що Виконавцем були виконані наступні роботи (надані такі послуги) у сфері " +
                             $"тваринництва по господарствах: {manager.Place}:";

            // Додаємо текст з вирівнюванням за шириною
            AddJustifiedText(body, actText);

            // Додаємо Purpose, з розбиттям тексту по символу "-"
            AddPurposeWithLineBreaks(body, manager.Purpose);

            // Додаємо порожній рядок для додаткового Enter
            AddJustifiedText(body, ""); // This creates a blank line before the cost line.

            // Додаємо текст про загальну вартість робіт (послуг) без ПДВ
            decimal cost = manager.Cost ?? 0;
            string amountInWords = ConvertNumberToWords(cost);
            AddJustifiedText(body, $"\tЗагальна вартість робіт (послуг) склала без ПДВ {manager.Cost},00 грн. ({amountInWords}).");

            // Додаємо три розриви рядків
            AddJustifiedText(body, ""); // Перший Enter

            // Додаємо текст "Сторони претензій одна до одної не мають."
            AddJustifiedText(body, "Сторони претензій одна до одної не мають.");

            // Додаємо текст "Місце складання: Україна, м. Київ"
            AddJustifiedText(body, "Місце складання: Україна, м. Київ");

            AddJustifiedText(body, ""); // Перший Enter

            // Додаємо таблицю
            Table table = new Table();
            TableProperties tblProperties = new TableProperties(new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct });
            table.AppendChild(tblProperties);

            // Встановлюємо шрифт "Times New Roman" для всієї таблиці
            TableStyle(table);

            // Додаємо перший рядок для "Від Замовника" та "Від Виконавця"
            TableRow row1 = new TableRow();
            AddTableCell(row1, "Від Замовника", true);
            AddTableCell(row1, "Від Виконавця", true);
            table.AppendChild(row1);

            // Додаємо другий рядок для "Вондроуш Йозеф" та `manager.Name`, з жирним шрифтом
            TableRow row2 = new TableRow();
            AddTableCell(row2, "Вондроуш Йозеф", true);
            AddTableCell(row2, manager.Name, true);
            table.AppendChild(row2);

            // Додаємо третій рядок для "ТОВ «Текро»" і наступних деталей без жирного шрифта
            TableRow row3 = new TableRow();
            AddTableCell(row3, "ТОВ «Текро»", false);
            table.AppendChild(row3);

            // Додаємо інші рядки без жирного шрифта
            AddTableRow(table, "ЄДРПОУ 25409463");
            AddTableRow(table, "ІПН 254094626567");
            AddTableRow(table, "номер свідоцтва 100029248");
            AddTableRow(table, "Адреса: 04071, м. Київ");
            AddTableRow(table, "вул. Cпаська, буд.5, офіс №60");

            // Додаємо таблицю до тіла документа
            body.AppendChild(table);

            // Додаємо DetailedInfo одночасно
            AddDetailedInfoRowsSimultaneously(table, manager.DetailedInfo);

            mainPart.Document.Save();
        }
    }

    // Метод для одночасного додавання рядків до обох стовпців
    private void AddDetailedInfoRowsSimultaneously(Table table, string detailedInfo)
    {
        if (!string.IsNullOrEmpty(detailedInfo))
        {
            // Розбиваємо DetailedInfo на рядки за \r\n
            var lines = detailedInfo.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            // Проходимо по кожному рядку та додаємо одночасно в обидва стовпці
            foreach (var line in lines)
            {
                TableRow row = new TableRow();

                // Для першого стовпця додаємо відповідну клітинку (можливо, порожню або з даними)
                AddTableCell(row, "", false); // Порожня клітинка для першого стовпця

                // Для другого стовпця додаємо відповідний рядок з DetailedInfo
                AddTableCell(row, line, false); // Рядок з DetailedInfo у другому стовпці

                table.AppendChild(row);
            }
        }
    }

    // Допоміжний метод для встановлення шрифта "Times New Roman" для всієї таблиці
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

    // Допоміжний метод для додавання клітинки в рядок
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

    // Допоміжний метод для додавання рядка таблиці без жирного шрифта
    private void AddTableRow(Table table, string text)
    {
        TableRow row = new TableRow();
        AddTableCell(row, text, false);
        table.AppendChild(row);
    }

    /// <summary>
    /// Додає текст Purpose, розбиваючи його за кожним дефісом та виводячи кожен пункт на новому рядку.
    /// </summary>
    private void AddPurposeWithLineBreaks(Body body, string purposeText)
    {
        if (string.IsNullOrEmpty(purposeText)) return;

        var purposeParts = purposeText.Split(new[] { " - " }, StringSplitOptions.None);

        foreach (var part in purposeParts)
        {
            // Додаємо кожну частину з вирівнюванням за шириною
            string text = part.Trim(); // Видалити зайві пробіли
            AddJustifiedText(body, $" - {text}");  // Додаємо дефіс перед кожним пунктом
        }
    }

    /// <summary>
    /// Додає жирний текст у центрі документа з шрифтом Times New Roman, розмір 11.
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
    /// Додає звичайний текст, вирівняний за шириною, шрифт Times New Roman, розмір 11.
    /// </summary>
    private void AddJustifiedText(Body body, string text)
    {
        var paragraph = new Paragraph();
        var run = new Run();
        var runProperties = new RunProperties();

        runProperties.Append(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" });
        runProperties.Append(new FontSize() { Val = "22" });

        run.Append(runProperties);

        // Додаємо символ табуляції перед текстом
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
        var ones = new[] { "", "один", "два", "три", "чотири", "п’ять", "шість", "сім", "вісім", "дев’ять", "десять", "одинадцять", "дванадцять", "тринадцять", "чотирнадцять", "п’ятнадцять", "шістнадцять", "сімнадцять", "вісімнадцять", "дев’ятнадцять" };
        var tens = new[] { "", "", "двадцять", "тридцять", "сорок", "п’ятдесят", "шістдесят", "сімдесят", "вісімдесят", "дев’яносто" };
        var hundreds = new[] { "", "сто", "двісті", "триста", "чотириста", "п’ятсот", "шістсот", "сімсот", "вісімсот", "дев’ятсот" };
        var thousands = new[] { "", "тисяча", "тисячі", "тисяч" };

        int wholePart = (int)number;
        int fractionalPart = (int)((number - wholePart) * 100); // Ділимо залишок на копійки

        var words = new List<string>();

        // Перетворення цілих частин
        if (wholePart > 0)
        {
            if (wholePart >= 1000)
            {
                var thousandPart = wholePart / 1000;
                // Розбираємо тисячі (перед тисячами: сотні, десятки, одиниці)
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

                // Правильний вибір відмінка для "тисяча"
                int thousandIndex = (thousandPart % 100 >= 10 && thousandPart % 100 <= 20) ? 3 : (thousandPart % 10 == 1 ? 1 : (thousandPart % 10 >= 2 && thousandPart % 10 <= 4 ? 2 : 3));

                // Виправлення для "одна тисяча"
                if (thousandPart == 1)
                {
                    words.Add("одна");
                }
                else
                {
                    words.Add(thousands[thousandIndex]);
                }

                wholePart %= 1000;
            }

            // Розбір сотень, десятків, одиниць
            words.Add(hundreds[wholePart / 100]);
            words.Add(tens[(wholePart % 100) / 10]);
            words.Add(ones[wholePart % 10]);
        }

        // Перетворення дробової частини
        string fractionalWords = fractionalPart > 0 ? $"{fractionalPart:D2} копійок" : "00 копійок";

        // Об'єднуємо результат
        string result = string.Join(" ", words).Trim() + " грн., " + fractionalWords;
        result = char.ToUpper(result[0]) + result.Substring(1); // Перша буква велика

        return result;
    }

    private void AddDetailedInfoRows(Table table, string detailedInfo)
    {
        // Розбиваємо DetailedInfo на окремі рядки
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