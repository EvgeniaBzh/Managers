using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace WebApplication1.Models;

public partial class Manager
{
    public int Id { get; set; }

    public string? Name { get; set; } = null!;

    public DateTime? Date { get; set; } = null!;

    public string? Purpose { get; set; } = null!;

    public string? Place { get; set; } = null!;

    public decimal? Cost { get; set; }

    public string? DetailedInfo { get; set; }
}