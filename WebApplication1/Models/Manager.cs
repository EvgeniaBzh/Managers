using System;
using System.Collections.Generic;

namespace WebApplication1.Models;

public partial class Manager
{
    public int Id { get; set; }

    public string? Name { get; set; } = null!;

    public DateTime? Date { get; set; } = null!;

    public string? Purposes { get; set; } = null!;

    public decimal? Cost { get; set; }

    public string? DetailedInfo { get; set; }
}
