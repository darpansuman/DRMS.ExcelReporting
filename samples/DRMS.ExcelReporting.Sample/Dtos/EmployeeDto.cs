namespace DRMS.ExcelReporting.Sample.Dtos;

/// <summary>
/// Sample DTO. Each property is mapped to an Excel column in appsettings.json
/// ("ExcelReporting:Reports:EmployeeReport"). The comments show which data type is used.
/// </summary>
public sealed class EmployeeDto
{
    /// <summary>Integer column.</summary>
    public int Id { get; init; }

    /// <summary>Text column. Keeps leading zeros such as "00042".</summary>
    public string EmployeeCode { get; init; } = string.Empty;

    /// <summary>Text column.</summary>
    public string FullName { get; init; } = string.Empty;

    /// <summary>Hyperlink column (mailto: link).</summary>
    public string Email { get; init; } = string.Empty;

    /// <summary>Nested object. Mapped with dotted paths such as "Address.City".</summary>
    public AddressDto? Address { get; init; }

    /// <summary>Date column.</summary>
    public DateOnly JoiningDate { get; init; }

    /// <summary>DateTime column. Null shows the column's DefaultValue.</summary>
    public DateTime? LastLogin { get; init; }

    /// <summary>Time column.</summary>
    public TimeOnly ShiftStart { get; init; }

    /// <summary>Currency column.</summary>
    public decimal Salary { get; init; }

    /// <summary>Percentage column. 0.1 is shown as 10.00%.</summary>
    public double BonusRate { get; init; }

    /// <summary>Text column with ValueMappings True → "Yes", False → "No".</summary>
    public bool IsActive { get; init; }

    /// <summary>Status code. ValueMappings turn "A" into "Active", and so on.</summary>
    public string StatusCode { get; init; } = string.Empty;

    /// <summary>Enum value. Written as its name ("FullTime").</summary>
    public EmploymentType EmploymentType { get; init; }

    /// <summary>A list. Each skill is written on its own line in the same cell.</summary>
    public List<string> Skills { get; init; } = [];

    /// <summary>
    /// Free text that may contain real line breaks ("\r\n") and "|" markers.
    /// Both become new lines in Excel.
    /// </summary>
    public string? Notes { get; init; }
}

/// <summary>Nested address object.</summary>
public sealed class AddressDto
{
    /// <summary>Street, may span several lines.</summary>
    public string Street { get; init; } = string.Empty;

    /// <summary>City.</summary>
    public string City { get; init; } = string.Empty;

    /// <summary>Country.</summary>
    public string Country { get; init; } = string.Empty;
}

/// <summary>Kind of employment.</summary>
public enum EmploymentType
{
    /// <summary>Full-time employee.</summary>
    FullTime,

    /// <summary>Part-time employee.</summary>
    PartTime,

    /// <summary>External contractor.</summary>
    Contractor,
}
