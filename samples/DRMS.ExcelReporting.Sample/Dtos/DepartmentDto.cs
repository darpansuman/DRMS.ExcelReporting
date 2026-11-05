namespace DRMS.ExcelReporting.Sample.Dtos;

/// <summary>
/// Second sample DTO, shown on its own sheet ("Departments") in the same Excel file.
/// </summary>
public sealed class DepartmentDto
{
    /// <summary>Department code, for example "FIN".</summary>
    public string Code { get; init; } = string.Empty;

    /// <summary>Department name.</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Manager's name.</summary>
    public string Manager { get; init; } = string.Empty;

    /// <summary>Number of people.</summary>
    public int HeadCount { get; init; }

    /// <summary>Yearly budget.</summary>
    public decimal Budget { get; init; }

    /// <summary>
    /// Office locations separated by ";" (for example "Pune;Mumbai").
    /// The column's line break settings put each location on its own line.
    /// </summary>
    public string Locations { get; init; } = string.Empty;
}
