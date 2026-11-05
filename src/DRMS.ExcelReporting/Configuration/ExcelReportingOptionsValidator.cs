using Microsoft.Extensions.Options;

namespace DRMS.ExcelReporting.Configuration;

/// <summary>
/// Checks the "ExcelReporting" configuration and reports clear error messages,
/// so mistakes in appsettings.json are found at startup instead of at download time.
/// </summary>
public sealed class ExcelReportingOptionsValidator : IValidateOptions<ExcelReportingOptions>
{
    /// <summary>Excel's limit for sheet (tab) names.</summary>
    public const int MaxSheetNameLength = 31;

    private static readonly char[] InvalidSheetNameChars = [':', '\\', '/', '?', '*', '[', ']'];

    private static readonly string[] AllowedAlignments = ["", "Left", "Center", "Right"];

    /// <inheritdoc />
    public ValidateOptionsResult Validate(string? name, ExcelReportingOptions options)
    {
        var errors = new List<string>();

        if (options.Defaults.MaxRowsPerSheet is < 1 or > 1_048_575)
        {
            errors.Add("Defaults.MaxRowsPerSheet must be between 1 and 1048575.");
        }

        foreach (var (reportName, report) in options.Reports)
        {
            errors.AddRange(ValidateReport(reportName, report));
        }

        return errors.Count == 0 ? ValidateOptionsResult.Success : ValidateOptionsResult.Fail(errors);
    }

    /// <summary>
    /// Checks one report definition and returns a list of problems (empty when it is valid).
    /// </summary>
    public static IReadOnlyList<string> ValidateReport(string reportName, ReportDefinition report)
    {
        var errors = new List<string>();
        var prefix = $"Report '{reportName}'";

        if (report.Sheets.Count == 0)
        {
            errors.Add($"{prefix} must have at least one sheet.");
        }

        var seenSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < report.Sheets.Count; i++)
        {
            var sheet = report.Sheets[i];
            var sheetPrefix = $"{prefix}, sheet #{i + 1} ('{sheet.Name}')";

            if (string.IsNullOrWhiteSpace(sheet.Name))
            {
                errors.Add($"{sheetPrefix}: Name is required.");
            }
            else
            {
                if (sheet.Name.Length > MaxSheetNameLength)
                {
                    errors.Add($"{sheetPrefix}: Name must be {MaxSheetNameLength} characters or fewer.");
                }

                if (sheet.Name.IndexOfAny(InvalidSheetNameChars) >= 0)
                {
                    errors.Add($"{sheetPrefix}: Name must not contain any of these characters: : \\ / ? * [ ]");
                }

                if (!seenSheetNames.Add(sheet.Name))
                {
                    errors.Add($"{sheetPrefix}: Name is used by more than one sheet. Sheet names must be unique.");
                }
            }

            if (sheet.Columns.Count == 0)
            {
                errors.Add($"{sheetPrefix}: must have at least one column.");
            }

            for (var c = 0; c < sheet.Columns.Count; c++)
            {
                var column = sheet.Columns[c];
                var columnPrefix = $"{sheetPrefix}, column #{c + 1} ('{column.GetHeaderText()}')";

                if (string.IsNullOrWhiteSpace(column.Property))
                {
                    errors.Add($"{columnPrefix}: Property is required.");
                }

                if (column.Width is <= 0 or > 255)
                {
                    errors.Add($"{columnPrefix}: Width must be between 0 and 255.");
                }

                if (!AllowedAlignments.Contains(column.Alignment, StringComparer.OrdinalIgnoreCase))
                {
                    errors.Add($"{columnPrefix}: Alignment must be Left, Center or Right.");
                }
            }
        }

        return errors;
    }
}
