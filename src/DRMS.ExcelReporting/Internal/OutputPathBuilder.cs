using System.Globalization;
using System.Text.RegularExpressions;

namespace DRMS.ExcelReporting.Internal;

/// <summary>
/// Builds the folder and file name of a report file from the configured patterns,
/// for example "{TempPath}/ExcelReports" and "{ReportName}_{Timestamp}.xlsx".
/// </summary>
internal static partial class OutputPathBuilder
{
    /// <summary>Placeholders allowed in a folder pattern.</summary>
    public static readonly string[] DirectoryTokens = ["TempPath", "BaseDirectory", "ReportName"];

    /// <summary>Placeholders allowed in a file name pattern.</summary>
    public static readonly string[] FileNameTokens = ["ReportName", "Date", "Time", "Timestamp", "Guid"];

    private const string Extension = ".xlsx";

    // Characters not allowed in file names on Windows, Linux or macOS.
    // A fixed list is used so names made on Linux also work when copied to Windows.
    private static readonly char[] InvalidFileNameChars =
        [.. Path.GetInvalidFileNameChars(), '<', '>', ':', '"', '/', '\\', '|', '?', '*'];

    // Matches {Name} or {Name:format}.
    [GeneratedRegex(@"\{(?<name>[A-Za-z]+)(?::(?<format>[^{}]+))?\}")]
    private static partial Regex TokenRegex();

    /// <summary>Returns placeholders in <paramref name="pattern"/> that are not in <paramref name="allowed"/>.</summary>
    public static IReadOnlyList<string> FindUnknownTokens(string pattern, IEnumerable<string> allowed)
    {
        var allowedSet = new HashSet<string>(allowed, StringComparer.OrdinalIgnoreCase);

        return TokenRegex().Matches(pattern)
            .Select(m => m.Groups["name"].Value)
            .Where(name => !allowedSet.Contains(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>
    /// Turns a folder pattern into a full path. Relative paths are relative to the application folder.
    /// </summary>
    public static string BuildDirectory(string pattern, string reportName)
    {
        var resolved = ReplaceTokens(pattern, (name, _) => name.ToUpperInvariant() switch
        {
            "TEMPPATH" => Path.TrimEndingDirectorySeparator(Path.GetTempPath()),
            "BASEDIRECTORY" => Path.TrimEndingDirectorySeparator(AppContext.BaseDirectory),
            "REPORTNAME" => SanitizeFileName(reportName),
            _ => throw UnknownToken(name, pattern),
        });

        resolved = Environment.ExpandEnvironmentVariables(resolved.Trim());

        if (string.IsNullOrWhiteSpace(resolved))
        {
            throw new InvalidOperationException("The output directory is empty. Set \"ExcelReporting:Defaults:Output:Directory\".");
        }

        return Path.GetFullPath(resolved, AppContext.BaseDirectory);
    }

    /// <summary>
    /// Turns a file name pattern into a safe file name that ends with ".xlsx".
    /// Folder parts and characters that are not allowed in file names are removed.
    /// </summary>
    public static string BuildFileName(string pattern, string reportName, DateTimeOffset now)
    {
        var resolved = ReplaceTokens(pattern, (name, format) => name.ToUpperInvariant() switch
        {
            "REPORTNAME" => reportName,
            "DATE" => now.ToString(format ?? "yyyyMMdd", CultureInfo.InvariantCulture),
            "TIME" => now.ToString(format ?? "HHmmss", CultureInfo.InvariantCulture),
            "TIMESTAMP" => now.ToString(format ?? "yyyyMMdd_HHmmss", CultureInfo.InvariantCulture),
            "GUID" => FormatGuid(format),
            _ => throw UnknownToken(name, pattern),
        });

        var fileName = SanitizeFileName(resolved);

        if (fileName.EndsWith(Extension, StringComparison.OrdinalIgnoreCase))
        {
            fileName = fileName[..^Extension.Length];
        }

        if (string.IsNullOrWhiteSpace(fileName) || fileName.Trim('.', ' ').Length == 0)
        {
            fileName = SanitizeFileName(reportName);
        }

        return fileName + Extension;
    }

    /// <summary>
    /// Replaces every character that is not allowed in a file name with "_".
    /// Slashes are replaced too, so a name can never point to another folder.
    /// </summary>
    public static string SanitizeFileName(string name)
    {
        var chars = name.Trim().ToCharArray();

        for (var i = 0; i < chars.Length; i++)
        {
            if (Array.IndexOf(InvalidFileNameChars, chars[i]) >= 0)
            {
                chars[i] = '_';
            }
        }

        return new string(chars).Trim('.', ' ');
    }

    private static string ReplaceTokens(string pattern, Func<string, string?, string> getValue) =>
        TokenRegex().Replace(pattern, match =>
        {
            var format = match.Groups["format"].Success ? match.Groups["format"].Value : null;
            return getValue(match.Groups["name"].Value, format);
        });

    // "{Guid}" gives 32 characters; "{Guid:8}" gives the first 8; "{Guid:D}" uses .NET's "D" format.
    private static string FormatGuid(string? format)
    {
        var id = Guid.NewGuid();

        if (int.TryParse(format, out var length))
        {
            var text = id.ToString("N");
            return text[..Math.Clamp(length, 1, text.Length)];
        }

        return id.ToString(format ?? "N", CultureInfo.InvariantCulture);
    }

    private static InvalidOperationException UnknownToken(string name, string pattern) =>
        new($"Unknown placeholder '{{{name}}}' in \"{pattern}\".");
}
