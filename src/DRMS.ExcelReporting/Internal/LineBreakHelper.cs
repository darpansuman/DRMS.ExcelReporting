using DRMS.ExcelReporting.Configuration;

namespace DRMS.ExcelReporting.Internal;

/// <summary>
/// Prepares text so that line breaks show correctly inside an Excel cell.
/// </summary>
/// <remarks>
/// Excel uses a single "\n" for a new line inside a cell. Windows line breaks ("\r\n") and
/// old Mac line breaks ("\r") are changed to "\n". Configured separators such as "|" or "&lt;br&gt;"
/// are also changed to "\n".
/// </remarks>
internal static class LineBreakHelper
{
    /// <summary>The new line character Excel uses inside a cell.</summary>
    public const string ExcelNewLine = "\n";

    /// <summary>
    /// Replaces every configured separator and every kind of line break with "\n".
    /// </summary>
    public static string Normalize(string text, LineBreakOptions options)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        var result = text.Replace("\r\n", ExcelNewLine).Replace('\r', '\n');

        foreach (var separator in options.Separators)
        {
            if (!string.IsNullOrEmpty(separator))
            {
                // Example: the separator "\\n" in JSON matches the literal text \n
                // (backslash + n) that some databases store instead of a real new line.
                result = result.Replace(separator, ExcelNewLine);
            }
        }

        return result;
    }

    /// <summary>True when the text contains a new line.</summary>
    public static bool HasLineBreak(string? text) => text is not null && text.Contains('\n');
}
