using System.Collections;
using System.Globalization;
using ClosedXML.Excel;
using DRMS.ExcelReporting.Configuration;

namespace DRMS.ExcelReporting.Internal;

/// <summary>
/// What should be written into one cell: the value, its number format and an optional link.
/// </summary>
internal readonly record struct CellContent(XLCellValue Value, string? NumberFormat = null, Uri? Hyperlink = null)
{
    public static CellContent Blank => new(ClosedXML.Excel.Blank.Value);

    /// <summary>True when the value is text that contains a line break.</summary>
    public bool HasLineBreak => Value.IsText && LineBreakHelper.HasLineBreak(Value.GetText());
}

/// <summary>
/// Converts a raw DTO value into an Excel cell value, following the column's settings.
/// </summary>
/// <remarks>
/// The steps are always the same, in this order:
/// 1. Lists are joined into one text (each item on its own line by default).
/// 2. ValueMappings replace raw values with friendly text.
/// 3. Empty values get the DefaultValue (or stay blank).
/// 4. The value is converted to the column's DataType and the number format is chosen.
/// If a value cannot be converted (for example "abc" in a Number column), it is written as text
/// and a warning is returned, so one bad value never breaks the whole report.
/// Text from ValueMappings or DefaultValue (for example "Never" in a DateTime column) is written as
/// text without a warning, because it was chosen on purpose.
/// </remarks>
internal sealed class CellValueConverter
{
    /// <summary>Excel's maximum number of characters in one cell.</summary>
    public const int MaxCellTextLength = 32_767;

    private readonly ReportDefaults _defaults;
    private readonly CultureInfo _culture;

    public CellValueConverter(ReportDefaults defaults)
    {
        _defaults = defaults;
        _culture = string.IsNullOrWhiteSpace(defaults.Culture)
            ? CultureInfo.InvariantCulture
            : CultureInfo.GetCultureInfo(defaults.Culture);
    }

    /// <summary>
    /// Converts <paramref name="raw"/> for the given column.
    /// <paramref name="warning"/> is set when the value could not be converted to the requested type.
    /// </summary>
    public CellContent Convert(object? raw, ColumnDefinition column, out string? warning)
    {
        warning = null;
        var lineBreaks = column.LineBreaks ?? _defaults.LineBreaks;

        // True when the value is text written on purpose in configuration (a mapping or default value).
        // Such text is never reported as a conversion problem.
        var isConfiguredText = false;

        // Step 1: a list becomes one text value, joined with the collection separator.
        if (raw is IEnumerable items and not string and not byte[])
        {
            raw = JoinItems(items, lineBreaks.CollectionSeparator);
        }

        // Step 2: friendly text from ValueMappings.
        if (column.ValueMappings.Count > 0 && TryMapValue(raw, column.ValueMappings, out var mapped))
        {
            raw = mapped;
            isConfiguredText = true;
        }

        // Step 3: empty values.
        if (raw is null || raw is string { Length: 0 })
        {
            if (string.IsNullOrEmpty(column.DefaultValue))
            {
                return CellContent.Blank;
            }

            raw = column.DefaultValue;
            isConfiguredText = true;
        }

        // Step 4: convert to the requested Excel data type.
        var dataType = column.DataType == ExcelDataType.Auto ? DetectDataType(raw) : column.DataType;
        var format = string.IsNullOrWhiteSpace(column.Format) ? null : column.Format;
        var formats = _defaults.Formats;

        switch (dataType)
        {
            case ExcelDataType.Integer when TryGetNumber(raw, out var number):
                return new CellContent(Math.Round(number, MidpointRounding.AwayFromZero), format ?? DefaultFormat(column, formats.Integer));

            case ExcelDataType.Number when TryGetNumber(raw, out var number):
                return new CellContent(number, format ?? DefaultFormat(column, formats.Number));

            case ExcelDataType.Currency when TryGetNumber(raw, out var number):
                return new CellContent(number, format ?? formats.Currency);

            case ExcelDataType.Percentage when TryGetNumber(raw, out var number):
                return new CellContent(number, format ?? formats.Percentage);

            case ExcelDataType.Date when TryGetDateTime(raw, out var date):
                return new CellContent(date, format ?? formats.Date);

            case ExcelDataType.DateTime when TryGetDateTime(raw, out var dateTime):
                return new CellContent(dateTime, format ?? formats.DateTime);

            case ExcelDataType.Time when TryGetTime(raw, out var time):
                return new CellContent(time, format ?? formats.Time);

            case ExcelDataType.Boolean when TryGetBoolean(raw, out var flag):
                return new CellContent(flag);

            case ExcelDataType.Hyperlink:
                var url = ToText(raw, null).Trim();

                // A plain e-mail address becomes a mailto: link but keeps showing the address.
                var target = url.Contains('@') && !url.Contains(':') ? "mailto:" + url : url;

                return Uri.TryCreate(target, UriKind.Absolute, out var uri)
                    ? new CellContent(url, Hyperlink: uri)
                    : TextCell(url, lineBreaks);

            case ExcelDataType.Text:
            case ExcelDataType.Auto:
                // For text, "Format" is a .NET format string, for example "N2" or "dd/MM/yyyy".
                return TextCell(ToText(raw, format), lineBreaks);

            default:
                // The value did not match the requested type. Keep it visible as text.
                warning = isConfiguredText ? null : $"Value '{ToText(raw, null)}' could not be written as {dataType}; it was written as text.";
                return TextCell(ToText(raw, null), lineBreaks);
        }
    }

    /// <summary>
    /// Picks a data type from the value itself, used when the column's DataType is "Auto".
    /// </summary>
    internal static ExcelDataType DetectDataType(object value) => value switch
    {
        bool => ExcelDataType.Boolean,
        byte or sbyte or short or ushort or int or uint or long or ulong => ExcelDataType.Integer,
        float or double or decimal => ExcelDataType.Number,
        DateOnly => ExcelDataType.Date,
        DateTime { TimeOfDay.Ticks: 0 } => ExcelDataType.Date,
        DateTime or DateTimeOffset => ExcelDataType.DateTime,
        TimeOnly or TimeSpan => ExcelDataType.Time,
        _ => ExcelDataType.Text,
    };

    // With "Auto", plain numbers keep Excel's "General" format unless a Format is configured.
    private static string? DefaultFormat(ColumnDefinition column, string configuredDefault) =>
        column.DataType == ExcelDataType.Auto ? null : configuredDefault;

    private static CellContent TextCell(string text, LineBreakOptions lineBreaks)
    {
        text = LineBreakHelper.Normalize(text, lineBreaks);

        if (text.Length > MaxCellTextLength)
        {
            text = text[..MaxCellTextLength];
        }

        return new CellContent(text);
    }

    private string JoinItems(IEnumerable items, string separator)
    {
        var parts = new List<string>();

        foreach (var item in items)
        {
            if (item is not null)
            {
                parts.Add(ToText(item, null));
            }
        }

        return string.Join(separator, parts);
    }

    private static bool TryMapValue(object? raw, Dictionary<string, string> mappings, out string mapped)
    {
        // null is looked up as an empty key "" so you can map missing values too.
        var key = raw switch
        {
            null => string.Empty,
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
            _ => raw.ToString() ?? string.Empty,
        };

        if (mappings.TryGetValue(key, out mapped!))
        {
            return true;
        }

        // The dictionary may not ignore case if it was created by the configuration binder,
        // so fall back to a case-insensitive search.
        foreach (var pair in mappings)
        {
            if (string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase))
            {
                mapped = pair.Value;
                return true;
            }
        }

        mapped = string.Empty;
        return false;
    }

    private string ToText(object value, string? dotNetFormat) => value switch
    {
        string text => text,
        Enum enumValue => enumValue.ToString(),
        IFormattable formattable => formattable.ToString(dotNetFormat, _culture),
        _ => value.ToString() ?? string.Empty,
    };

    private bool TryGetNumber(object value, out double number)
    {
        switch (value)
        {
            case byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal:
                number = System.Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return !double.IsNaN(number) && !double.IsInfinity(number);
            case Enum enumValue:
                number = System.Convert.ToDouble(enumValue, CultureInfo.InvariantCulture);
                return true;
            case string text:
                return double.TryParse(text.Trim(), NumberStyles.Any, _culture, out number);
            default:
                number = 0;
                return false;
        }
    }

    private bool TryGetDateTime(object value, out DateTime dateTime)
    {
        switch (value)
        {
            case DateTime dt:
                dateTime = dt;
                return true;
            case DateTimeOffset dto:
                // Keep the clock time the user sees; Excel has no time zones.
                dateTime = dto.DateTime;
                return true;
            case DateOnly dateOnly:
                dateTime = dateOnly.ToDateTime(TimeOnly.MinValue);
                return true;
            case string text:
                return DateTime.TryParse(text.Trim(), _culture, DateTimeStyles.None, out dateTime);
            default:
                dateTime = default;
                return false;
        }
    }

    private bool TryGetTime(object value, out TimeSpan time)
    {
        switch (value)
        {
            case TimeSpan span:
                time = span;
                return true;
            case TimeOnly timeOnly:
                time = timeOnly.ToTimeSpan();
                return true;
            case DateTime dt:
                time = dt.TimeOfDay;
                return true;
            case DateTimeOffset dto:
                time = dto.TimeOfDay;
                return true;
            case string text:
                return TimeSpan.TryParse(text.Trim(), _culture, out time);
            default:
                time = default;
                return false;
        }
    }

    private static bool TryGetBoolean(object value, out bool flag)
    {
        switch (value)
        {
            case bool b:
                flag = b;
                return true;
            case byte or sbyte or short or ushort or int or uint or long or ulong:
                flag = System.Convert.ToInt64(value, CultureInfo.InvariantCulture) != 0;
                return true;
            case string text:
                switch (text.Trim().ToUpperInvariant())
                {
                    case "TRUE" or "YES" or "Y" or "1":
                        flag = true;
                        return true;
                    case "FALSE" or "NO" or "N" or "0":
                        flag = false;
                        return true;
                }

                break;
        }

        flag = false;
        return false;
    }
}
