namespace DRMS.ExcelReporting.Configuration;

/// <summary>
/// The kind of value written to an Excel cell.
/// Set it per column in appsettings.json, for example: "DataType": "Currency".
/// </summary>
public enum ExcelDataType
{
    /// <summary>
    /// Default. The type is picked from the value itself
    /// (numbers stay numbers, dates stay dates, everything else becomes text).
    /// </summary>
    Auto = 0,

    /// <summary>Plain text. Values like "00123" keep their leading zeros.</summary>
    Text,

    /// <summary>A whole number. Default format: <c>0</c>.</summary>
    Integer,

    /// <summary>A number with decimals. Default format: <c>#,##0.00</c>.</summary>
    Number,

    /// <summary>A money value. Default format: <c>#,##0.00</c> (change it with "Format").</summary>
    Currency,

    /// <summary>A percentage. The value 0.25 is shown as 25%. Default format: <c>0.00%</c>.</summary>
    Percentage,

    /// <summary>A date without time. Default format: <c>yyyy-mm-dd</c>.</summary>
    Date,

    /// <summary>A date with time. Default format: <c>yyyy-mm-dd hh:mm:ss</c>.</summary>
    DateTime,

    /// <summary>A time of day or a duration. Default format: <c>hh:mm:ss</c>.</summary>
    Time,

    /// <summary>A TRUE / FALSE cell. Use "Text" with ValueMappings if you want "Yes" / "No".</summary>
    Boolean,

    /// <summary>A clickable link. The value must be a URL (or a mailto: address).</summary>
    Hyperlink,
}
