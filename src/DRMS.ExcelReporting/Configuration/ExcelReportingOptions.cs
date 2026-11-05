namespace DRMS.ExcelReporting.Configuration;

// These classes mirror the "ExcelReporting" section in appsettings.json.
// Every property has a sensible default, so you only need to write what you want to change.

/// <summary>
/// Root of the "ExcelReporting" configuration section.
/// </summary>
public sealed class ExcelReportingOptions
{
    /// <summary>The default name of the configuration section.</summary>
    public const string SectionName = "ExcelReporting";

    /// <summary>Settings shared by every report (formats, header style, line breaks).</summary>
    public ReportDefaults Defaults { get; set; } = new();

    /// <summary>
    /// All reports, keyed by report name. The key is the name you pass to
    /// <see cref="IExcelReportGenerator.Generate(string, ExcelReportData)"/>.
    /// Keys are matched without caring about upper or lower case.
    /// </summary>
    public Dictionary<string, ReportDefinition> Reports { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

/// <summary>
/// Settings used by all reports unless a report, sheet or column overrides them.
/// </summary>
public sealed class ReportDefaults
{
    /// <summary>
    /// Culture used to read numbers and dates that arrive as text (for example "1,234.50").
    /// Empty means the invariant culture.
    /// </summary>
    public string Culture { get; set; } = string.Empty;

    /// <summary>Default Excel number formats for each data type.</summary>
    public DataTypeFormats Formats { get; set; } = new();

    /// <summary>Look of the header row.</summary>
    public HeaderStyleOptions HeaderStyle { get; set; } = new();

    /// <summary>Default line break rules. A column can override them.</summary>
    public LineBreakOptions LineBreaks { get; set; } = new();

    /// <summary>
    /// When true, column widths are fitted to their content (up to <see cref="MaxColumnWidth"/>).
    /// A column with a fixed "Width" is never auto-fitted.
    /// </summary>
    public bool AutoFitColumns { get; set; } = true;

    /// <summary>Upper limit for auto-fitted columns, in Excel character units.</summary>
    public double MaxColumnWidth { get; set; } = 60;

    /// <summary>Keep the header row visible when scrolling.</summary>
    public bool FreezeHeaderRow { get; set; } = true;

    /// <summary>Add filter drop-downs to the header row.</summary>
    public bool AutoFilter { get; set; } = true;

    /// <summary>
    /// Maximum data rows on one sheet. When there are more rows, the sheet is continued
    /// on extra sheets named "Name (2)", "Name (3)" and so on.
    /// Excel allows 1,048,576 rows including the header.
    /// </summary>
    public int MaxRowsPerSheet { get; set; } = 1_048_575;
}

/// <summary>
/// Default Excel number formats, one per data type.
/// These are Excel format codes, not .NET format strings.
/// </summary>
public sealed class DataTypeFormats
{
    /// <summary>Format for <see cref="ExcelDataType.Integer"/>.</summary>
    public string Integer { get; set; } = "0";

    /// <summary>Format for <see cref="ExcelDataType.Number"/>.</summary>
    public string Number { get; set; } = "#,##0.00";

    /// <summary>Format for <see cref="ExcelDataType.Currency"/>.</summary>
    public string Currency { get; set; } = "#,##0.00";

    /// <summary>Format for <see cref="ExcelDataType.Percentage"/>.</summary>
    public string Percentage { get; set; } = "0.00%";

    /// <summary>Format for <see cref="ExcelDataType.Date"/>.</summary>
    public string Date { get; set; } = "yyyy-mm-dd";

    /// <summary>Format for <see cref="ExcelDataType.DateTime"/>.</summary>
    public string DateTime { get; set; } = "yyyy-mm-dd hh:mm:ss";

    /// <summary>Format for <see cref="ExcelDataType.Time"/>.</summary>
    public string Time { get; set; } = "hh:mm:ss";
}

/// <summary>
/// Look of the header row.
/// </summary>
public sealed class HeaderStyleOptions
{
    /// <summary>Show header text in bold.</summary>
    public bool Bold { get; set; } = true;

    /// <summary>Header fill colour as a hex value, for example "#D9E1F2". Empty means no fill.</summary>
    public string BackgroundColor { get; set; } = "#D9E1F2";

    /// <summary>Header font colour as a hex value, for example "#000000". Empty means the default colour.</summary>
    public string FontColor { get; set; } = string.Empty;

    /// <summary>Wrap header text. Needed when a header contains a line break.</summary>
    public bool WrapText { get; set; } = true;
}

/// <summary>
/// Rules for turning parts of a value into new lines inside one Excel cell.
/// </summary>
public sealed class LineBreakOptions
{
    /// <summary>
    /// Text markers in the data that should become a new line, for example "|" or "&lt;br&gt;".
    /// Real line breaks (\r\n, \r, \n) are always kept; you do not need to list them.
    /// </summary>
    public List<string> Separators { get; set; } = [];

    /// <summary>
    /// Text placed between items when a property is a list (for example List&lt;string&gt;).
    /// The default is a new line, so each item gets its own line in the cell.
    /// </summary>
    public string CollectionSeparator { get; set; } = "\n";

    /// <summary>
    /// When true, any cell that ends up with a line break gets "Wrap Text" turned on,
    /// otherwise Excel would show everything on one line.
    /// </summary>
    public bool WrapTextWhenLineBreakFound { get; set; } = true;
}

/// <summary>
/// One Excel file.
/// </summary>
public sealed class ReportDefinition
{
    /// <summary>Suggested file name, for example "EmployeeReport.xlsx". Used by the Save helpers.</summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>The sheets in the file, in the order they appear in Excel.</summary>
    public List<SheetDefinition> Sheets { get; set; } = [];
}

/// <summary>
/// One worksheet inside the Excel file.
/// </summary>
public sealed class SheetDefinition
{
    /// <summary>
    /// The tab name shown in Excel. Maximum 31 characters, and it must not contain : \ / ? * [ ].
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Which list of data fills this sheet. It matches the key used in
    /// <see cref="ExcelReportData.Add{T}(string, IEnumerable{T})"/>.
    /// When empty, the sheet <see cref="Name"/> is used as the key.
    /// </summary>
    public string DataSource { get; set; } = string.Empty;

    /// <summary>Optional title written above the header row (for example "Employee List - 2026").</summary>
    public string Title { get; set; } = string.Empty;

    /// <summary>Overrides <see cref="ReportDefaults.FreezeHeaderRow"/> for this sheet.</summary>
    public bool? FreezeHeaderRow { get; set; }

    /// <summary>Overrides <see cref="ReportDefaults.AutoFilter"/> for this sheet.</summary>
    public bool? AutoFilter { get; set; }

    /// <summary>The columns, in the order they appear in Excel.</summary>
    public List<ColumnDefinition> Columns { get; set; } = [];

    /// <summary>Gets the data source key, falling back to the sheet name.</summary>
    public string GetDataSourceKey() => string.IsNullOrWhiteSpace(DataSource) ? Name : DataSource;
}

/// <summary>
/// One column: which DTO property it reads, its header text and how the value is written.
/// </summary>
public sealed class ColumnDefinition
{
    /// <summary>
    /// Header text shown in the first row. Use "\n" in JSON to split the header over two lines.
    /// When empty, the property name is used.
    /// </summary>
    public string Header { get; set; } = string.Empty;

    /// <summary>
    /// The DTO property to read. Use dots for nested objects, for example "Address.City".
    /// Matching ignores upper or lower case.
    /// </summary>
    public string Property { get; set; } = string.Empty;

    /// <summary>How the value is written to Excel. Default: <see cref="ExcelDataType.Auto"/>.</summary>
    public ExcelDataType DataType { get; set; } = ExcelDataType.Auto;

    /// <summary>
    /// Excel number format for this column, for example "dd-mmm-yyyy" or "$#,##0.00".
    /// When empty, the default format for the data type is used.
    /// </summary>
    public string Format { get; set; } = string.Empty;

    /// <summary>
    /// Replaces raw values with friendly text before writing, for example
    /// { "A": "Active", "I": "Inactive" } or { "True": "Yes", "False": "No" }.
    /// Keys are matched without caring about upper or lower case.
    /// </summary>
    public Dictionary<string, string> ValueMappings { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Text written when the value is null or empty. Leave empty for a blank cell.</summary>
    public string DefaultValue { get; set; } = string.Empty;

    /// <summary>Fixed column width in Excel character units. When not set, the column is auto-fitted.</summary>
    public double? Width { get; set; }

    /// <summary>Always wrap text in this column, even without line breaks.</summary>
    public bool WrapText { get; set; }

    /// <summary>Horizontal alignment: "Left", "Center" or "Right". Empty means Excel's default.</summary>
    public string Alignment { get; set; } = string.Empty;

    /// <summary>
    /// Line break rules for this column only. When not set, <see cref="ReportDefaults.LineBreaks"/> is used.
    /// </summary>
    public LineBreakOptions? LineBreaks { get; set; }

    /// <summary>Gets the header text, falling back to the property name.</summary>
    public string GetHeaderText() => string.IsNullOrWhiteSpace(Header) ? Property : Header;
}
