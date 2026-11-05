using System.Collections;
using ClosedXML.Excel;
using DRMS.ExcelReporting.Configuration;
using DRMS.ExcelReporting.Internal;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace DRMS.ExcelReporting;

/// <summary>
/// Creates Excel (.xlsx) files from lists of DTOs. The layout of each report — sheets, columns,
/// header text, DTO property mapping, data types, formats and line breaks — comes from the
/// "ExcelReporting" section of appsettings.json, so you can change a report without changing code.
/// </summary>
/// <remarks>
/// How a report is built:
/// 1. Find the report definition by name in configuration.
/// 2. For each sheet, find its list of DTOs in <see cref="ExcelReportData"/>.
/// 3. Write an optional title row, then the header row.
/// 4. For each DTO and each column, read the configured property and convert it to the column's data type.
/// 5. Apply column widths, alignment, frozen header and filter.
/// 6. When a sheet has more rows than allowed, continue on a new sheet named "Name (2)", "Name (3)", ...
///
/// The result can be returned as bytes, written to any stream, returned as a <see cref="MemoryStream"/>,
/// or saved as a file on disk (<see cref="GenerateToFile(string, ExcelReportData, string?, string?)"/>),
/// for example to upload it to object storage and delete it afterwards.
///
/// This class is thread-safe and can be registered as a singleton.
/// </remarks>
public sealed class ExcelReportGenerator : IExcelReportGenerator
{
    /// <summary>The MIME type of an .xlsx file. Use it when returning the file from a web API.</summary>
    public const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /// <summary>The maximum number of rows in one Excel sheet.</summary>
    private const int ExcelMaxRows = 1_048_576;

    /// <summary>Only this many rows are measured when auto-fitting column widths, to keep large reports fast.</summary>
    private const int AutoFitSampleRows = 1_000;

    /// <summary>Extra width added to auto-fitted columns, in Excel character units.</summary>
    private const double AutoFitPadding = 3;

    /// <summary>Extension of the temporary file used while the workbook is being written.</summary>
    private const string TempFileExtension = ".xlsx.tmp";

    /// <summary>"{ReportName}" value used for reports built in code.</summary>
    private const string CodeReportName = "Report";

    private readonly Func<ExcelReportingOptions> _getOptions;
    private readonly ILogger _logger;
    private readonly TimeProvider _timeProvider;

    /// <summary>
    /// Constructor used by dependency injection. Reads the latest configuration on every call,
    /// so changes to appsettings.json are picked up without restarting the application.
    /// </summary>
    /// <param name="options">The "ExcelReporting" settings.</param>
    /// <param name="logger">Optional logger.</param>
    /// <param name="timeProvider">Optional clock used for {Date}/{Time}/{Timestamp} in file names. Default: system clock.</param>
    public ExcelReportGenerator(
        IOptionsMonitor<ExcelReportingOptions> options,
        ILogger<ExcelReportGenerator>? logger = null,
        TimeProvider? timeProvider = null)
    {
        ArgumentNullException.ThrowIfNull(options);

        _getOptions = () => options.CurrentValue;
        _logger = logger ?? NullLogger<ExcelReportGenerator>.Instance;
        _timeProvider = timeProvider ?? TimeProvider.System;
    }

    /// <summary>
    /// Constructor for use without dependency injection, for example in a console tool or a test.
    /// </summary>
    /// <param name="options">The "ExcelReporting" settings.</param>
    /// <param name="logger">Optional logger.</param>
    /// <param name="timeProvider">Optional clock used for {Date}/{Time}/{Timestamp} in file names. Default: system clock.</param>
    public ExcelReportGenerator(
        ExcelReportingOptions options,
        ILogger<ExcelReportGenerator>? logger = null,
        TimeProvider? timeProvider = null)
    {
        ArgumentNullException.ThrowIfNull(options);

        _getOptions = () => options;
        _logger = logger ?? NullLogger<ExcelReportGenerator>.Instance;
        _timeProvider = timeProvider ?? TimeProvider.System;
    }

    /// <inheritdoc />
    public byte[] Generate<T>(string reportName, IEnumerable<T> items) =>
        Generate(reportName, ExcelReportData.ForAllSheets(items));

    /// <inheritdoc />
    public byte[] Generate(string reportName, ExcelReportData data)
    {
        using var stream = new MemoryStream();
        Generate(reportName, data, stream);
        return stream.ToArray();
    }

    /// <inheritdoc />
    public void Generate(string reportName, ExcelReportData data, Stream output)
    {
        var options = _getOptions();
        var report = FindReport(options, reportName);
        WriteReport(reportName, report, options.Defaults, data, output);
    }

    /// <inheritdoc />
    public byte[] Generate(ReportDefinition report, ExcelReportData data)
    {
        using var stream = new MemoryStream();
        Generate(report, data, stream);
        return stream.ToArray();
    }

    /// <inheritdoc />
    public void Generate(ReportDefinition report, ExcelReportData data, Stream output)
    {
        EnsureValid(report);
        WriteReport(CodeReportName, report, _getOptions().Defaults, data, output);
    }

    /// <inheritdoc />
    public MemoryStream GenerateToStream(string reportName, ExcelReportData data)
    {
        var stream = new MemoryStream();
        try
        {
            Generate(reportName, data, stream);
            stream.Position = 0;
            return stream;
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    /// <inheritdoc />
    public ExcelReportFile GenerateToFile(string reportName, ExcelReportData data, string? outputDirectory = null, string? fileName = null)
    {
        var options = _getOptions();
        var report = FindReport(options, reportName);
        return WriteReportFile(reportName, report, options.Defaults, data, outputDirectory, fileName);
    }

    /// <inheritdoc />
    public ExcelReportFile GenerateToFile(ReportDefinition report, ExcelReportData data, string? outputDirectory = null, string? fileName = null)
    {
        EnsureValid(report);
        return WriteReportFile(CodeReportName, report, _getOptions().Defaults, data, outputDirectory, fileName);
    }

    /// <inheritdoc />
    public string GetFileName(string reportName)
    {
        var options = _getOptions();
        var report = FindReport(options, reportName);

        return OutputPathBuilder.BuildFileName(GetFileNamePattern(report, options.Defaults.Output), reportName, _timeProvider.GetLocalNow());
    }

    /// <inheritdoc />
    public string GetOutputDirectory(string reportName)
    {
        var options = _getOptions();
        var report = FindReport(options, reportName);

        return OutputPathBuilder.BuildDirectory(GetDirectoryPattern(report, options.Defaults.Output), reportName);
    }

    /// <inheritdoc />
    public int DeleteOldReportFiles(TimeSpan? olderThan = null)
    {
        var options = _getOptions();
        var output = options.Defaults.Output;
        var cutoffUtc = _timeProvider.GetUtcNow().UtcDateTime - (olderThan ?? output.CleanupFilesOlderThan);

        // Every folder a report can be saved in: the default folder and each report's own folder.
        var directories = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (name, report) in options.Reports)
        {
            directories.Add(OutputPathBuilder.BuildDirectory(GetDirectoryPattern(report, output), name));
        }

        if (!output.Directory.Contains("{ReportName}", StringComparison.OrdinalIgnoreCase))
        {
            directories.Add(OutputPathBuilder.BuildDirectory(output.Directory, CodeReportName));
        }

        var deleted = 0;
        foreach (var directory in directories)
        {
            deleted += DeleteOldFilesIn(directory, cutoffUtc);
        }

        if (deleted > 0)
        {
            _logger.LogInformation("Deleted {Count} old report file(s).", deleted);
        }

        return deleted;
    }

    // ----------------------------------------------------------------------------------------
    // Saving to disk
    // ----------------------------------------------------------------------------------------

    /// <summary>
    /// Saves the report to disk. The workbook is first written to a hidden temporary file and then
    /// renamed, so nobody ever sees (or uploads) a half-written file.
    /// </summary>
    private ExcelReportFile WriteReportFile(
        string reportName,
        ReportDefinition report,
        ReportDefaults defaults,
        ExcelReportData data,
        string? directoryOverride,
        string? fileNameOverride)
    {
        var output = defaults.Output;
        var now = _timeProvider.GetLocalNow();

        var directory = OutputPathBuilder.BuildDirectory(
            string.IsNullOrWhiteSpace(directoryOverride) ? GetDirectoryPattern(report, output) : directoryOverride,
            reportName);

        var fileName = OutputPathBuilder.BuildFileName(
            string.IsNullOrWhiteSpace(fileNameOverride) ? GetFileNamePattern(report, output) : fileNameOverride,
            reportName,
            now);

        var createdDirectories = CreateDirectory(directory);
        var tempPath = Path.Combine(directory, $".{Guid.NewGuid():N}{TempFileExtension}");

        try
        {
            using (var stream = CreateTempFile(tempPath))
            {
                WriteReport(reportName, report, defaults, data, stream);
            }

            var finalPath = MoveToFinalName(tempPath, directory, fileName, output.OverwriteExisting);

            _logger.LogInformation("Excel report '{ReportName}' saved to '{FilePath}'.", reportName, finalPath);

            return new ExcelReportFile(
                reportName,
                finalPath,
                now,
                createdDirectories,
                output.DeleteFileOnDispose,
                output.DeleteEmptyDirectories,
                _logger);
        }
        catch
        {
            // Leave nothing behind when the report could not be created.
            TryDeleteFile(tempPath);
            if (output.DeleteEmptyDirectories)
            {
                ExcelReportFile.DeleteEmptyCreatedDirectories(createdDirectories, _logger);
            }

            throw;
        }
    }

    /// <summary>
    /// Creates the folder and returns the folders that did not exist before, deepest first,
    /// so they can be removed again when the file is deleted.
    /// </summary>
    private static List<string> CreateDirectory(string directory)
    {
        var missing = new List<string>();
        for (var current = directory; current is not null && !Directory.Exists(current); current = Path.GetDirectoryName(current))
        {
            missing.Add(current);
        }

        Directory.CreateDirectory(directory);
        return missing;
    }

    private static FileStream CreateTempFile(string tempPath)
    {
        try
        {
            return new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        }
        catch (DirectoryNotFoundException)
        {
            // Another report may have just removed this (empty) folder during its cleanup. Create it again.
            Directory.CreateDirectory(Path.GetDirectoryName(tempPath)!);
            return new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        }
    }

    /// <summary>
    /// Renames the temporary file to its final name. When the name is taken and overwriting is off,
    /// tries "Name_1.xlsx", "Name_2.xlsx", ... until a free name is found.
    /// </summary>
    private static string MoveToFinalName(string tempPath, string directory, string fileName, bool overwrite)
    {
        var target = Path.Combine(directory, fileName);

        if (overwrite)
        {
            File.Move(tempPath, target, overwrite: true);
            return target;
        }

        var stem = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);

        for (var attempt = 1; attempt <= 10_000; attempt++)
        {
            try
            {
                File.Move(tempPath, target, overwrite: false);
                return target;
            }
            catch (IOException) when (File.Exists(target))
            {
                target = Path.Combine(directory, $"{stem}_{attempt}{extension}");
            }
        }

        throw new IOException($"Could not find a free file name for '{fileName}' in '{directory}'.");
    }

    private int DeleteOldFilesIn(string directory, DateTime cutoffUtc)
    {
        // Safety net: never clean the root of a drive.
        if (!Directory.Exists(directory) || string.Equals(Path.GetPathRoot(directory), directory, StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        var deleted = 0;
        var candidates = Directory.EnumerateFiles(directory, "*.xlsx")
            .Concat(Directory.EnumerateFiles(directory, "*" + TempFileExtension));

        foreach (var file in candidates)
        {
            try
            {
                if (File.GetLastWriteTimeUtc(file) < cutoffUtc)
                {
                    File.Delete(file);
                    deleted++;
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // The file may be open (for example still being uploaded). It is tried again next time.
                _logger.LogDebug(ex, "Old report file '{File}' was not deleted.", file);
            }
        }

        return deleted;
    }

    private void TryDeleteFile(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _logger.LogWarning(ex, "Could not delete temporary file '{FilePath}'.", path);
        }
    }

    private static string GetFileNamePattern(ReportDefinition report, OutputOptions output) =>
        string.IsNullOrWhiteSpace(report.FileName) ? output.FileName : report.FileName;

    private static string GetDirectoryPattern(ReportDefinition report, OutputOptions output) =>
        string.IsNullOrWhiteSpace(report.OutputDirectory) ? output.Directory : report.OutputDirectory;

    private static void EnsureValid(ReportDefinition report)
    {
        ArgumentNullException.ThrowIfNull(report);

        var errors = ExcelReportingOptionsValidator.ValidateReport(CodeReportName, report);
        if (errors.Count > 0)
        {
            throw new InvalidOperationException("The report definition is not valid:" + Environment.NewLine + string.Join(Environment.NewLine, errors));
        }
    }

    // ----------------------------------------------------------------------------------------
    // Report and sheet writing
    // ----------------------------------------------------------------------------------------

    private void WriteReport(string reportName, ReportDefinition report, ReportDefaults defaults, ExcelReportData data, Stream output)
    {
        ArgumentNullException.ThrowIfNull(data);
        ArgumentNullException.ThrowIfNull(output);

        var converter = new CellValueConverter(defaults);
        using var workbook = new XLWorkbook();

        foreach (var sheet in report.Sheets)
        {
            var dataSource = sheet.GetDataSourceKey();
            if (!data.TryGetSource(dataSource, out var items))
            {
                throw new InvalidOperationException(
                    $"Report '{reportName}', sheet '{sheet.Name}' needs data source '{dataSource}', but it was not provided. " +
                    $"Add it with ExcelReportData.Add(\"{dataSource}\", items). Provided sources: [{string.Join(", ", data.SourceNames)}].");
            }

            WriteSheetWithOverflow(workbook, sheet, defaults, converter, items);
        }

        workbook.SaveAs(output);

        _logger.LogDebug("Excel report '{ReportName}' created with {SheetCount} sheet(s).", reportName, workbook.Worksheets.Count);
    }

    /// <summary>
    /// Writes one configured sheet. If there are more rows than one sheet allows,
    /// the rows continue on extra sheets: "Name (2)", "Name (3)", and so on.
    /// </summary>
    private void WriteSheetWithOverflow(XLWorkbook workbook, SheetDefinition sheet, ReportDefaults defaults, CellValueConverter converter, IEnumerable items)
    {
        var headerRow = string.IsNullOrWhiteSpace(sheet.Title) ? 1 : 2;
        var maxRowsPerSheet = Math.Min(defaults.MaxRowsPerSheet, ExcelMaxRows - headerRow);

        // Walk the list only once, so it also works with data streamed from a database.
        var enumerator = items.GetEnumerator();
        try
        {
            var part = 1;
            var hasMore = enumerator.MoveNext();

            do
            {
                var sheetName = part == 1 ? sheet.Name : BuildContinuationSheetName(workbook, sheet.Name, part);
                var worksheet = workbook.Worksheets.Add(sheetName);

                hasMore = WriteSheet(worksheet, sheet, defaults, converter, enumerator, hasMore, headerRow, maxRowsPerSheet);
                part++;
            }
            while (hasMore);
        }
        finally
        {
            (enumerator as IDisposable)?.Dispose();
        }
    }

    /// <summary>
    /// Writes the title, header and up to <paramref name="maxRows"/> data rows to one worksheet.
    /// Returns true when there are still rows left for another sheet.
    /// </summary>
    private bool WriteSheet(
        IXLWorksheet worksheet,
        SheetDefinition sheet,
        ReportDefaults defaults,
        CellValueConverter converter,
        IEnumerator items,
        bool hasCurrent,
        int headerRow,
        int maxRows)
    {
        var columns = sheet.Columns;

        WriteTitle(worksheet, sheet, columns.Count);
        WriteHeader(worksheet, columns, defaults, headerRow);

        var problems = new ConversionProblemCollector();
        var row = headerRow;
        var rowsWritten = 0;

        while (hasCurrent && rowsWritten < maxRows)
        {
            var item = items.Current;

            // Skip null entries in the list instead of failing.
            if (item is not null)
            {
                row++;
                rowsWritten++;
                WriteDataRow(worksheet, sheet, defaults, converter, item, row, problems);
            }

            hasCurrent = items.MoveNext();
        }

        var lastRow = Math.Max(row, headerRow);
        FormatColumns(worksheet, columns, defaults, headerRow, lastRow);

        if (sheet.FreezeHeaderRow ?? defaults.FreezeHeaderRow)
        {
            worksheet.SheetView.FreezeRows(headerRow);
        }

        if (sheet.AutoFilter ?? defaults.AutoFilter)
        {
            worksheet.Range(headerRow, 1, lastRow, columns.Count).SetAutoFilter();
        }

        problems.LogTo(_logger, worksheet.Name);
        return hasCurrent;
    }

    private static void WriteTitle(IXLWorksheet worksheet, SheetDefinition sheet, int columnCount)
    {
        if (string.IsNullOrWhiteSpace(sheet.Title))
        {
            return;
        }

        var titleRange = worksheet.Range(1, 1, 1, Math.Max(columnCount, 1));
        titleRange.Merge();

        var titleCell = worksheet.Cell(1, 1);
        titleCell.Value = sheet.Title;
        titleCell.Style.Font.Bold = true;
        titleCell.Style.Font.FontSize = 14;
    }

    private void WriteHeader(IXLWorksheet worksheet, List<ColumnDefinition> columns, ReportDefaults defaults, int headerRow)
    {
        var style = defaults.HeaderStyle;

        for (var c = 0; c < columns.Count; c++)
        {
            var cell = worksheet.Cell(headerRow, c + 1);

            // Header text can contain line breaks too, for example "Joining\nDate".
            cell.Value = LineBreakHelper.Normalize(columns[c].GetHeaderText(), defaults.LineBreaks);
        }

        var headerRange = worksheet.Range(headerRow, 1, headerRow, columns.Count);
        headerRange.Style.Font.Bold = style.Bold;
        headerRange.Style.Alignment.WrapText = style.WrapText;
        headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

        if (TryParseColor(style.BackgroundColor, out var background))
        {
            headerRange.Style.Fill.BackgroundColor = background;
        }

        if (TryParseColor(style.FontColor, out var fontColor))
        {
            headerRange.Style.Font.FontColor = fontColor;
        }
    }

    /// <summary>
    /// Maps one DTO to one Excel row: for each configured column, read the DTO property
    /// and write it as the configured data type.
    /// </summary>
    private static void WriteDataRow(
        IXLWorksheet worksheet,
        SheetDefinition sheet,
        ReportDefaults defaults,
        CellValueConverter converter,
        object item,
        int row,
        ConversionProblemCollector problems)
    {
        for (var c = 0; c < sheet.Columns.Count; c++)
        {
            var column = sheet.Columns[c];

            object? rawValue;
            try
            {
                rawValue = PropertyValueResolver.GetValue(item, column.Property);
            }
            catch (InvalidOperationException ex)
            {
                // A wrong property name is a configuration mistake. Stop with a clear message.
                throw new InvalidOperationException(
                    $"Sheet '{sheet.Name}', column '{column.GetHeaderText()}' (Property \"{column.Property}\"): {ex.Message}", ex);
            }

            var content = converter.Convert(rawValue, column, out var warning);
            if (warning is not null)
            {
                problems.Add(column.GetHeaderText(), row, warning);
            }

            WriteCell(worksheet.Cell(row, c + 1), content, column, defaults);
        }
    }

    private static void WriteCell(IXLCell cell, CellContent content, ColumnDefinition column, ReportDefaults defaults)
    {
        cell.Value = content.Value;

        if (content.NumberFormat is not null)
        {
            cell.Style.NumberFormat.Format = content.NumberFormat;
        }

        if (content.Hyperlink is not null)
        {
            cell.SetHyperlink(new XLHyperlink(content.Hyperlink));
            cell.Style.Font.FontColor = XLColor.Blue;
            cell.Style.Font.Underline = XLFontUnderlineValues.Single;
        }

        // Excel shows "\n" as a new line only when "Wrap Text" is on.
        var lineBreaks = column.LineBreaks ?? defaults.LineBreaks;
        if (column.WrapText || (lineBreaks.WrapTextWhenLineBreakFound && content.HasLineBreak))
        {
            cell.Style.Alignment.WrapText = true;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        }
    }

    private static void FormatColumns(IXLWorksheet worksheet, List<ColumnDefinition> columns, ReportDefaults defaults, int headerRow, int lastRow)
    {
        for (var c = 0; c < columns.Count; c++)
        {
            var column = columns[c];
            var excelColumn = worksheet.Column(c + 1);

            if (lastRow > headerRow && TryParseAlignment(column.Alignment, out var alignment))
            {
                worksheet.Range(headerRow + 1, c + 1, lastRow, c + 1).Style.Alignment.Horizontal = alignment;
            }

            if (column.Width is { } width)
            {
                excelColumn.Width = width;
            }
            else if (defaults.AutoFitColumns)
            {
                excelColumn.AdjustToContents(headerRow, Math.Min(lastRow, headerRow + AutoFitSampleRows));

                // Extra room for the filter drop-down arrow and a small margin.
                excelColumn.Width = Math.Min(excelColumn.Width + AutoFitPadding, defaults.MaxColumnWidth);
            }
        }
    }

    // ----------------------------------------------------------------------------------------
    // Small helpers
    // ----------------------------------------------------------------------------------------

    private static ReportDefinition FindReport(ExcelReportingOptions options, string reportName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(reportName);

        if (options.Reports.TryGetValue(reportName, out var report))
        {
            return report;
        }

        // The configuration binder may create a case-sensitive dictionary, so search again ignoring case.
        foreach (var pair in options.Reports)
        {
            if (string.Equals(pair.Key, reportName, StringComparison.OrdinalIgnoreCase))
            {
                return pair.Value;
            }
        }

        throw new KeyNotFoundException(
            $"No report named '{reportName}' was found in the \"{ExcelReportingOptions.SectionName}:Reports\" configuration. " +
            $"Available reports: [{string.Join(", ", options.Reports.Keys)}].");
    }

    /// <summary>
    /// Builds "Name (2)", "Name (3)", ... and shortens the base name so the result fits Excel's 31-character limit.
    /// </summary>
    private static string BuildContinuationSheetName(XLWorkbook workbook, string baseName, int part)
    {
        while (true)
        {
            var suffix = $" ({part})";
            var maxBaseLength = ExcelReportingOptionsValidator.MaxSheetNameLength - suffix.Length;
            var name = (baseName.Length > maxBaseLength ? baseName[..maxBaseLength] : baseName) + suffix;

            if (!workbook.Worksheets.Contains(name))
            {
                return name;
            }

            part++;
        }
    }

    private static bool TryParseColor(string value, out XLColor color)
    {
        color = XLColor.NoColor;
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        try
        {
            color = XLColor.FromHtml(value.Trim());
            return true;
        }
        catch (Exception ex) when (ex is FormatException or ArgumentException or OverflowException)
        {
            return false;
        }
    }

    private static bool TryParseAlignment(string value, out XLAlignmentHorizontalValues alignment)
    {
        alignment = value.Trim().ToUpperInvariant() switch
        {
            "LEFT" => XLAlignmentHorizontalValues.Left,
            "CENTER" => XLAlignmentHorizontalValues.Center,
            "RIGHT" => XLAlignmentHorizontalValues.Right,
            _ => XLAlignmentHorizontalValues.General,
        };

        return alignment != XLAlignmentHorizontalValues.General;
    }

    /// <summary>
    /// Collects values that could not be converted to their column's data type,
    /// then logs one summary line per column instead of one line per cell.
    /// </summary>
    private sealed class ConversionProblemCollector
    {
        private readonly Dictionary<string, (int Count, int FirstRow, string FirstMessage)> _byColumn = new();

        public void Add(string column, int row, string message)
        {
            _byColumn[column] = _byColumn.TryGetValue(column, out var existing)
                ? existing with { Count = existing.Count + 1 }
                : (1, row, message);
        }

        public void LogTo(ILogger logger, string sheetName)
        {
            foreach (var (column, problem) in _byColumn)
            {
                logger.LogWarning(
                    "Sheet '{Sheet}', column '{Column}': {Count} value(s) were written as text because they did not match the data type. First at row {Row}: {Message}",
                    sheetName, column, problem.Count, problem.FirstRow, problem.FirstMessage);
            }
        }
    }
}
