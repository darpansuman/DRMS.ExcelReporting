using DRMS.ExcelReporting.Configuration;

namespace DRMS.ExcelReporting;

/// <summary>
/// Creates Excel (.xlsx) files from lists of DTOs, using report layouts from configuration.
/// </summary>
public interface IExcelReportGenerator
{
    /// <summary>
    /// Creates the report named <paramref name="reportName"/> where every sheet shows <paramref name="items"/>.
    /// </summary>
    /// <returns>The .xlsx file as bytes.</returns>
    byte[] Generate<T>(string reportName, IEnumerable<T> items);

    /// <summary>
    /// Creates the report named <paramref name="reportName"/>. Each sheet reads its own list from <paramref name="data"/>.
    /// </summary>
    /// <returns>The .xlsx file as bytes.</returns>
    byte[] Generate(string reportName, ExcelReportData data);

    /// <summary>
    /// Creates the report named <paramref name="reportName"/> and writes it to <paramref name="output"/>
    /// (for example a FileStream or an HTTP response body).
    /// </summary>
    void Generate(string reportName, ExcelReportData data, Stream output);

    /// <summary>
    /// Creates a report from a definition built in code instead of appsettings.json.
    /// </summary>
    /// <returns>The .xlsx file as bytes.</returns>
    byte[] Generate(ReportDefinition report, ExcelReportData data);

    /// <summary>
    /// Creates a report from a definition built in code and writes it to <paramref name="output"/>.
    /// </summary>
    void Generate(ReportDefinition report, ExcelReportData data, Stream output);

    /// <summary>
    /// Creates the report and returns it in a <see cref="MemoryStream"/> that is already at position 0,
    /// ready to upload without touching the disk. The caller must dispose the stream.
    /// </summary>
    MemoryStream GenerateToStream(string reportName, ExcelReportData data);

    /// <summary>
    /// Creates the report and saves it as a file on disk.
    /// </summary>
    /// <param name="reportName">The report key in configuration.</param>
    /// <param name="data">The lists of DTOs for the sheets.</param>
    /// <param name="outputDirectory">
    /// Optional folder for this call only. Overrides the report's "OutputDirectory" and "Output:Directory".
    /// Placeholders such as {TempPath} are allowed.
    /// </param>
    /// <param name="fileName">
    /// Optional file name for this call only. Overrides the report's "FileName" and "Output:FileName".
    /// Placeholders such as {Timestamp} are allowed; ".xlsx" is added when missing.
    /// </param>
    /// <returns>
    /// The saved file: use <see cref="ExcelReportFile.FilePath"/> to get its path and
    /// <see cref="ExcelReportFile.OpenRead"/> to read it. Dispose it to delete the file.
    /// </returns>
    ExcelReportFile GenerateToFile(string reportName, ExcelReportData data, string? outputDirectory = null, string? fileName = null);

    /// <summary>
    /// Creates a report from a definition built in code and saves it as a file on disk.
    /// "{ReportName}" in paths becomes "Report".
    /// </summary>
    ExcelReportFile GenerateToFile(ReportDefinition report, ExcelReportData data, string? outputDirectory = null, string? fileName = null);

    /// <summary>
    /// Returns the file name the report would get right now, with placeholders filled in
    /// (for example "EmployeeReport_20261105_101500.xlsx"). Useful as a download name.
    /// </summary>
    string GetFileName(string reportName);

    /// <summary>
    /// Returns the full path of the folder where the report's files are saved.
    /// </summary>
    string GetOutputDirectory(string reportName);

    /// <summary>
    /// Deletes left-over report files (*.xlsx and unfinished *.xlsx.tmp) older than <paramref name="olderThan"/>
    /// from every configured output folder. Use it on start-up or on a timer to clean up files that were
    /// not deleted, for example after a crash. Sub-folders are not touched.
    /// </summary>
    /// <param name="olderThan">Minimum file age. Default: "Output:CleanupFilesOlderThan".</param>
    /// <returns>The number of files deleted.</returns>
    int DeleteOldReportFiles(TimeSpan? olderThan = null);
}
