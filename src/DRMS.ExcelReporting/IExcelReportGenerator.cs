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
    /// Returns the configured file name for a report, or "{reportName}.xlsx" when none is set.
    /// </summary>
    string GetFileName(string reportName);
}
