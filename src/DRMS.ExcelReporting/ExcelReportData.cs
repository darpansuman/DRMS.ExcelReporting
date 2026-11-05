using System.Collections;

namespace DRMS.ExcelReporting;

/// <summary>
/// The data for one report: one list of DTOs per data source (usually one per sheet).
/// </summary>
/// <example>
/// <code>
/// var data = new ExcelReportData()
///     .Add("Employees", employees)      // fills every sheet with "DataSource": "Employees"
///     .Add("Departments", departments); // fills every sheet with "DataSource": "Departments"
/// </code>
/// </example>
public sealed class ExcelReportData
{
    private readonly Dictionary<string, IEnumerable> _sources = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Data used by any sheet whose data source was not added by name.
    /// Handy when every sheet shows the same list in a different way.
    /// </summary>
    public IEnumerable? DefaultSource { get; private set; }

    /// <summary>The names of the data sources that were added.</summary>
    public IReadOnlyCollection<string> SourceNames => _sources.Keys;

    /// <summary>
    /// Adds (or replaces) a named list of DTOs. The name matches the sheet's "DataSource"
    /// (or the sheet "Name" when "DataSource" is empty). Matching ignores upper or lower case.
    /// </summary>
    public ExcelReportData Add<T>(string dataSourceName, IEnumerable<T> items)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(dataSourceName);
        ArgumentNullException.ThrowIfNull(items);

        _sources[dataSourceName] = items;
        return this;
    }

    /// <summary>
    /// Sets the list used by every sheet that has no data source of its own.
    /// </summary>
    public ExcelReportData WithDefault<T>(IEnumerable<T> items)
    {
        ArgumentNullException.ThrowIfNull(items);

        DefaultSource = items;
        return this;
    }

    /// <summary>Creates report data where every sheet uses the same list.</summary>
    public static ExcelReportData ForAllSheets<T>(IEnumerable<T> items) => new ExcelReportData().WithDefault(items);

    /// <summary>
    /// Finds the list for a data source name, falling back to <see cref="DefaultSource"/>.
    /// </summary>
    public bool TryGetSource(string dataSourceName, out IEnumerable items)
    {
        if (_sources.TryGetValue(dataSourceName, out var named))
        {
            items = named;
            return true;
        }

        if (DefaultSource is not null)
        {
            items = DefaultSource;
            return true;
        }

        items = Array.Empty<object>();
        return false;
    }
}
