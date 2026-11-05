using DRMS.ExcelReporting;
using DRMS.ExcelReporting.Configuration;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Options;

// Placed in the Microsoft namespace so "services.AddExcelReporting(...)" is found without an extra using.
namespace Microsoft.Extensions.DependencyInjection;

/// <summary>
/// Registers the Excel report generator in the dependency injection container.
/// </summary>
public static class ExcelReportingServiceCollectionExtensions
{
    /// <summary>
    /// Registers <see cref="IExcelReportGenerator"/> and binds its settings from configuration.
    /// </summary>
    /// <param name="services">The service collection.</param>
    /// <param name="configuration">The application configuration (usually builder.Configuration).</param>
    /// <param name="sectionName">The configuration section to read. Default: "ExcelReporting".</param>
    /// <example>
    /// <code>
    /// builder.Services.AddExcelReporting(builder.Configuration);
    /// </code>
    /// </example>
    public static IServiceCollection AddExcelReporting(
        this IServiceCollection services,
        IConfiguration configuration,
        string sectionName = ExcelReportingOptions.SectionName)
    {
        ArgumentNullException.ThrowIfNull(services);
        ArgumentNullException.ThrowIfNull(configuration);

        services
            .AddOptions<ExcelReportingOptions>()
            .Bind(configuration.GetSection(sectionName))
            .ValidateOnStart(); // Configuration mistakes are reported when the app starts.

        services.TryAddEnumerable(ServiceDescriptor.Singleton<IValidateOptions<ExcelReportingOptions>, ExcelReportingOptionsValidator>());
        services.TryAddSingleton<IExcelReportGenerator, ExcelReportGenerator>();

        return services;
    }
}
