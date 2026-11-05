using System.Text;
using ClosedXML.Excel;
using DRMS.ExcelReporting.Configuration;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;

namespace DRMS.ExcelReporting.Tests;

/// <summary>
/// Checks that settings written in appsettings.json reach the generator correctly.
/// </summary>
public sealed class ConfigurationTests
{
    private const string Json = """
        {
          "ExcelReporting": {
            "Defaults": {
              "Formats": { "Currency": "\"€\"#,##0.00" },
              "LineBreaks": { "Separators": [ "<br>" ] }
            },
            "Reports": {
              "Orders": {
                "FileName": "Orders",
                "Sheets": [
                  {
                    "Name": "Order List",
                    "Columns": [
                      { "Header": "Total", "Property": "Total", "DataType": "Currency" },
                      { "Header": "State", "Property": "State", "ValueMappings": { "": "Unknown", "P": "Paid" } },
                      { "Header": "Note",  "Property": "Note" }
                    ]
                  }
                ]
              }
            }
          }
        }
        """;

    private sealed record Order(decimal Total, string State, string Note);

    private static ServiceProvider BuildServices(string json)
    {
        var configuration = new ConfigurationBuilder()
            .AddJsonStream(new MemoryStream(Encoding.UTF8.GetBytes(json)))
            .Build();

        return new ServiceCollection()
            .AddLogging()
            .AddExcelReporting(configuration)
            .BuildServiceProvider();
    }

    [Fact]
    public void Generates_report_from_json_configuration()
    {
        using var services = BuildServices(Json);
        var generator = services.GetRequiredService<IExcelReportGenerator>();

        // Report names ignore upper or lower case.
        var bytes = generator.Generate("orders", new[] { new Order(10.5m, "P", "a<br>b"), new Order(1, "", "") });

        using var workbook = new XLWorkbook(new MemoryStream(bytes));
        var sheet = workbook.Worksheet("Order List");

        Assert.Equal(10.5, sheet.Cell(2, 1).GetDouble());
        Assert.Equal("\"€\"#,##0.00", sheet.Cell(2, 1).Style.NumberFormat.Format);
        Assert.Equal("Paid", sheet.Cell(2, 2).GetText());
        Assert.Equal("Unknown", sheet.Cell(3, 2).GetText());
        Assert.Equal("a\nb", sheet.Cell(2, 3).GetText());
        Assert.Equal("Orders.xlsx", generator.GetFileName("Orders"));
    }

    [Fact]
    public void Unknown_report_name_lists_the_available_reports()
    {
        using var services = BuildServices(Json);
        var generator = services.GetRequiredService<IExcelReportGenerator>();

        var ex = Assert.Throws<KeyNotFoundException>(() => generator.Generate("Missing", Array.Empty<Order>()));
        Assert.Contains("Orders", ex.Message);
    }

    [Fact]
    public void Invalid_configuration_is_reported_by_the_validator()
    {
        const string badJson = """
            { "ExcelReporting": { "Reports": { "R": { "Sheets": [ { "Name": "A name that is far too long for Excel", "Columns": [ { "Header": "x" } ] } ] } } } }
            """;

        using var services = BuildServices(badJson);

        var ex = Assert.Throws<OptionsValidationException>(() => services.GetRequiredService<IOptions<ExcelReportingOptions>>().Value);
        Assert.Contains(ex.Failures, f => f.Contains("31 characters"));
        Assert.Contains(ex.Failures, f => f.Contains("Property is required"));
    }
}
