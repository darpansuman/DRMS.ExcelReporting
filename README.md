# DRMS.ExcelReporting

A small, reusable .NET 10 library that turns lists of DTOs into Excel (`.xlsx`) files.
The whole layout of each report lives in `appsettings.json`:

- column header text
- which DTO property fills each column (including nested ones like `Address.City`)
- the Excel data type and number format of each column
- value mappings (for example `true` → `Yes`, `A` → `Active`)
- sheet names, and which list fills each sheet
- line break rules inside cells

Change a report by editing configuration — no code change, no rebuild.

Excel files are written with [ClosedXML](https://github.com/ClosedXML/ClosedXML) (MIT licence, no Excel installation needed).

## Project layout

```
src/DRMS.ExcelReporting/                  The reusable library (copy this folder)
  ExcelReportGenerator.cs                 Main class: builds the workbook
  IExcelReportGenerator.cs                Interface (easy to mock in tests)
  ExcelReportData.cs                      Holds one list of DTOs per sheet
  Configuration/
    ExcelReportingOptions.cs              Classes that mirror the appsettings.json section
    ExcelDataType.cs                      Supported Excel data types
    ExcelReportingOptionsValidator.cs     Finds configuration mistakes at startup
  Internal/
    PropertyValueResolver.cs              Reads "Address.City" style paths from a DTO (cached, fast)
    CellValueConverter.cs                 Value mappings, defaults, data type conversion
    LineBreakHelper.cs                    Turns separators and \r\n into Excel new lines
  DependencyInjection/
    ServiceCollectionExtensions.cs        services.AddExcelReporting(configuration)
samples/DRMS.ExcelReporting.Sample/       Console demo with appsettings.json and sample DTOs
tests/DRMS.ExcelReporting.Tests/          xUnit tests
```

## Quick start

Run the sample:

```bash
dotnet run --project samples/DRMS.ExcelReporting.Sample
```

It writes `EmployeeReport.xlsx` (3 sheets) and `Contacts.xlsx` into `samples/DRMS.ExcelReporting.Sample/bin/Debug/net10.0/output/`.

Run the tests:

```bash
dotnet test
```

## Using it in your own project

1. Copy `src/DRMS.ExcelReporting` into your solution and add a project reference
   (or copy the `.cs` files into an existing project and add the `ClosedXML` NuGet package).
2. Add an `ExcelReporting` section to your `appsettings.json` (see below).
3. Register the generator:

   ```csharp
   builder.Services.AddExcelReporting(builder.Configuration);
   ```

4. Use it:

   ```csharp
   public class ReportsController(IExcelReportGenerator generator) : ControllerBase
   {
       [HttpGet("employees")]
       public IActionResult Employees()
       {
           var data = new ExcelReportData()
               .Add("Employees", employeeService.GetAll())    // matches "DataSource": "Employees"
               .Add("Departments", departmentService.GetAll()); // matches the "Departments" sheet

           return File(
               generator.Generate("EmployeeReport", data),
               ExcelReportGenerator.ContentType,
               generator.GetFileName("EmployeeReport"));
       }
   }
   ```

Without dependency injection:

```csharp
var options = configuration.GetSection("ExcelReporting").Get<ExcelReportingOptions>()!;
var generator = new ExcelReportGenerator(options);
byte[] file = generator.Generate("EmployeeReport", employees);
```

You can also build a `ReportDefinition` in code and pass it to `Generate(report, data)`.

## Configuration reference

A minimal report:

```json
"ExcelReporting": {
  "Reports": {
    "EmployeeReport": {
      "Sheets": [
        {
          "Name": "Employees",
          "Columns": [
            { "Header": "Name", "Property": "FullName" },
            { "Header": "City", "Property": "Address.City" },
            { "Header": "Salary", "Property": "Salary", "DataType": "Currency" }
          ]
        }
      ]
    }
  }
}
```

See [`samples/DRMS.ExcelReporting.Sample/appsettings.json`](samples/DRMS.ExcelReporting.Sample/appsettings.json) for a full example with every option.

### `Defaults` (shared by all reports)

| Setting | Default | Meaning |
|---|---|---|
| `Culture` | invariant | Culture used to read numbers/dates that arrive as text, e.g. `"en-US"`, `"de-DE"`. |
| `Formats.Integer` / `Number` / `Currency` / `Percentage` / `Date` / `DateTime` / `Time` | `0`, `#,##0.00`, `#,##0.00`, `0.00%`, `yyyy-mm-dd`, `yyyy-mm-dd hh:mm:ss`, `hh:mm:ss` | Default Excel number format per data type. |
| `HeaderStyle.Bold` / `BackgroundColor` / `FontColor` / `WrapText` | `true`, `#D9E1F2`, none, `true` | Look of the header row. |
| `LineBreaks.Separators` | `[]` | Text markers that become a new line, e.g. `["|", "<br>"]`. |
| `LineBreaks.CollectionSeparator` | `"\n"` | Joins list properties; the default puts each item on its own line. |
| `LineBreaks.WrapTextWhenLineBreakFound` | `true` | Turn on "Wrap Text" for any cell that has a line break. |
| `AutoFitColumns` / `MaxColumnWidth` | `true`, `60` | Fit columns to content (first 1,000 rows are measured), with an upper limit. |
| `FreezeHeaderRow` / `AutoFilter` | `true`, `true` | Freeze the header and add filter drop-downs. |
| `MaxRowsPerSheet` | `1048575` | More rows continue on extra sheets: `Name (2)`, `Name (3)`, … |

### Report

| Setting | Meaning |
|---|---|
| key (e.g. `"EmployeeReport"`) | The name you pass to `Generate(...)`. Upper/lower case is ignored. |
| `FileName` | Returned by `GetFileName(...)`. `.xlsx` is added if missing. |
| `Sheets` | The sheets, in order. |

### Sheet

| Setting | Meaning |
|---|---|
| `Name` (required) | Tab name. Max 31 characters; must not contain `: \ / ? * [ ]`; unique in the report. |
| `DataSource` | Key of the list in `ExcelReportData`. Empty = use the sheet `Name`. |
| `Title` | Optional bold title row above the header. |
| `FreezeHeaderRow`, `AutoFilter` | Override the defaults for this sheet. |
| `Columns` | The columns, in order. |

### Column

| Setting | Meaning |
|---|---|
| `Property` (required) | DTO property to read. Dots for nested objects: `Address.City`. Case is ignored. Public fields and dictionary keys also work. |
| `Header` | Header text. `\n` splits it over two lines. Empty = property name. |
| `DataType` | See the table below. Default `Auto`. |
| `Format` | Excel number format for this column, e.g. `dd-mmm-yyyy`, `"$"#,##0.00`. For `Text` columns it is a .NET format string (e.g. `N2`). |
| `ValueMappings` | Replace raw values with text: `{ "A": "Active", "": "Not set" }`. The key `""` matches null/empty values. Case is ignored. |
| `DefaultValue` | Written when the value is null or empty. |
| `Width` | Fixed width (characters). When not set, the column is auto-fitted. |
| `WrapText` | Always wrap this column. |
| `Alignment` | `Left`, `Center` or `Right`. |
| `LineBreaks` | Line break rules for this column only (same fields as `Defaults.LineBreaks`). |

### Data types

| `DataType` | Accepts | Written as |
|---|---|---|
| `Auto` | anything | Picked from the value: numbers → number, dates → date, `bool` → TRUE/FALSE, rest → text. |
| `Text` | anything | Text. Keeps leading zeros (`"00042"`). Enums use their name. |
| `Integer` | numbers, numeric text | Rounded number, format `0`. |
| `Number`, `Currency`, `Percentage` | numbers, numeric text | Number with the configured format. `0.25` shows as `25.00%`. |
| `Date`, `DateTime` | `DateTime`, `DateOnly`, `DateTimeOffset`, date text | Real Excel date (sortable, filterable). |
| `Time` | `TimeOnly`, `TimeSpan`, `DateTime`, time text | Real Excel time. |
| `Boolean` | `bool`, `0/1`, `yes/no`, `true/false` | TRUE / FALSE. For "Yes"/"No" use `Text` + `ValueMappings`. |
| `Hyperlink` | URL or e-mail address | Clickable link (e-mails get `mailto:`). |

## How values are processed

For every cell, in this order:

1. **Read** the property from the DTO (`Address.City`). A null anywhere on the path gives an empty value.
2. **Lists** (`List<string>`, arrays, …) are joined with `CollectionSeparator` — one item per line by default.
3. **ValueMappings** replace the raw value if a key matches.
4. **DefaultValue** fills empty values.
5. **Convert** to the column's `DataType` and apply the number format.
6. **Line breaks**: `\r\n` and `\r` become `\n` (the only new line Excel understands), and every configured
   separator becomes `\n`. If the cell now has a line break, "Wrap Text" is turned on so Excel shows it.

If a value cannot be converted (for example `"abc"` in a `Number` column), it is written as text and one
warning per column is logged — one bad value never breaks the whole report. Text that comes from
`ValueMappings` or `DefaultValue` (for example `"Never"` in a `DateTime` column) is written as text without a warning.

Configuration mistakes (unknown property name, bad sheet name, missing data source, unknown report)
throw an exception with a clear message. With `AddExcelReporting`, sheet and column settings are also
checked when the application starts.

## Notes

- The generator is thread-safe; it is registered as a singleton.
- With dependency injection the latest configuration is read on every call, so edits to `appsettings.json`
  apply without a restart (when the file is loaded with `reloadOnChange`, the default in ASP.NET Core).
- The list for each sheet is read only once, so a streamed `IEnumerable<T>` (for example from a database) works.
- `ClosedXML` keeps the whole workbook in memory. For very large exports (hundreds of thousands of rows),
  test the memory use for your data.
