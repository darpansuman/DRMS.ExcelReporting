# DRMS.ExcelReporting

A small, reusable .NET 10 library that turns lists of DTOs into Excel (`.xlsx`) files.
The whole layout of each report lives in `appsettings.json`:

- column header text
- which DTO property fills each column (including nested ones like `Address.City`)
- the Excel data type and number format of each column
- value mappings (for example `true` → `Yes`, `A` → `Active`)
- sheet names, and which list fills each sheet
- line break rules inside cells
- where the file is saved (folder and file name pattern) and when it is deleted again

Change a report by editing configuration — no code change, no rebuild.

Excel files are written with [ClosedXML](https://github.com/ClosedXML/ClosedXML) (MIT licence, no Excel installation needed).

## Project layout

```
src/DRMS.ExcelReporting/                  The reusable library (copy this folder)
  ExcelReportGenerator.cs                 Main class: builds the workbook
  IExcelReportGenerator.cs                Interface (easy to mock in tests)
  ExcelReportData.cs                      Holds one list of DTOs per sheet
  ExcelReportFile.cs                      A saved report file: path, size, read stream, auto-delete
  Configuration/
    ExcelReportingOptions.cs              Classes that mirror the appsettings.json section
    ExcelDataType.cs                      Supported Excel data types
    ExcelReportingOptionsValidator.cs     Finds configuration mistakes at startup
  Internal/
    PropertyValueResolver.cs              Reads "Address.City" style paths from a DTO (cached, fast)
    CellValueConverter.cs                 Value mappings, defaults, data type conversion
    LineBreakHelper.cs                    Turns separators and \r\n into Excel new lines
    OutputPathBuilder.cs                  Fills in {ReportName}/{Timestamp}/... and makes file names safe
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

It shows the three ways to get a report out (file + delete after upload, file kept, memory stream) and
keeps a couple of files in `samples/DRMS.ExcelReporting.Sample/bin/Debug/net10.0/output/` so you can open them.

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

## Saving to disk, uploading and cleaning up

The generator can hand the report over in three ways. Pick the one your storage client (MinIO, S3,
Azure Blob, a network share, …) needs. This library does not upload anything itself.

| Method | Returns | Use when |
|---|---|---|
| `GenerateToFile(reportName, data, outputDirectory?, fileName?)` | `ExcelReportFile` | You need a real file path, or want the file on disk while uploading. |
| `GenerateToStream(reportName, data)` | `MemoryStream` at position 0 | You want to upload straight from memory, without touching the disk. |
| `Generate(reportName, data, stream)` | — | You already have a target stream (HTTP response, file, …). |

### `ExcelReportFile`

| Member | Meaning |
|---|---|
| `FilePath` | Full path of the saved file. |
| `FileName` | File name only, e.g. `EmployeeReport_20261105_101500.xlsx` — handy as the object name. |
| `DirectoryPath` | Folder that holds the file. |
| `Length` | Size in bytes (most upload APIs ask for it). |
| `ContentType` | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` |
| `OpenRead()` | Opens a read stream for the upload. The file can still be deleted while it is open. |
| `Delete()` | Deletes the file now (safe to call more than once). |
| `DeleteOnDispose` | When `true` (default from config), disposing the object deletes the file. Set `false` to keep it. |
| `Dispose()` / `DisposeAsync()` | Deletes the file (if `DeleteOnDispose`) and the folders created for it, if now empty. |

Typical use — the local file is always removed, even if the upload throws:

```csharp
await using var report = generator.GenerateToFile("EmployeeReport", data);

await using (var stream = report.OpenRead())
{
    // Your storage call goes here, for example:
    // await storage.UploadAsync(bucket, report.FileName, stream, report.Length, report.ContentType);
}
// report is disposed here -> local file deleted.
```

Without a file:

```csharp
using var stream = generator.GenerateToStream("EmployeeReport", data);
var objectName = generator.GetFileName("EmployeeReport");
// await storage.UploadAsync(bucket, objectName, stream, stream.Length, ExcelReportGenerator.ContentType);
```

### Cleaning up left-over files

If the application stops before a file is disposed, the file stays on disk. Call
`generator.DeleteOldReportFiles()` at start-up or on a timer: it deletes `*.xlsx` (and unfinished
`*.xlsx.tmp`) files older than `Output:CleanupFilesOlderThan` from every configured output folder.
Sub-folders and other file types are never touched — still, **use a folder that only holds reports**.

### How a file is written

1. The folder is created if needed (the library remembers which folders it created).
2. The workbook is written to a hidden temporary file `.<id>.xlsx.tmp` in that folder.
3. The temporary file is renamed to the final name, so nobody sees a half-written file.
   If the name is taken: `OverwriteExisting: true` replaces it, otherwise the file becomes `Name_1.xlsx`, `Name_2.xlsx`, …
4. If anything fails, the temporary file and any folder created for it are removed.

File names given in config or in code are cleaned: characters that are not allowed in file names
(including `/` and `\`) become `_`, so a name can never point outside the output folder.

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
| `Output.Directory` | `{TempPath}/ExcelReports` | Folder for `GenerateToFile`. Relative paths are relative to the app folder; `%VARS%` are expanded. |
| `Output.FileName` | `{ReportName}_{Timestamp}` | File name when a report has no `FileName`. `.xlsx` is added when missing. |
| `Output.OverwriteExisting` | `false` | `false`: an existing name becomes `Name_1.xlsx`, `Name_2.xlsx`, … |
| `Output.DeleteFileOnDispose` | `true` | Delete the file when the `ExcelReportFile` is disposed. |
| `Output.DeleteEmptyDirectories` | `true` | Also delete folders that were created for the file, once they are empty. Folders that already existed are never deleted. |
| `Output.CleanupFilesOlderThan` | `1.00:00:00` | Age used by `DeleteOldReportFiles()` (`d.hh:mm:ss`). |

#### Path placeholders

| Placeholder | Where | Value |
|---|---|---|
| `{ReportName}` | folder, file name | The report key, e.g. `EmployeeReport`. |
| `{TempPath}` | folder | The system temp folder. |
| `{BaseDirectory}` | folder | The application folder. |
| `{Date}` / `{Time}` / `{Timestamp}` | file name | Local time: `yyyyMMdd` / `HHmmss` / `yyyyMMdd_HHmmss`. Custom: `{Date:yyyy-MM-dd}`. |
| `{Guid}` | file name | A new unique id (32 characters). `{Guid:8}` keeps the first 8. |

Unknown placeholders are reported at start-up.

### Report

| Setting | Meaning |
|---|---|
| key (e.g. `"EmployeeReport"`) | The name you pass to `Generate(...)`. Upper/lower case is ignored. |
| `FileName` | File name pattern for this report, e.g. `EmployeeReport_{Timestamp}.xlsx`. Empty = `Output.FileName`. Also returned by `GetFileName(...)`. |
| `OutputDirectory` | Folder for this report's files. Empty = `Output.Directory`. |
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
