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
See [Example 5](#example-5--remove-left-over-files-on-a-timer) for a ready-made background service.

### How a file is written

1. The folder is created if needed (the library remembers which folders it created).
2. The workbook is written to a hidden temporary file `.<id>.xlsx.tmp` in that folder.
3. The temporary file is renamed to the final name, so nobody sees a half-written file.
   If the name is taken: `OverwriteExisting: true` replaces it, otherwise the file becomes `Name_1.xlsx`, `Name_2.xlsx`, …
4. If anything fails, the temporary file and any folder created for it are removed.

File names given in config or in code are cleaned: characters that are not allowed in file names
(including `/` and `\`) become `_`, so a name can never point outside the output folder.

## Examples

These examples show how to use the library from your own application code. They are not part of the
library itself. The MinIO calls use the official [Minio .NET SDK](https://github.com/minio/minio-dotnet)
(`Minio` NuGet package; checked against versions 6.0.5 and 7.0.0).

### Example 1 – Scheduled alert: generate, upload to MinIO, delete the local file

A scheduler job loads alert rows, adds an alert id per agreement, writes them to Excel, uploads the
file to MinIO and removes the local copy.

**appsettings.json**

```jsonc
{
  "ExcelReporting": {
    "Defaults": {
      "Output": {
        "Directory": "{TempPath}/ExcelReports",
        "DeleteFileOnDispose": true,        // local file is removed after the upload
        "DeleteEmptyDirectories": true,
        "CleanupFilesOlderThan": "1.00:00:00"
      }
    },
    "Reports": {
      "FastOnboardingAlert": {
        "FileName": "FastOnboardingAlert_{Timestamp}_{Guid:8}.xlsx",
        "OutputDirectory": "{TempPath}/ExcelReports/{ReportName}",
        "Sheets": [
          {
            "Name": "Fast Onboarding",
            "DataSource": "FastOnboarding",
            "Title": "Fast Onboarding Alerts",
            "Columns": [
              { "Header": "Alert Id",     "Property": "AlertId",     "DataType": "Integer" },
              { "Header": "Alert Type",   "Property": "AlertType",   "DataType": "Text" },
              { "Header": "Agreement Id", "Property": "AgreementId", "DataType": "Text" }
              // Add one line per FastOnboardingResponse property you want in the file, for example:
              // { "Header": "Onboarded At", "Property": "OnboardedAt", "DataType": "DateTime" }
            ]
          }
        ]
      }
    }
  },

  // Your own MinIO settings (not read by this library).
  "MinioStorage": {
    "Endpoint": "minio.local:9000",
    "AccessKey": "",
    "SecretKey": "",
    "UseSsl": true,
    "BucketName": "fraud-alerts",
    "ObjectPrefix": "fast-onboarding"
  }
}
```

> `"Property"` must match a real property name on `FastOnboardingResponse` (upper/lower case is ignored).
> A wrong name stops the report with a message naming the column, so you find it on the first run.
> `AlertId` is a string in the DTO; `"DataType": "Integer"` writes it as a number so Excel sorts it correctly.

**Program.cs (registration)**

```csharp
using Minio;

builder.Services.AddExcelReporting(builder.Configuration);

builder.Services.Configure<MinioStorageOptions>(builder.Configuration.GetSection("MinioStorage"));
builder.Services.AddMinio(client =>
{
    var minio = builder.Configuration.GetSection("MinioStorage").Get<MinioStorageOptions>()!;
    client.WithEndpoint(minio.Endpoint)
          .WithCredentials(minio.AccessKey, minio.SecretKey)
          .WithSSL(minio.UseSsl);
});
```

```csharp
public sealed class MinioStorageOptions
{
    public string Endpoint { get; set; } = string.Empty;
    public string AccessKey { get; set; } = string.Empty;
    public string SecretKey { get; set; } = string.Empty;
    public bool UseSsl { get; set; } = true;
    public string BucketName { get; set; } = string.Empty;
    public string ObjectPrefix { get; set; } = string.Empty;
}
```

**The scheduler service**

```csharp
using DRMS.ExcelReporting;
using Minio;
using Minio.DataModel.Args;

public class FraudAlertScheduler
{
    private const string FastOnboardingReport = "FastOnboardingAlert"; // key in "ExcelReporting:Reports"
    private const string FastOnboardingSource = "FastOnboarding";      // "DataSource" of the sheet

    private readonly IDeviceRentalProxy _deviceRentalProxy;
    private readonly IExcelReportGenerator _excelReportGenerator;
    private readonly IMinioClient _minioClient;
    private readonly MinioStorageOptions _minioOptions;
    private readonly FraudSchedulerOptions _fraudSchedulerOptions;
    private readonly ILogger<FraudAlertScheduler> _logger;

    public FraudAlertScheduler(
        IDeviceRentalProxy deviceRentalProxy,
        IExcelReportGenerator excelReportGenerator,
        IMinioClient minioClient,
        IOptions<MinioStorageOptions> minioOptions,
        IOptions<FraudSchedulerOptions> fraudSchedulerOptions,
        ILogger<FraudAlertScheduler> logger)
    {
        _deviceRentalProxy = deviceRentalProxy;
        _excelReportGenerator = excelReportGenerator;
        _minioClient = minioClient;
        _minioOptions = minioOptions.Value;
        _fraudSchedulerOptions = fraudSchedulerOptions.Value;
        _logger = logger;
    }

    public async Task<bool> SendFastOnboardingAlert(CancellationToken cancellationToken = default)
    {
        DateTimeOffset timeNow = DateTimeHelper.GetSouthAfricaTimeNowWithOffSet();

        var startDateTime = new DateTimeOffset(timeNow.Date, timeNow.Offset);
        var endDateTime = startDateTime.AddDays(1).AddTicks(-1);

        var request = new Requests.DeviceRental.GetFastOnboardingRequest
        {
            TriggerThresholdInMinutes = _fraudSchedulerOptions.FastOnboarding.TriggerThresholdInMinutes,
            StartDateTime = startDateTime,
            EndDateTime = endDateTime,
        };

        List<FastOnboardingResponse>? response = await _deviceRentalProxy.GetFastOnboardingAsync(request, cancellationToken);

        if (response is null || response.Count == 0)
        {
            _logger.LogInformation("No Results were found for SendFastOnboardingAlert! No alert generated !");
            return false;
        }

        // One alert id per agreement: 1, 2, 3, ...
        var alertNumber = 0;
        foreach (var group in response.GroupBy(x => x.AgreementId))
        {
            alertNumber++;
            foreach (var item in group)
            {
                item.AlertId = alertNumber.ToString();
                item.AlertType = _fraudSchedulerOptions.FastOnboarding.AlertType;
            }
        }

        // Sort by the number, not the text ("10" would come before "2" as text).
        response = response.OrderBy(x => int.Parse(x.AlertId)).ToList();

        // Create the Excel file, upload it, delete the local copy.
        var objectName = await UploadFastOnboardingReportAsync(response, startDateTime, cancellationToken);

        _logger.LogInformation(
            "Fast onboarding alert report with {RowCount} row(s) uploaded to '{Bucket}/{ObjectName}'.",
            response.Count, _minioOptions.BucketName, objectName);

        return true;
    }

    /// <summary>
    /// Writes the alerts to an Excel file, uploads it to MinIO and deletes the local file.
    /// Returns the MinIO object name.
    /// </summary>
    private async Task<string> UploadFastOnboardingReportAsync(
        List<FastOnboardingResponse> alerts,
        DateTimeOffset reportDate,
        CancellationToken cancellationToken)
    {
        var data = new ExcelReportData().Add(FastOnboardingSource, alerts);

        // 1. Create the file. Folder and file name come from appsettings.json.
        //    "await using" deletes the local file at the end of this method,
        //    also when the upload throws an exception.
        await using var reportFile = _excelReportGenerator.GenerateToFile(FastOnboardingReport, data);

        _logger.LogDebug("Fast onboarding report created at {FilePath} ({Size} bytes).", reportFile.FilePath, reportFile.Length);

        // 2. Object name in the bucket, e.g. "fast-onboarding/2026/11/05/FastOnboardingAlert_20261105_101500_ab12cd34.xlsx".
        var objectName = $"{_minioOptions.ObjectPrefix}/{reportDate:yyyy/MM/dd}/{reportFile.FileName}";

        // 3. Upload from the file stream.
        await using (var stream = reportFile.OpenRead())
        {
            var putObjectArgs = new PutObjectArgs()
                .WithBucket(_minioOptions.BucketName)
                .WithObject(objectName)
                .WithStreamData(stream)
                .WithObjectSize(reportFile.Length)
                .WithContentType(reportFile.ContentType);

            await _minioClient.PutObjectAsync(putObjectArgs, cancellationToken);
        }

        return objectName;
    } // 4. reportFile is disposed here -> local file and its empty folder are deleted.
}
```

What happens step by step:

| Step | Call | Result |
|---|---|---|
| 1 | `GenerateToFile("FastOnboardingAlert", data)` | `/tmp/ExcelReports/FastOnboardingAlert/FastOnboardingAlert_20261105_101500_ab12cd34.xlsx` is created. |
| 2 | `reportFile.FileName` / `FilePath` / `Length` / `ContentType` | Everything the upload needs. |
| 3 | `reportFile.OpenRead()` + `PutObjectAsync` | The file is streamed to MinIO; it is not loaded into memory. |
| 4 | end of `await using` | The local file is deleted, and so is the `FastOnboardingAlert` folder if the library created it and it is now empty. |

### Example 2 – Same upload, without a local file

When you do not need a file on disk, generate into memory. Good for small and medium reports;
the whole file is held in memory while uploading.

```csharp
private async Task<string> UploadFastOnboardingReportFromMemoryAsync(
    List<FastOnboardingResponse> alerts,
    CancellationToken cancellationToken)
{
    var data = new ExcelReportData().Add("FastOnboarding", alerts);

    using var stream = _excelReportGenerator.GenerateToStream("FastOnboardingAlert", data); // already at position 0
    var objectName = $"{_minioOptions.ObjectPrefix}/{_excelReportGenerator.GetFileName("FastOnboardingAlert")}";

    await _minioClient.PutObjectAsync(new PutObjectArgs()
        .WithBucket(_minioOptions.BucketName)
        .WithObject(objectName)
        .WithStreamData(stream)
        .WithObjectSize(stream.Length)
        .WithContentType(ExcelReportGenerator.ContentType), cancellationToken);

    return objectName;
}
```

### Example 3 – Keep the local file when the upload fails

By default the file is always deleted. If you would rather keep it for a retry or to look at it,
switch off the delete when the upload throws:

```csharp
await using var reportFile = _excelReportGenerator.GenerateToFile("FastOnboardingAlert", data);

try
{
    await using var stream = reportFile.OpenRead();
    await _minioClient.PutObjectAsync(/* ... as in Example 1 ... */, cancellationToken);
}
catch (Exception ex)
{
    reportFile.DeleteOnDispose = false; // keep it
    _logger.LogError(ex, "Upload failed. The report was kept at {FilePath}.", reportFile.FilePath);
    throw;
}
```

Kept files are removed later by `DeleteOldReportFiles()` (see Example 5) once they are older than
`Output:CleanupFilesOlderThan`.

### Example 4 – Choose the folder or file name in code

The two optional arguments override the configured folder and file name for one call only.
Placeholders work here too.

```csharp
await using var reportFile = _excelReportGenerator.GenerateToFile(
    "FastOnboardingAlert",
    data,
    outputDirectory: "/var/app/reports/{ReportName}",
    fileName: $"FastOnboarding_{startDateTime:yyyyMMdd}_{{Guid:8}}");   // {{ }} = a literal { } in C# interpolation

Console.WriteLine(reportFile.FilePath); // /var/app/reports/FastOnboardingAlert/FastOnboarding_20261029_1a2b3c4d.xlsx
```

### Example 5 – Remove left-over files on a timer

If the process stops between creating and disposing a file, the file stays on disk.
A small background service removes such files regularly:

```csharp
public sealed class ReportFileCleanupService(IExcelReportGenerator generator, ILogger<ReportFileCleanupService> logger)
    : BackgroundService
{
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        using var timer = new PeriodicTimer(TimeSpan.FromHours(1));

        do
        {
            try
            {
                generator.DeleteOldReportFiles(); // uses Output:CleanupFilesOlderThan
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "Cleaning up old report files failed.");
            }
        }
        while (await timer.WaitForNextTickAsync(stoppingToken));
    }
}

// Program.cs
builder.Services.AddHostedService<ReportFileCleanupService>();
```

### Example 6 – Several sheets in one alert file

Each sheet reads its own list. Here a summary sheet (one row per agreement) is added next to the
detail sheet. Add a second sheet with `"DataSource": "FastOnboardingSummary"` to the report in
appsettings.json, then:

```csharp
var summary = response
    .GroupBy(x => x.AgreementId)
    .Select(g => new { AgreementId = g.Key, AlertId = g.First().AlertId, Rows = g.Count() })
    .ToList();

var data = new ExcelReportData()
    .Add("FastOnboarding", response)          // detail sheet
    .Add("FastOnboardingSummary", summary);   // summary sheet (anonymous types work too)

await using var reportFile = _excelReportGenerator.GenerateToFile("FastOnboardingAlert", data);
```

### Example 7 – Return the report from a Web API

```csharp
app.MapGet("/reports/fast-onboarding", async (IExcelReportGenerator generator, IDeviceRentalProxy proxy, CancellationToken ct) =>
{
    var rows = await proxy.GetFastOnboardingAsync(/* request */, ct);
    var bytes = generator.Generate("FastOnboardingAlert", new ExcelReportData().Add("FastOnboarding", rows));

    return Results.File(bytes, ExcelReportGenerator.ContentType, generator.GetFileName("FastOnboardingAlert"));
});
```

### Example 8 – A report defined in code (no appsettings.json)

Useful for one-off exports or tests.

```csharp
var report = new ReportDefinition
{
    FileName = "AdHoc_{Timestamp}",
    Sheets =
    [
        new SheetDefinition
        {
            Name = "Alerts",
            Columns =
            [
                new ColumnDefinition { Header = "Alert Id", Property = "AlertId", DataType = ExcelDataType.Integer },
                new ColumnDefinition { Header = "Agreement", Property = "AgreementId" },
            ],
        },
    ],
};

await using var reportFile = generator.GenerateToFile(report, ExcelReportData.ForAllSheets(response));
```

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
