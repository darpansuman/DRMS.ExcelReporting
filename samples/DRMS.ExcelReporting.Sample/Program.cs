// Sample: build Excel reports from DTOs using the layout in appsettings.json,
// save them to disk or to a stream, "upload" them, and clean up the local files.
//
// Run:  dotnet run --project samples/DRMS.ExcelReporting.Sample [folder-for-kept-files]
//
// The "upload" here only copies the stream into a local "uploaded" folder. Replace
// UploadAsync with your storage client (for example MinIO's PutObjectAsync) in your own project.

using DRMS.ExcelReporting;
using DRMS.ExcelReporting.Sample.Dtos;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

// ContentRootPath makes sure appsettings.json is found next to the executable,
// even when the app is started from another folder.
var builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings
{
    Args = args,
    ContentRootPath = AppContext.BaseDirectory,
});

// 1. Register the generator. Settings come from the "ExcelReporting" section of appsettings.json.
builder.Services.AddExcelReporting(builder.Configuration);

using var host = builder.Build();

var generator = host.Services.GetRequiredService<IExcelReportGenerator>();
var outputFolder = args.Length > 0 ? args[0] : Path.Combine(AppContext.BaseDirectory, "output");

// 2. Remove files left behind by earlier runs (for example after a crash).
//    Deletes *.xlsx files older than "Output:CleanupFilesOlderThan" in the configured folders.
var removed = generator.DeleteOldReportFiles();
Console.WriteLine($"Old report files removed: {removed}");

// 3. Get your data (normally from a database or an API).
var employees = SampleData.Employees();
var reportData = new ExcelReportData()
    .Add("Employees", employees)
    .Add("ActiveEmployees", employees.Where(e => e.IsActive))
    .Add("Departments", SampleData.Departments());

// ---------------------------------------------------------------------------------------
// Option A: save to a file, upload it, then delete it.
// The folder and file name come from "Output" / the report's "OutputDirectory" and "FileName".
// ---------------------------------------------------------------------------------------
string uploadedFilePath;
await using (var reportFile = generator.GenerateToFile("EmployeeReport", reportData))
{
    Console.WriteLine();
    Console.WriteLine("Option A - file on disk");
    Console.WriteLine($"  Path:         {reportFile.FilePath}");
    Console.WriteLine($"  File name:    {reportFile.FileName}");
    Console.WriteLine($"  Size:         {reportFile.Length:N0} bytes");
    Console.WriteLine($"  Content type: {reportFile.ContentType}");

    await using (var stream = reportFile.OpenRead())
    {
        await UploadAsync(reportFile.FileName, stream, reportFile.Length, reportFile.ContentType);
    }

    uploadedFilePath = reportFile.FilePath;
} // <- the local file (and the folder created for it, if now empty) is deleted here.

Console.WriteLine($"  Local file still exists after dispose: {File.Exists(uploadedFilePath)}");

// ---------------------------------------------------------------------------------------
// Option B: choose the folder and file name in code, and keep the file.
// ---------------------------------------------------------------------------------------
using (var keptFile = generator.GenerateToFile(
    "ContactList",
    ExcelReportData.ForAllSheets(employees),
    outputDirectory: outputFolder,
    fileName: "Contacts_{Date:yyyy-MM-dd}"))
{
    keptFile.DeleteOnDispose = false; // Keep this one so you can open it.

    Console.WriteLine();
    Console.WriteLine("Option B - file kept");
    Console.WriteLine($"  Path: {keptFile.FilePath}");

    // You can also delete it yourself at any time: keptFile.Delete();
}

// Same, but with the full employee report, so you can open it and look at it.
using (var keptReport = generator.GenerateToFile("EmployeeReport", reportData, outputDirectory: outputFolder))
{
    keptReport.DeleteOnDispose = false;
    Console.WriteLine($"  Path: {keptReport.FilePath}");
}

// ---------------------------------------------------------------------------------------
// Option C: no file at all - generate into memory and upload the stream.
// ---------------------------------------------------------------------------------------
using (var memoryStream = generator.GenerateToStream("EmployeeReport", reportData))
{
    var objectName = generator.GetFileName("EmployeeReport");

    Console.WriteLine();
    Console.WriteLine("Option C - memory stream");
    Console.WriteLine($"  Object name: {objectName}, size: {memoryStream.Length:N0} bytes");

    await UploadAsync(objectName, memoryStream, memoryStream.Length, ExcelReportGenerator.ContentType);
}

// Stand-in for a real storage upload. Put your MinIO (or other storage) call here.
// A typical storage API needs exactly these four values: object name, stream, size and content type.
async Task UploadAsync(string objectName, Stream content, long size, string contentType)
{
    var uploadFolder = Path.Combine(outputFolder, "uploaded");
    Directory.CreateDirectory(uploadFolder);

    await using var target = File.Create(Path.Combine(uploadFolder, objectName));
    await content.CopyToAsync(target);

    Console.WriteLine($"  Uploaded '{objectName}' ({size:N0} bytes, {contentType}) to {uploadFolder}");
}

/// <summary>Creates sample data for the demo.</summary>
internal static class SampleData
{
    public static List<EmployeeDto> Employees() =>
    [
        new()
        {
            Id = 1,
            EmployeeCode = "00042",
            FullName = "Asha Patel",
            Email = "asha.patel@example.com",
            Address = new AddressDto { Street = "12 MG Road\r\nBlock B", City = "Pune", Country = "India" },
            JoiningDate = new DateOnly(2019, 4, 1),
            LastLogin = new DateTime(2026, 11, 4, 9, 30, 0),
            ShiftStart = new TimeOnly(9, 0),
            Salary = 85000.50m,
            BonusRate = 0.10,
            IsActive = true,
            StatusCode = "A",
            EmploymentType = EmploymentType.FullTime,
            Skills = ["C#", "SQL", "Azure"],
            Notes = "Team lead for reporting.|Mentors new joiners.",
        },
        new()
        {
            Id = 2,
            EmployeeCode = "00107",
            FullName = "John Smith",
            Email = "john.smith@example.com",
            Address = new AddressDto { Street = "221B Baker Street", City = "London", Country = "UK" },
            JoiningDate = new DateOnly(2021, 9, 15),
            LastLogin = null, // Shows "Never" (DefaultValue).
            ShiftStart = new TimeOnly(13, 30),
            Salary = 62000m,
            BonusRate = 0.075,
            IsActive = false,
            StatusCode = "L",
            EmploymentType = EmploymentType.PartTime,
            Skills = ["Excel"],
            Notes = "First line\r\nSecond line\\nThird line (stored as literal \\n)",
        },
        new()
        {
            Id = 3,
            EmployeeCode = "00230",
            FullName = "Maria Garcia",
            Email = "maria.garcia@example.com",
            Address = null, // City shows "Unknown" (DefaultValue).
            JoiningDate = new DateOnly(2024, 1, 8),
            LastLogin = new DateTime(2026, 10, 30, 17, 5, 12),
            ShiftStart = new TimeOnly(8, 0),
            Salary = 47250.75m,
            BonusRate = 0.05,
            IsActive = true,
            StatusCode = "", // Shows "Not Set" through ValueMappings.
            EmploymentType = EmploymentType.Contractor,
            Skills = [],
            Notes = null,
        },
    ];

    public static List<DepartmentDto> Departments() =>
    [
        new() { Code = "ENG", Name = "Engineering", Manager = "Asha Patel", HeadCount = 42, Budget = 1_250_000m, Locations = "Pune;London" },
        new() { Code = "FIN", Name = "Finance", Manager = "Ravi Kumar", HeadCount = 8, Budget = 300_000m, Locations = "Mumbai" },
    ];
}
