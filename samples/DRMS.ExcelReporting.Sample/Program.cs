// Sample: build Excel reports from DTOs using the layout in appsettings.json.
//
// Run:  dotnet run --project samples/DRMS.ExcelReporting.Sample
// Output files are written to the "output" folder next to the executable
// (or to the folder passed as the first argument).

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
Directory.CreateDirectory(outputFolder);

// 2. Get your data (normally from a database or an API).
var employees = SampleData.Employees();
var departments = SampleData.Departments();

// 3a. A report with several sheets: give each sheet its own list.
//     The names match "DataSource" (or the sheet "Name") in appsettings.json.
var reportData = new ExcelReportData()
    .Add("Employees", employees)
    .Add("ActiveEmployees", employees.Where(e => e.IsActive))
    .Add("Departments", departments);

var employeeReportPath = Path.Combine(outputFolder, generator.GetFileName("EmployeeReport"));
File.WriteAllBytes(employeeReportPath, generator.Generate("EmployeeReport", reportData));
Console.WriteLine($"Created {employeeReportPath}");

// 3b. A report where every sheet uses the same list. You can also write straight to a stream.
var contactsPath = Path.Combine(outputFolder, generator.GetFileName("ContactList"));
using (var file = File.Create(contactsPath))
{
    generator.Generate("ContactList", ExcelReportData.ForAllSheets(employees), file);
}

Console.WriteLine($"Created {contactsPath}");

// In an ASP.NET Core API you would return the file like this:
//   app.MapGet("/reports/employees", (IExcelReportGenerator generator) =>
//       Results.File(generator.Generate("EmployeeReport", data),
//                    ExcelReportGenerator.ContentType,
//                    generator.GetFileName("EmployeeReport")));

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
