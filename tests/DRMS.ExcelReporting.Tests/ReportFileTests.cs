using ClosedXML.Excel;
using DRMS.ExcelReporting.Configuration;
using DRMS.ExcelReporting.Internal;

namespace DRMS.ExcelReporting.Tests;

/// <summary>
/// Tests for saving reports to disk, reading them back as streams and cleaning up.
/// Each test works in its own temporary folder.
/// </summary>
public sealed class ReportFileTests : IDisposable
{
    private sealed record Row(string Name, int Value);

    /// <summary>A clock that always returns the same time, so file names are predictable.</summary>
    private sealed class FixedTimeProvider(DateTimeOffset now) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => now;

        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }

    private static readonly DateTimeOffset Now = new(2026, 3, 4, 5, 6, 7, TimeSpan.Zero);

    private readonly string _root = Path.Combine(Path.GetTempPath(), "ExcelReportingTests", Guid.NewGuid().ToString("N"));

    public ReportFileTests() => Directory.CreateDirectory(_root);

    public void Dispose()
    {
        if (Directory.Exists(_root))
        {
            Directory.Delete(_root, recursive: true);
        }
    }

    private ExcelReportingOptions Options(string fileName = "", string outputDirectory = "")
    {
        var options = new ExcelReportingOptions();
        options.Defaults.Output.Directory = _root;
        options.Reports["Rows"] = new ReportDefinition
        {
            FileName = fileName,
            OutputDirectory = outputDirectory,
            Sheets = [new SheetDefinition { Name = "Data", Columns = [new() { Property = "Name" }, new() { Property = "Value" }] }],
        };

        return options;
    }

    private static ExcelReportGenerator Generator(ExcelReportingOptions options) =>
        new(options, timeProvider: new FixedTimeProvider(Now));

    private static ExcelReportData Data() => ExcelReportData.ForAllSheets(new[] { new Row("a", 1), new Row("b", 2) });

    [Fact]
    public void Saves_file_using_configured_folder_and_file_name_pattern()
    {
        var generator = Generator(Options(fileName: "{ReportName}_{Timestamp}", outputDirectory: Path.Combine(_root, "{ReportName}")));

        using var file = generator.GenerateToFile("Rows", Data());

        Assert.Equal(Path.Combine(_root, "Rows", "Rows_20260304_050607.xlsx"), file.FilePath);
        Assert.Equal("Rows_20260304_050607.xlsx", file.FileName);
        Assert.Equal(Path.Combine(_root, "Rows"), file.DirectoryPath);
        Assert.Equal(ExcelReportGenerator.ContentType, file.ContentType);
        Assert.True(File.Exists(file.FilePath));
        Assert.Equal(new FileInfo(file.FilePath).Length, file.Length);
        Assert.Equal(file.DirectoryPath, generator.GetOutputDirectory("Rows"));
    }

    [Fact]
    public void Default_file_name_pattern_is_used_when_report_has_none()
    {
        using var file = Generator(Options()).GenerateToFile("Rows", Data());

        Assert.Equal("Rows_20260304_050607.xlsx", file.FileName);
        Assert.Equal(_root, file.DirectoryPath);
    }

    [Fact]
    public void Folder_and_file_name_can_be_set_per_call()
    {
        var folder = Path.Combine(_root, "custom");

        using var file = Generator(Options()).GenerateToFile("Rows", Data(), folder, "My {Date:yyyy-MM-dd} report");

        Assert.Equal(Path.Combine(folder, "My 2026-03-04 report.xlsx"), file.FilePath);
    }

    [Fact]
    public void File_name_cannot_escape_the_output_folder()
    {
        using var file = Generator(Options()).GenerateToFile("Rows", Data(), fileName: "../../evil");

        Assert.Equal(_root, file.DirectoryPath);
        Assert.DoesNotContain("/", file.FileName);
        Assert.DoesNotContain("\\", file.FileName);
    }

    [Fact]
    public void Open_read_returns_the_complete_workbook()
    {
        using var file = Generator(Options()).GenerateToFile("Rows", Data());
        using var stream = file.OpenRead();

        Assert.Equal(file.Length, stream.Length);

        using var workbook = new XLWorkbook(stream);
        Assert.Equal("b", workbook.Worksheet("Data").Cell(3, 1).GetText());
    }

    [Fact]
    public void Dispose_deletes_the_file_and_the_folders_created_for_it()
    {
        var generator = Generator(Options(outputDirectory: Path.Combine(_root, "new", "nested")));

        string path;
        using (var file = generator.GenerateToFile("Rows", Data()))
        {
            path = file.FilePath;
            Assert.True(File.Exists(path));
        }

        Assert.False(File.Exists(path));
        Assert.False(Directory.Exists(Path.Combine(_root, "new")));
        Assert.True(Directory.Exists(_root)); // Existed before, so it is kept.
    }

    [Fact]
    public void Folders_that_still_contain_files_are_kept()
    {
        var folder = Path.Combine(_root, "shared");
        var generator = Generator(Options(outputDirectory: folder));

        var first = generator.GenerateToFile("Rows", Data());
        using var second = generator.GenerateToFile("Rows", Data());

        first.Dispose();

        Assert.True(Directory.Exists(folder));
        Assert.True(File.Exists(second.FilePath));
    }

    [Fact]
    public void File_is_kept_when_delete_on_dispose_is_off_and_delete_can_be_called_later()
    {
        var file = Generator(Options()).GenerateToFile("Rows", Data());
        file.DeleteOnDispose = false;
        file.Dispose();

        Assert.True(File.Exists(file.FilePath));

        Assert.True(file.Delete());
        Assert.False(file.Delete()); // Second call does nothing.
        Assert.True(file.IsDeleted);
        Assert.False(File.Exists(file.FilePath));
    }

    [Fact]
    public void Delete_on_dispose_default_comes_from_configuration()
    {
        var options = Options();
        options.Defaults.Output.DeleteFileOnDispose = false;

        var file = Generator(options).GenerateToFile("Rows", Data());
        file.Dispose();

        Assert.True(File.Exists(file.FilePath));
    }

    [Fact]
    public void Existing_file_gets_a_numbered_name_or_is_overwritten()
    {
        var options = Options(fileName: "Fixed");
        var generator = Generator(options);

        using var first = generator.GenerateToFile("Rows", Data());
        using var second = generator.GenerateToFile("Rows", Data());

        Assert.Equal("Fixed.xlsx", first.FileName);
        Assert.Equal("Fixed_1.xlsx", second.FileName);

        options.Defaults.Output.OverwriteExisting = true;
        var third = generator.GenerateToFile("Rows", Data());
        third.DeleteOnDispose = false;

        Assert.Equal("Fixed.xlsx", third.FileName);
    }

    [Fact]
    public void Failed_generation_leaves_no_files_or_folders_behind()
    {
        var folder = Path.Combine(_root, "failed");
        var generator = Generator(Options(outputDirectory: folder));

        Assert.Throws<InvalidOperationException>(() => generator.GenerateToFile("Rows", new ExcelReportData()));

        Assert.False(Directory.Exists(folder));
        Assert.Empty(Directory.EnumerateFileSystemEntries(_root));
    }

    [Fact]
    public void Generate_to_stream_returns_a_stream_at_position_zero()
    {
        using var stream = Generator(Options()).GenerateToStream("Rows", Data());

        Assert.Equal(0, stream.Position);
        Assert.True(stream.Length > 0);

        using var workbook = new XLWorkbook(stream);
        Assert.Equal("a", workbook.Worksheet("Data").Cell(2, 1).GetText());
        Assert.Empty(Directory.EnumerateFileSystemEntries(_root)); // Nothing written to disk.
    }

    [Fact]
    public void Delete_old_report_files_removes_only_old_report_files()
    {
        var reportFolder = Path.Combine(_root, "reports");
        var generator = Generator(Options(outputDirectory: reportFolder));
        Directory.CreateDirectory(reportFolder);

        string Create(string folder, string name, TimeSpan age)
        {
            var path = Path.Combine(folder, name);
            File.WriteAllText(path, "x");
            File.SetLastWriteTimeUtc(path, (Now - age).UtcDateTime);
            return path;
        }

        var oldReport = Create(reportFolder, "old.xlsx", TimeSpan.FromDays(2));
        var oldTemp = Create(reportFolder, ".abc.xlsx.tmp", TimeSpan.FromDays(2));
        var oldDefaultFolderReport = Create(_root, "old-default.xlsx", TimeSpan.FromDays(2));
        var newReport = Create(reportFolder, "new.xlsx", TimeSpan.FromHours(1));
        var otherFile = Create(reportFolder, "keep.txt", TimeSpan.FromDays(10));

        var deleted = generator.DeleteOldReportFiles();

        Assert.Equal(3, deleted);
        Assert.False(File.Exists(oldReport));
        Assert.False(File.Exists(oldTemp));
        Assert.False(File.Exists(oldDefaultFolderReport));
        Assert.True(File.Exists(newReport));
        Assert.True(File.Exists(otherFile));

        Assert.Equal(1, generator.DeleteOldReportFiles(TimeSpan.FromMinutes(30)));
        Assert.False(File.Exists(newReport));
    }

    [Theory]
    [InlineData("{ReportName}_{Guid:8}", 13)]
    [InlineData("{ReportName}_{Guid}", 37)]
    public void Guid_placeholder_makes_unique_names(string pattern, int expectedLengthWithoutExtension)
    {
        var a = OutputPathBuilder.BuildFileName(pattern, "Rows", Now);
        var b = OutputPathBuilder.BuildFileName(pattern, "Rows", Now);

        Assert.NotEqual(a, b);
        Assert.Equal(expectedLengthWithoutExtension, Path.GetFileNameWithoutExtension(a).Length);
    }

    [Fact]
    public void Unknown_placeholders_are_reported_by_the_validator()
    {
        var options = Options(fileName: "{ReportName}_{Nope}", outputDirectory: "{TempPath}/{Timestamp}");

        var result = new ExcelReportingOptionsValidator().Validate(null, options);

        Assert.True(result.Failed);
        Assert.Contains(result.Failures!, f => f.Contains("FileName") && f.Contains("{Nope}"));
        Assert.Contains(result.Failures!, f => f.Contains("OutputDirectory") && f.Contains("{Timestamp}"));
    }
}
