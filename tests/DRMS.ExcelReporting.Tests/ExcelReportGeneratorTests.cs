using ClosedXML.Excel;
using DRMS.ExcelReporting.Configuration;

namespace DRMS.ExcelReporting.Tests;

public sealed class ExcelReportGeneratorTests
{
    private sealed class Address
    {
        public string City { get; init; } = string.Empty;
    }

    private enum Level { Junior, Senior }

    private sealed class Person
    {
        public int Id { get; init; }
        public string Code { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public Address? Address { get; init; }
        public DateOnly Born { get; init; }
        public DateTime? Updated { get; init; }
        public TimeOnly Start { get; init; }
        public decimal Salary { get; init; }
        public double Rate { get; init; }
        public bool Active { get; init; }
        public string Status { get; init; } = string.Empty;
        public Level Level { get; init; }
        public List<string> Tags { get; init; } = [];
        public string? Notes { get; init; }
        public string Website { get; init; } = string.Empty;
        public string Amount { get; init; } = string.Empty;
    }

    private static Person SamplePerson(int id = 1) => new()
    {
        Id = id,
        Code = "007",
        Name = "Ann",
        Address = new Address { City = "Pune" },
        Born = new DateOnly(1990, 5, 17),
        Updated = new DateTime(2026, 1, 2, 3, 4, 5),
        Start = new TimeOnly(9, 30),
        Salary = 1234.5m,
        Rate = 0.25,
        Active = true,
        Status = "A",
        Level = Level.Senior,
        Tags = ["one", "two"],
        Notes = "line1\r\nline2|line3",
        Website = "https://example.com",
        Amount = "1,234.56",
    };

    private static ReportDefinition SingleSheet(params ColumnDefinition[] columns) => new()
    {
        Sheets = [new SheetDefinition { Name = "Data", Columns = [.. columns] }],
    };

    private static XLWorkbook Open(byte[] bytes) => new(new MemoryStream(bytes));

    private static IXLCell DataCell(byte[] bytes, int column = 1, string sheet = "Data") =>
        Open(bytes).Worksheet(sheet).Cell(2, column);

    private static byte[] Generate(ReportDefinition report, params object[] items) =>
        new ExcelReportGenerator(new ExcelReportingOptions()).Generate(report, ExcelReportData.ForAllSheets(items));

    [Fact]
    public void Writes_header_and_maps_properties_including_nested_paths()
    {
        var report = SingleSheet(
            new ColumnDefinition { Header = "Name", Property = "Name" },
            new ColumnDefinition { Header = "City", Property = "Address.City" },
            new ColumnDefinition { Property = "id" }); // No header: property name is used; case is ignored.

        var people = new[] { SamplePerson(), new Person { Id = 2, Name = "Bob", Address = null } };
        var sheet = Open(Generate(report, people)).Worksheet("Data");

        Assert.Equal("Name", sheet.Cell(1, 1).GetText());
        Assert.Equal("City", sheet.Cell(1, 2).GetText());
        Assert.Equal("id", sheet.Cell(1, 3).GetText());
        Assert.Equal("Ann", sheet.Cell(2, 1).GetText());
        Assert.Equal("Pune", sheet.Cell(2, 2).GetText());
        Assert.Equal(1, sheet.Cell(2, 3).GetDouble());
        Assert.True(sheet.Cell(3, 2).IsEmpty()); // Null nested object gives an empty cell, not an error.
    }

    [Theory]
    [InlineData("Id", ExcelDataType.Integer, "0")]
    [InlineData("Salary", ExcelDataType.Number, "#,##0.00")]
    [InlineData("Salary", ExcelDataType.Currency, "#,##0.00")]
    [InlineData("Rate", ExcelDataType.Percentage, "0.00%")]
    public void Writes_numeric_types_as_numbers_with_default_formats(string property, ExcelDataType type, string expectedFormat)
    {
        var cell = DataCell(Generate(SingleSheet(new ColumnDefinition { Property = property, DataType = type }), SamplePerson()));

        Assert.Equal(XLDataType.Number, cell.DataType);
        Assert.Equal(expectedFormat, cell.Style.NumberFormat.Format);
    }

    [Fact]
    public void Writes_dates_times_and_booleans_with_their_native_types()
    {
        var report = SingleSheet(
            new ColumnDefinition { Property = "Born", DataType = ExcelDataType.Date },
            new ColumnDefinition { Property = "Updated", DataType = ExcelDataType.DateTime, Format = "dd/mm/yyyy hh:mm" },
            new ColumnDefinition { Property = "Start", DataType = ExcelDataType.Time },
            new ColumnDefinition { Property = "Active", DataType = ExcelDataType.Boolean });

        var sheet = Open(Generate(report, SamplePerson())).Worksheet("Data");

        Assert.Equal(new DateTime(1990, 5, 17), sheet.Cell(2, 1).GetDateTime());
        Assert.Equal("yyyy-mm-dd", sheet.Cell(2, 1).Style.NumberFormat.Format);
        Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5), sheet.Cell(2, 2).GetDateTime());
        Assert.Equal("dd/mm/yyyy hh:mm", sheet.Cell(2, 2).Style.NumberFormat.Format);
        Assert.Equal(new TimeSpan(9, 30, 0), sheet.Cell(2, 3).GetTimeSpan());
        Assert.True(sheet.Cell(2, 4).GetBoolean());
    }

    [Fact]
    public void Text_keeps_leading_zeros_and_numbers_in_text_are_parsed_for_number_columns()
    {
        var report = SingleSheet(
            new ColumnDefinition { Property = "Code", DataType = ExcelDataType.Text },
            new ColumnDefinition { Property = "Amount", DataType = ExcelDataType.Number });

        var sheet = Open(Generate(report, SamplePerson())).Worksheet("Data");

        Assert.Equal("007", sheet.Cell(2, 1).GetText());
        Assert.Equal(1234.56, sheet.Cell(2, 2).GetDouble(), 3);
    }

    [Fact]
    public void Auto_type_follows_the_value()
    {
        var report = SingleSheet(
            new ColumnDefinition { Property = "Id" },
            new ColumnDefinition { Property = "Born" },
            new ColumnDefinition { Property = "Active" },
            new ColumnDefinition { Property = "Level" });

        var sheet = Open(Generate(report, SamplePerson())).Worksheet("Data");

        Assert.Equal(XLDataType.Number, sheet.Cell(2, 1).DataType);
        Assert.Equal(XLDataType.DateTime, sheet.Cell(2, 2).DataType);
        Assert.Equal(XLDataType.Boolean, sheet.Cell(2, 3).DataType);
        Assert.Equal("Senior", sheet.Cell(2, 4).GetText());
    }

    [Fact]
    public void Value_that_does_not_fit_the_type_is_written_as_text_instead_of_failing()
    {
        var person = new Person { Amount = "not a number" };
        var cell = DataCell(Generate(SingleSheet(new ColumnDefinition { Property = "Amount", DataType = ExcelDataType.Number }), person));

        Assert.Equal("not a number", cell.GetText());
    }

    [Fact]
    public void Value_mappings_and_default_values_are_applied()
    {
        var mappings = new Dictionary<string, string> { ["true"] = "Yes", ["false"] = "No" };
        var statusMappings = new Dictionary<string, string> { ["A"] = "Active", [""] = "Not set" };

        var report = SingleSheet(
            new ColumnDefinition { Property = "Active", DataType = ExcelDataType.Text, ValueMappings = mappings },
            new ColumnDefinition { Property = "Status", ValueMappings = statusMappings },
            new ColumnDefinition { Property = "Updated", DataType = ExcelDataType.DateTime, DefaultValue = "Never" });

        var people = new[] { SamplePerson(), new Person { Active = false, Status = "", Updated = null } };
        var sheet = Open(Generate(report, people)).Worksheet("Data");

        Assert.Equal("Yes", sheet.Cell(2, 1).GetText());
        Assert.Equal("Active", sheet.Cell(2, 2).GetText());
        Assert.Equal("No", sheet.Cell(3, 1).GetText());
        Assert.Equal("Not set", sheet.Cell(3, 2).GetText());
        Assert.Equal("Never", sheet.Cell(3, 3).GetText());
    }

    [Fact]
    public void Line_breaks_and_separators_become_new_lines_with_wrap_text()
    {
        var report = SingleSheet(
            new ColumnDefinition { Header = "Notes\nColumn", Property = "Notes" },
            new ColumnDefinition { Property = "Tags" },
            new ColumnDefinition { Property = "Name" });

        var options = new ExcelReportingOptions();
        options.Defaults.LineBreaks.Separators.Add("|");

        var bytes = new ExcelReportGenerator(options).Generate(report, ExcelReportData.ForAllSheets([SamplePerson()]));
        var sheet = Open(bytes).Worksheet("Data");

        Assert.Equal("Notes\nColumn", sheet.Cell(1, 1).GetText());
        Assert.Equal("line1\nline2\nline3", sheet.Cell(2, 1).GetText());
        Assert.True(sheet.Cell(2, 1).Style.Alignment.WrapText);
        Assert.Equal("one\ntwo", sheet.Cell(2, 2).GetText());
        Assert.True(sheet.Cell(2, 2).Style.Alignment.WrapText);
        Assert.False(sheet.Cell(2, 3).Style.Alignment.WrapText);
    }

    [Fact]
    public void Column_line_break_settings_override_defaults()
    {
        var column = new ColumnDefinition
        {
            Property = "Notes",
            LineBreaks = new LineBreakOptions { Separators = [";"], CollectionSeparator = ", " },
        };

        var person = new Person { Notes = "a;b|c" };
        var cell = DataCell(Generate(SingleSheet(column), person));

        Assert.Equal("a\nb|c", cell.GetText());
    }

    [Fact]
    public void Hyperlink_columns_create_clickable_links()
    {
        var report = SingleSheet(
            new ColumnDefinition { Property = "Website", DataType = ExcelDataType.Hyperlink },
            new ColumnDefinition { Property = "Name", DataType = ExcelDataType.Hyperlink });

        var sheet = Open(Generate(report, new Person { Website = "https://example.com", Name = "a@b.com" })).Worksheet("Data");

        Assert.Equal("https://example.com/", sheet.Cell(2, 1).GetHyperlink().ExternalAddress.ToString());
        Assert.Equal("a@b.com", sheet.Cell(2, 2).GetText());
        Assert.Equal("mailto:a@b.com", sheet.Cell(2, 2).GetHyperlink().ExternalAddress.ToString());
    }

    [Fact]
    public void Creates_one_sheet_per_definition_each_with_its_own_data_source()
    {
        var report = new ReportDefinition
        {
            Sheets =
            [
                new SheetDefinition { Name = "People", Title = "All people", Columns = [new() { Property = "Name" }] },
                new SheetDefinition { Name = "Cities", DataSource = "CityList", Columns = [new() { Property = "City" }] },
            ],
        };

        var data = new ExcelReportData()
            .Add("people", new[] { SamplePerson() })
            .Add("CityList", new[] { new Address { City = "Paris" }, new Address { City = "Rome" } });

        using var workbook = Open(new ExcelReportGenerator(new ExcelReportingOptions()).Generate(report, data));

        Assert.Equal(["People", "Cities"], workbook.Worksheets.Select(w => w.Name));

        var people = workbook.Worksheet("People");
        Assert.Equal("All people", people.Cell(1, 1).GetText());
        Assert.Equal("Name", people.Cell(2, 1).GetText());
        Assert.Equal("Ann", people.Cell(3, 1).GetText());

        var cities = workbook.Worksheet("Cities");
        Assert.Equal("Rome", cities.Cell(3, 1).GetText());
    }

    [Fact]
    public void Missing_data_source_gives_a_clear_error()
    {
        var generator = new ExcelReportGenerator(new ExcelReportingOptions());
        var ex = Assert.Throws<InvalidOperationException>(() =>
            generator.Generate(SingleSheet(new ColumnDefinition { Property = "Name" }), new ExcelReportData().Add("Other", new[] { 1 })));

        Assert.Contains("'Data'", ex.Message);
    }

    [Fact]
    public void Rows_over_the_limit_continue_on_extra_sheets()
    {
        var options = new ExcelReportingOptions();
        options.Defaults.MaxRowsPerSheet = 2;

        var longName = new string('X', 31);
        var report = new ReportDefinition
        {
            Sheets = [new SheetDefinition { Name = longName, Columns = [new() { Property = "Id" }] }],
        };

        var people = Enumerable.Range(1, 5).Select(i => SamplePerson(i));
        using var workbook = Open(new ExcelReportGenerator(options).Generate(report, ExcelReportData.ForAllSheets(people)));

        Assert.Equal(3, workbook.Worksheets.Count);
        Assert.Equal(new string('X', 27) + " (2)", workbook.Worksheet(2).Name);
        Assert.Equal(5, workbook.Worksheet(3).Cell(2, 1).GetDouble());
        Assert.True(workbook.Worksheet(3).Cell(3, 1).IsEmpty());
    }

    [Fact]
    public void Empty_list_gives_a_sheet_with_only_the_header()
    {
        using var workbook = Open(Generate(SingleSheet(new ColumnDefinition { Property = "Name" })));

        Assert.Equal("Name", workbook.Worksheet("Data").Cell(1, 1).GetText());
        Assert.True(workbook.Worksheet("Data").Cell(2, 1).IsEmpty());
    }

    [Fact]
    public void Works_with_dictionaries_for_dynamic_data()
    {
        var row = new Dictionary<string, object?> { ["Name"] = "Dyn", ["Score"] = 9.5 };
        var report = SingleSheet(new ColumnDefinition { Property = "name" }, new ColumnDefinition { Property = "Score" });

        var sheet = Open(Generate(report, row)).Worksheet("Data");

        Assert.Equal("Dyn", sheet.Cell(2, 1).GetText());
        Assert.Equal(9.5, sheet.Cell(2, 2).GetDouble());
    }

    [Fact]
    public void Unknown_property_gives_a_clear_error()
    {
        var ex = Assert.Throws<InvalidOperationException>(() =>
            Generate(SingleSheet(new ColumnDefinition { Header = "Oops", Property = "DoesNotExist" }), SamplePerson()));

        Assert.Contains("DoesNotExist", ex.Message);
        Assert.Contains("Oops", ex.Message);
    }

    [Fact]
    public void Invalid_report_definition_is_rejected()
    {
        var report = new ReportDefinition
        {
            Sheets =
            [
                new SheetDefinition { Name = "Bad/Name", Columns = [new() { Property = "Id" }] },
                new SheetDefinition { Name = "Same", Columns = [new() { Property = "Id" }] },
                new SheetDefinition { Name = "same", Columns = [] },
            ],
        };

        var errors = ExcelReportingOptionsValidator.ValidateReport("R", report);

        Assert.Contains(errors, e => e.Contains("must not contain"));
        Assert.Contains(errors, e => e.Contains("unique"));
        Assert.Contains(errors, e => e.Contains("at least one column"));
        Assert.Throws<InvalidOperationException>(() => Generate(report));
    }
}
