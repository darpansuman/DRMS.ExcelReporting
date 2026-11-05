using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace DRMS.ExcelReporting;

/// <summary>
/// A report file saved on disk. Gives you everything needed to upload it somewhere
/// (path, name, size, content type, a read stream) and deletes it when you are done.
/// </summary>
/// <remarks>
/// Use it with "using" so the local file is always cleaned up, even when the upload fails:
/// <code>
/// await using var report = generator.GenerateToFile("EmployeeReport", data);
///
/// await using (var stream = report.OpenRead())
/// {
///     await storage.UploadAsync(report.FileName, stream, report.Length, report.ContentType);
/// }
/// // The file (and any empty folder created for it) is deleted here.
/// </code>
/// To keep the file, set <see cref="DeleteOnDispose"/> to false.
/// </remarks>
public sealed class ExcelReportFile : IDisposable, IAsyncDisposable
{
    private readonly IReadOnlyList<string> _createdDirectories;
    private readonly bool _deleteEmptyDirectories;
    private readonly ILogger _logger;
    private int _deleted;

    internal ExcelReportFile(
        string reportName,
        string filePath,
        DateTimeOffset createdAt,
        IReadOnlyList<string> createdDirectories,
        bool deleteOnDispose,
        bool deleteEmptyDirectories,
        ILogger? logger)
    {
        ReportName = reportName;
        FilePath = filePath;
        CreatedAt = createdAt;
        Length = new FileInfo(filePath).Length;
        DeleteOnDispose = deleteOnDispose;
        _createdDirectories = createdDirectories;
        _deleteEmptyDirectories = deleteEmptyDirectories;
        _logger = logger ?? NullLogger.Instance;
    }

    /// <summary>The report key from configuration, for example "EmployeeReport".</summary>
    public string ReportName { get; }

    /// <summary>Full path of the file on disk.</summary>
    public string FilePath { get; }

    /// <summary>File name only, for example "EmployeeReport_20261105_101500.xlsx". Handy as an object name in storage.</summary>
    public string FileName => Path.GetFileName(FilePath);

    /// <summary>The folder that holds the file.</summary>
    public string DirectoryPath => Path.GetDirectoryName(FilePath)!;

    /// <summary>File size in bytes. Most storage APIs need this when uploading a stream.</summary>
    public long Length { get; }

    /// <summary>The MIME type of an .xlsx file.</summary>
    public string ContentType => ExcelReportGenerator.ContentType;

    /// <summary>When the file was created (local time).</summary>
    public DateTimeOffset CreatedAt { get; }

    /// <summary>
    /// When true (the default from "Output:DeleteFileOnDispose"), the file is deleted when this object is disposed.
    /// Set to false to keep the file.
    /// </summary>
    public bool DeleteOnDispose { get; set; }

    /// <summary>True after the file was deleted.</summary>
    public bool IsDeleted => Volatile.Read(ref _deleted) == 1;

    /// <summary>
    /// Opens the file for reading, ready to pass to an upload API.
    /// The caller must dispose the stream. The file can still be deleted while this stream is open.
    /// </summary>
    public FileStream OpenRead()
    {
        ObjectDisposedException.ThrowIf(IsDeleted, this);

        return new FileStream(
            FilePath,
            new FileStreamOptions
            {
                Mode = FileMode.Open,
                Access = FileAccess.Read,
                Share = FileShare.Read | FileShare.Delete,
                Options = FileOptions.Asynchronous | FileOptions.SequentialScan,
            });
    }

    /// <summary>
    /// Deletes the file now, plus the folders that were created for it if they are now empty
    /// (when "Output:DeleteEmptyDirectories" is true). Folders that already existed are never deleted.
    /// Safe to call more than once. Returns true when the file was deleted by this call.
    /// </summary>
    public bool Delete()
    {
        if (Interlocked.Exchange(ref _deleted, 1) == 1)
        {
            return false;
        }

        try
        {
            File.Delete(FilePath);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Volatile.Write(ref _deleted, 0); // Allow a later retry.
            _logger.LogWarning(ex, "Could not delete report file '{FilePath}'.", FilePath);
            return false;
        }

        if (_deleteEmptyDirectories)
        {
            DeleteEmptyCreatedDirectories(_createdDirectories, _logger);
        }

        _logger.LogDebug("Deleted report file '{FilePath}'.", FilePath);
        return true;
    }

    /// <summary>Deletes the file when <see cref="DeleteOnDispose"/> is true.</summary>
    public void Dispose()
    {
        if (DeleteOnDispose)
        {
            Delete();
        }
    }

    /// <summary>Same as <see cref="Dispose"/>; lets you write "await using".</summary>
    public ValueTask DisposeAsync()
    {
        Dispose();
        return ValueTask.CompletedTask;
    }

    /// <inheritdoc />
    public override string ToString() => FilePath;

    /// <summary>
    /// Deletes the given folders (deepest first) when they are empty. Errors are logged, not thrown,
    /// because another report may be writing to the same folder at the same time.
    /// </summary>
    internal static void DeleteEmptyCreatedDirectories(IReadOnlyList<string> directoriesDeepestFirst, ILogger logger)
    {
        foreach (var directory in directoriesDeepestFirst)
        {
            try
            {
                if (!Directory.Exists(directory) || Directory.EnumerateFileSystemEntries(directory).Any())
                {
                    return; // Not empty: the parent folders are not empty either.
                }

                Directory.Delete(directory, recursive: false);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                logger.LogDebug(ex, "Report folder '{Directory}' was not deleted.", directory);
                return;
            }
        }
    }
}
