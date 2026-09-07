using Repriori.DocProfile.Validation;
using Xunit;

namespace Repriori.DocProfile.Tests.Validation;

public class FileValidationTests
{
    [Fact]
    public void EnsureFileIsReadable_fails_clearly_for_a_missing_file()
    {
        var result = FileValidation.EnsureFileIsReadable(@"C:\does\not\exist.docx");

        Assert.False(result.IsOk);
        Assert.Contains("not found", result.Reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EnsureFileIsReadable_succeeds_for_a_real_readable_file()
    {
        var path = Path.Combine(Path.GetTempPath(), $"readable-{Guid.NewGuid():N}.txt");
        File.WriteAllText(path, "hello");

        try
        {
            var result = FileValidation.EnsureFileIsReadable(path);
            Assert.True(result.IsOk);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void EnsureFileIsReadable_fails_clearly_when_the_file_is_exclusively_locked()
    {
        var path = Path.Combine(Path.GetTempPath(), $"locked-{Guid.NewGuid():N}.txt");
        File.WriteAllText(path, "hello");

        try
        {
            using var exclusiveLock = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

            var result = FileValidation.EnsureFileIsReadable(path);

            Assert.False(result.IsOk);
            Assert.Contains("in use", result.Reason, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void EnsureDirectoryIsWritable_creates_the_folder_if_it_does_not_exist_yet()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"writable-{Guid.NewGuid():N}");
        Assert.False(Directory.Exists(dir));

        try
        {
            var result = FileValidation.EnsureDirectoryIsWritable(dir);

            Assert.True(result.IsOk);
            Assert.True(Directory.Exists(dir));
        }
        finally
        {
            if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void EnsureDirectoryIsWritable_leaves_no_leftover_probe_file_behind()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"writable-clean-{Guid.NewGuid():N}");

        try
        {
            FileValidation.EnsureDirectoryIsWritable(dir);

            Assert.Empty(Directory.EnumerateFileSystemEntries(dir));
        }
        finally
        {
            if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true);
        }
    }
}
