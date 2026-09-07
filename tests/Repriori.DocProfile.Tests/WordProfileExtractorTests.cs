using System.Text.Json;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests;

public class WordProfileExtractorTests
{
    [Fact]
    public void FromDocx_stream_sets_the_four_always_present_fields()
    {
        using var docx = TestDocxBuilder.Minimal();

        var json = WordProfileExtractor.FromDocx(docx, "example.docx");
        var profile = JsonDocument.Parse(json).RootElement;

        Assert.Equal("example-extracted", profile.GetProperty("profileId").GetString());
        Assert.Equal("example.docx", profile.GetProperty("sourceFile").GetString());
        Assert.Equal("1.1.0", profile.GetProperty("extractorVersion").GetString());
        Assert.False(string.IsNullOrWhiteSpace(profile.GetProperty("extractedAt").GetString()));
    }

    [Fact]
    public void FromDocx_produces_well_formed_json_matching_the_schema_shape()
    {
        using var docx = TestDocxBuilder.WithHeadings();

        var json = WordProfileExtractor.FromDocx(docx, "headings.docx");
        var profile = JsonDocument.Parse(json).RootElement; // throws if not valid JSON

        // Spot-check that every top-level section from WordFormatComparer.json
        // actually exists in the output — a missing section would mean an
        // extractor silently didn't run, not merely that its answer was null.
        foreach (var section in new[] { "document", "revisions", "comments", "page", "headersFooters", "logo", "structure", "styles", "fonts", "colors", "emphasis", "paragraphs", "tables", "figures", "lists" })
        {
            Assert.True(profile.TryGetProperty(section, out _), $"expected top-level section '{section}' to exist");
        }
    }

    [Fact]
    public void FromDocx_path_throws_IOException_with_a_clear_message_when_file_is_missing()
    {
        var ex = Assert.Throws<IOException>(() => WordProfileExtractor.FromDocx(@"C:\does\not\exist.docx"));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FromDocx_path_throws_a_clear_message_when_the_containing_folder_is_missing()
    {
        // Regression test: DirectoryNotFoundException must not be reported as
        // "file is in use by another program" — see the comment in
        // WordProfileExtractor.FromDocx explaining why this needs its own catch.
        var ex = Assert.Throws<IOException>(() => WordProfileExtractor.FromDocx(@"C:\this\folder\does\not\exist\file.docx"));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("in use", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FromDocx_path_gives_a_clear_message_when_the_file_is_locked_by_another_process()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"locked-{Guid.NewGuid():N}.docx");
        using var docx = TestDocxBuilder.Minimal();
        File.WriteAllBytes(tempFile, docx.ToArray());

        try
        {
            // Opening with FileShare.None reproduces exactly what Word does while
            // a document is open — no other process, including this library,
            // should be able to read it either.
            using var exclusiveLock = new FileStream(tempFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

            var ex = Assert.Throws<IOException>(() => WordProfileExtractor.FromDocx(tempFile));
            Assert.Contains("in use", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void FromDocx_throws_a_recognisable_error_for_a_file_that_is_not_really_a_docx()
    {
        using var notADocx = TestDocxBuilder.NotAZipFile();

        // The OpenXml SDK's own FileFormatException is already a clear, specific
        // signal ("not a valid package") — this test exists to prove the library
        // fails loudly and predictably here rather than hanging, crashing with a
        // stack overflow, or silently returning an empty/wrong profile.
        Assert.ThrowsAny<Exception>(() => WordProfileExtractor.FromDocx(notADocx, "fake.docx"));
    }
}
