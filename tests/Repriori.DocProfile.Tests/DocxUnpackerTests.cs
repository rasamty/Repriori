using Repriori.DocProfile.Tests.Fixtures;
using Repriori.DocProfile.Unpacking;
using Xunit;

namespace Repriori.DocProfile.Tests;

public class DocxUnpackerTests
{
    [Fact]
    public void Unpack_extracts_the_real_internal_parts_of_a_docx()
    {
        var docxPath = WriteTempDocx(TestDocxBuilder.WithTable());
        var outputDir = Path.Combine(Path.GetTempPath(), $"unpack-{Guid.NewGuid():N}");

        try
        {
            var result = DocxUnpacker.Unpack(docxPath, outputDir);

            // These are the actual, real file names inside every .docx — not
            // anything specific to this fixture. Their presence is what proves
            // this genuinely unpacked a ZIP rather than doing something bespoke.
            Assert.Contains("word/document.xml", result.ExtractedFiles.Select(NormalizeSlashes));
            Assert.Contains("[Content_Types].xml", result.ExtractedFiles.Select(NormalizeSlashes));
            Assert.True(File.Exists(Path.Combine(outputDir, "word", "document.xml")));
        }
        finally
        {
            File.Delete(docxPath);
            if (Directory.Exists(outputDir)) Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void Unpack_refuses_to_write_into_a_non_empty_folder()
    {
        var docxPath = WriteTempDocx(TestDocxBuilder.Minimal());
        var outputDir = Path.Combine(Path.GetTempPath(), $"unpack-nonempty-{Guid.NewGuid():N}");
        Directory.CreateDirectory(outputDir);
        File.WriteAllText(Path.Combine(outputDir, "already-here.txt"), "hi");

        try
        {
            var ex = Assert.Throws<IOException>(() => DocxUnpacker.Unpack(docxPath, outputDir));
            Assert.Contains("not empty", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(docxPath);
            Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void Unpack_throws_a_clear_error_when_the_input_file_does_not_exist()
    {
        var outputDir = Path.Combine(Path.GetTempPath(), $"unpack-missing-{Guid.NewGuid():N}");

        var ex = Assert.Throws<IOException>(() => DocxUnpacker.Unpack(@"C:\does\not\exist.docx", outputDir));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static string WriteTempDocx(MemoryStream content)
    {
        using (content)
        {
            var path = Path.Combine(Path.GetTempPath(), $"unpack-src-{Guid.NewGuid():N}.docx");
            File.WriteAllBytes(path, content.ToArray());
            return path;
        }
    }

    private static string NormalizeSlashes(string path) => path.Replace('\\', '/');
}
