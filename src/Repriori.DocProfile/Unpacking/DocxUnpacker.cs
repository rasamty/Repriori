using System.IO.Compression;
using Repriori.DocProfile.Validation;

namespace Repriori.DocProfile.Unpacking;

/// <summary>Where an unpacked .docx ended up, and every file that came out of it.</summary>
public sealed record UnpackResult(string OutputDirectory, IReadOnlyList<string> ExtractedFiles);

/// <summary>
/// Extracts a .docx file's real internal contents to a folder you can open and
/// browse in a normal file explorer or text editor.
///
/// The thing worth actually understanding here: a .docx is not a special binary
/// format at all. It is an ordinary ZIP archive, containing a folder structure of
/// plain XML files — <c>word/document.xml</c> holds the visible text, <c>word/styles.xml</c>
/// holds every named style, <c>docProps/core.xml</c> holds the author/title metadata, and
/// so on. Every extractor in the <c>Profiling</c> folder is really just a program that reads
/// specific pieces of that same XML — this class exists purely so you can see that
/// XML directly yourself, with nothing pre-interpreted, rather than reading a summary
/// of it.
/// </summary>
public static class DocxUnpacker
{
    /// <summary>
    /// Unpacks <paramref name="docxPath"/> into <paramref name="outputDirectory"/>.
    /// The output directory must not already contain files — this deliberately
    /// refuses to guess whether it is safe to merge into or overwrite an existing
    /// folder, rather than silently deleting or mixing in someone else's files.
    /// </summary>
    public static UnpackResult Unpack(string docxPath, string outputDirectory)
    {
        var readable = FileValidation.EnsureFileIsReadable(docxPath);
        if (!readable.IsOk)
            throw new IOException(readable.Reason);

        if (Directory.Exists(outputDirectory) && Directory.EnumerateFileSystemEntries(outputDirectory).Any())
            throw new IOException($"Output folder is not empty: {outputDirectory}. Choose an empty or new folder so nothing gets overwritten.");

        var writable = FileValidation.EnsureDirectoryIsWritable(outputDirectory);
        if (!writable.IsOk)
            throw new IOException(writable.Reason);

        // This one line is the entire "unpacking" — .NET's own ZipFile class
        // already knows how to walk a ZIP archive's internal folder structure and
        // recreate it on disk. There is no OOXML-specific step here at all.
        //
        // The checks above are a fast, clear first line of defence, but they run
        // *before* this line, not *during* it — a file can still be locked here a
        // moment later (this is the exact race that showed up testing
        // WordProfileExtractor.FromDocx). Catching the real failure and
        // translating it is what actually closes that gap.
        try
        {
            ZipFile.ExtractToDirectory(docxPath, outputDirectory);
        }
        catch (IOException ex)
        {
            throw new IOException($"File is in use by another program (is it open in Word?): {docxPath} ({ex.Message})");
        }

        var extractedFiles = Directory.EnumerateFiles(outputDirectory, "*", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(outputDirectory, path))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return new UnpackResult(outputDirectory, extractedFiles);
    }
}
