using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;

namespace Repriori.DocProfile;

/// <summary>
/// The public entry point for turning a .docx file into a "profile" — a single
/// JSON document describing its structure, styles, fonts, tables, comments,
/// tracked changes, and more, matching the shape defined by the embedded
/// WordFormatComparer.json schema.
///
/// This class itself does almost no work. Its only job is: load the schema
/// template, open the file, run each of the twelve extractors in
/// <see cref="Profiling"/> in turn (each one fills in its own section of the
/// JSON and nothing else), and serialise the result. Reading the actual OOXML —
/// the interesting part — lives entirely in those extractor classes, which is
/// what makes each one independently readable and independently testable.
/// </summary>
public static class WordProfileExtractor
{
    /// <summary>Extracts a profile from a .docx file on disk.</summary>
    /// <exception cref="IOException">The file does not exist, or cannot currently be read (for example because it is open in Word).</exception>
    public static string FromDocx(string path)
    {
        // Deliberately not "check readability, then separately open" — a file that
        // was readable the moment it was checked can still be locked a millisecond
        // later (Word's autosave takes and releases its lock constantly while a
        // document is open). Opening it exactly once and translating whatever
        // failure actually happens is the only version of this that isn't racy.
        Stream stream;
        try
        {
            stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }
        catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException)
        {
            // Both are subclasses of IOException, so both must be caught here,
            // before the generic IOException catch below — otherwise "the folder
            // doesn't exist" and "the file is locked" collapse into the same wrong
            // message. This exact mistake is what the "missing file" smoke test
            // case below caught the first time this method was written.
            throw new IOException($"File not found: {path}");
        }
        catch (UnauthorizedAccessException)
        {
            throw new IOException($"Permission denied reading: {path}");
        }
        catch (IOException ex)
        {
            throw new IOException($"File is in use by another program (is it open in Word?): {path} ({ex.Message})");
        }

        using (stream)
        {
            return FromDocx(stream, Path.GetFileName(path));
        }
    }

    /// <summary>
    /// Extracts a profile from an already-open stream — used by <see cref="FromDocx(string)"/>
    /// above, and directly by tests, which build a .docx in memory rather than
    /// writing a temporary file to disk for every single test case.
    /// </summary>
    public static string FromDocx(Stream docxStream, string fileName)
    {
        using var word = WordprocessingDocument.Open(docxStream, isEditable: false);

        var root = LoadTemplate();
        JsonHelpers.Set(root, "profileId", Path.GetFileNameWithoutExtension(fileName) + "-extracted");
        JsonHelpers.Set(root, "sourceFile", fileName);
        JsonHelpers.Set(root, "extractedAt", DateTimeOffset.UtcNow.ToString("o"));
        JsonHelpers.Set(root, "extractorVersion", "1.1.0");

        // Order matters only once: FontsParagraphsColorsExtractor copies some of
        // its output from whatever StylesExtractor already wrote, so styles must
        // run first. Every other extractor is independent of the others and could
        // run in any order (or in parallel, though there is no need to bother —
        // profiling a single document takes milliseconds).
        DocumentPropertiesExtractor.Extract(word, root);
        RevisionsExtractor.Extract(word, root);
        CommentsExtractor.Extract(word, root);
        PageExtractor.Extract(word, root);
        HeadersFootersLogoExtractor.Extract(word, root);
        StructureAndTocExtractor.Extract(word, root);
        StylesExtractor.Extract(word, root);
        FontsParagraphsColorsExtractor.Extract(word, root);
        EmphasisExtractor.Extract(word, root);
        TablesExtractor.Extract(word, root);
        FiguresExtractor.Extract(word, root);
        ListsExtractor.Extract(word, root);

        return root.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    /// Loads WordFormatComparer.json out of this assembly's own embedded
    /// resources — see the &lt;EmbeddedResource&gt; entry in the .csproj. Embedding it
    /// means the schema always travels with the compiled library: there is no
    /// separate file a caller could forget to copy alongside the DLL, and no path
    /// to configure.
    /// </summary>
    private static JsonObject LoadTemplate()
    {
        var assembly = Assembly.GetExecutingAssembly();
        // An embedded resource's runtime name is "<DefaultNamespace>.<path-with-dots-instead-of-slashes>",
        // which for a file at the project root is just "<AssemblyName>.<FileName>".
        var resourceName = $"{assembly.GetName().Name}.WordFormatComparer.json";

        using var resourceStream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException(
                $"Embedded resource '{resourceName}' was not found. Check the <EmbeddedResource> entry in Repriori.DocProfile.csproj.");

        return JsonNode.Parse(resourceStream)!.AsObject();
    }
}
