// Repriori.DocProfile.Cli — the console wrapper around Repriori.DocProfile.
//
// This file deliberately contains no rules of its own: no format-checking logic,
// no OOXML reading. Its only two jobs are (1) turn command-line arguments into a
// call to the library, and (2) turn whatever the library returns — or throws —
// into console output and an exit code. If this file is ever wrong about *what*
// a profile contains, the bug is in Repriori.DocProfile, not here.

using Repriori.DocProfile;
using Repriori.DocProfile.Unpacking;

if (args.Length == 0)
{
    PrintUsage();
    return 1;
}

// A separate try/catch around each command (rather than one around all of Main)
// means a command that partly succeeds before failing — extract, for instance,
// writing nothing to disk until the JSON is fully built — behaves predictably:
// either the whole command worked, or nothing it was supposed to produce exists.
try
{
    return args[0] switch
    {
        "extract" => RunExtract(args[1..]),
        "unpack" => RunUnpack(args[1..]),
        "-h" or "--help" or "help" => RunHelp(),
        var unknown => UnknownCommand(unknown),
    };
}
catch (Exception ex)
{
    // Every exception the library throws on purpose (see FileValidation and the
    // try/catch blocks in WordProfileExtractor/DocxUnpacker) already has a clear,
    // one-sentence message meant to be read by a person — this is why the CLI
    // only ever needs to print ex.Message, never a stack trace, for the tool to
    // still be genuinely useful when something goes wrong.
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}

int RunExtract(string[] rest)
{
    if (rest.Length == 0)
    {
        Console.Error.WriteLine("Usage: docprofile extract <file.docx> [-o <output.json>]");
        return 1;
    }

    var inputPath = rest[0];
    var outputPath = ReadOptionValue(rest, "-o");

    var json = WordProfileExtractor.FromDocx(inputPath);

    if (outputPath is null)
    {
        // No -o given: print to stdout so the tool is still useful in a pipeline
        // (e.g. piped into another program, or just eyeballed in the terminal)
        // without forcing a file to be written every time.
        Console.WriteLine(json);
        return 0;
    }

    File.WriteAllText(outputPath, json);
    Console.WriteLine($"Wrote profile to {outputPath}");
    return 0;
}

int RunUnpack(string[] rest)
{
    if (rest.Length == 0 || ReadOptionValue(rest, "-o") is not { } outputDir)
    {
        Console.Error.WriteLine("Usage: docprofile unpack <file.docx> -o <output-folder>");
        return 1;
    }

    var inputPath = rest[0];
    var result = DocxUnpacker.Unpack(inputPath, outputDir);

    Console.WriteLine($"Unpacked {result.ExtractedFiles.Count} files to {result.OutputDirectory}");
    foreach (var extractedFile in result.ExtractedFiles)
    {
        Console.WriteLine($"  {extractedFile}");
    }

    return 0;
}

int RunHelp()
{
    PrintUsage();
    return 0;
}

int UnknownCommand(string command)
{
    Console.Error.WriteLine($"Unknown command: {command}");
    PrintUsage();
    return 1;
}

void PrintUsage()
{
    Console.WriteLine(
        """
        docprofile — extract or unpack a .docx file

        Usage:
          docprofile extract <file.docx> [-o <output.json>]   Extract a format profile as JSON.
                                                                Without -o, prints to the terminal.
          docprofile unpack <file.docx> -o <output-folder>     Extract the file's real internal XML
                                                                parts to a folder, so you can open and
                                                                read them yourself.
          docprofile --help                                    Show this message.
        """);
}

// Looks for a flag (e.g. "-o") in the argument list and returns the value that
// immediately follows it, or null if the flag wasn't given at all. Returning
// null (rather than throwing) for a missing flag lets each command decide for
// itself whether that flag was actually required.
static string? ReadOptionValue(string[] args, string optionName)
{
    var index = Array.IndexOf(args, optionName);
    return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
}
