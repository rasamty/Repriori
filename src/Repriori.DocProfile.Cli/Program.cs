// Repriori.DocProfile.Cli — the console wrapper around Repriori.DocProfile.
//
// This file deliberately contains no rules of its own: no format-checking logic,
// no OOXML reading. Its only two jobs are (1) turn command-line arguments into a
// call to the library, and (2) turn whatever the library returns — or throws —
// into console output and an exit code. If this file is ever wrong about *what*
// a profile contains, the bug is in Repriori.DocProfile, not here.

using Repriori.DocProfile;
using Repriori.DocProfile.Comparing;
using Repriori.DocProfile.Unpacking;
using Repriori.DocProfile.Validation;

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
        "compare" => RunCompare(args[1..]),
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

int RunCompare(string[] rest)
{
    // golden and submission are the first two arguments that are not part of
    // an option — not necessarily rest[0] and rest[1], since --rules <path> and
    // --show-diffs could come first. This walks the list once, skipping an
    // option and (for --rules specifically) the value straight after it.
    var positional = new List<string>();
    for (var i = 0; i < rest.Length; i++)
    {
        if (rest[i] == "--show-diffs") continue;
        if (rest[i] == "--rules") { i++; continue; } // also skip the path that follows it
        positional.Add(rest[i]);
    }

    if (positional.Count < 2)
    {
        Console.Error.WriteLine("Usage: docprofile compare <golden> <submission> [--rules <rules.json>] [--show-diffs]");
        Console.Error.WriteLine("       <golden> and <submission> may each be a .docx file or an already-extracted profile .json file.");
        return 1;
    }

    var goldenPath = positional[0];
    var submissionPath = positional[1];
    var rulesPath = ReadOptionValue(rest, "--rules");
    var showDiffs = rest.Contains("--show-diffs");

    var ruleSet = rulesPath is null ? ComparisonRuleSet.Default() : ComparisonRuleSet.FromFile(rulesPath);
    var result = ProfileComparisonEngine.Compare(LoadProfileJson(goldenPath), LoadProfileJson(submissionPath), ruleSet);

    PrintComparisonReport(result, showDiffs);
    return result.Passed ? 0 : 1;
}

// A .docx is extracted on the fly (reusing Phase 3's extractor); anything else
// is read as an already-extracted profile .json. This lets compare take either
// two raw documents or two already-extracted profiles, without needing two
// separate commands for what is really one decision (extract first, or not).
string LoadProfileJson(string path)
{
    if (Path.GetExtension(path).Equals(".docx", StringComparison.OrdinalIgnoreCase))
    {
        return WordProfileExtractor.FromDocx(path);
    }

    var readable = FileValidation.EnsureFileIsReadable(path);
    if (!readable.IsOk) throw new IOException(readable.Reason);
    return File.ReadAllText(path);
}

void PrintComparisonReport(ComparisonResult result, bool showDiffs)
{
    Console.WriteLine(result.Passed ? "PASS" : "FAIL");
    Console.WriteLine();

    if (result.Violations.Count == 0)
    {
        Console.WriteLine("No rule violations.");
    }
    else
    {
        Console.WriteLine($"{result.Violations.Count} rule violation(s):");
        foreach (var violation in result.Violations)
        {
            Console.WriteLine($"  - {violation.Rule.Field}: {violation.Rule.Description}");
            Console.WriteLine($"      {violation.Detail}");
        }
    }

    Console.WriteLine();
    Console.WriteLine($"{result.Diffs.Count} field(s) differ between golden and submission (informational — not all of these are rule violations, and not all rule violations show up here; see 7.2 in the Milestone 1 document for why those are two separate lists).");
    if (showDiffs)
    {
        foreach (var diff in result.Diffs)
        {
            Console.WriteLine($"  {diff}");
        }
    }
    else if (result.Diffs.Count > 0)
    {
        Console.WriteLine("  (use --show-diffs to list them)");
    }
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
        docprofile — extract, unpack, or compare .docx files

        Usage:
          docprofile extract <file.docx> [-o <output.json>]   Extract a format profile as JSON.
                                                                Without -o, prints to the terminal.
          docprofile unpack <file.docx> -o <output-folder>     Extract the file's real internal XML
                                                                parts to a folder, so you can open and
                                                                read them yourself.
          docprofile compare <golden> <submission>             Compare two profiles (or two .docx
                             [--rules <rules.json>]             files, extracted on the fly) against
                             [--show-diffs]                     a rule set. Exit code 0 = pass, 1 = fail.
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
