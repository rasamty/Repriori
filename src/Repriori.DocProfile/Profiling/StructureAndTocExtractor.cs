using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "structure" section: the document's heading outline (H1-H4), whether
/// it has a Table of Contents field, and whether it contains a handful of headings
/// SaMD documents conventionally need (Purpose, Scope, References, Revision
/// History). This is the closest thing to "does this document look like a
/// properly structured controlled document" that can be checked mechanically.
/// </summary>
internal static class StructureAndTocExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var section = JsonHelpers.Obj(root, "structure");
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return;

        var paragraphs = body.Elements<Paragraph>().ToList();
        JsonHelpers.Set(section, "paragraphCount", paragraphs.Count);
        JsonHelpers.Set(section, "wordCountEstimate",
            body.Descendants<Text>().SelectMany(t => t.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).Count());

        var heading1 = new JsonArray();
        var heading2 = new JsonArray();
        var heading3 = new JsonArray();
        var heading4 = new JsonArray();
        var sectionOrder = new JsonArray();

        foreach (var paragraph in paragraphs)
        {
            var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";
            var text = JsonHelpers.InnerText(paragraph).Trim();
            if (string.IsNullOrWhiteSpace(text)) continue;

            // H1 doubles as the document's top-level section order — a reasonable
            // proxy for "table of contents at a glance" even where no real TOC field exists.
            if (JsonHelpers.IsStyle(style, "Heading1", "Heading 1")) { heading1.Add(text); sectionOrder.Add(text); }
            else if (JsonHelpers.IsStyle(style, "Heading2", "Heading 2")) heading2.Add(text);
            else if (JsonHelpers.IsStyle(style, "Heading3", "Heading 3")) heading3.Add(text);
            else if (JsonHelpers.IsStyle(style, "Heading4", "Heading 4")) heading4.Add(text);
        }

        section["heading1Texts"] = heading1;
        section["heading2Texts"] = heading2;
        section["heading3Texts"] = heading3;
        section["heading4Texts"] = heading4;
        section["sectionOrder"] = sectionOrder;

        // A Table of Contents in Word is a "field" whose instruction text (e.g.
        // \o "1-3" or \t "Heading Style,1") controls how it was built — reading
        // that instruction is the only way to know *how* the TOC was generated,
        // as opposed to merely that one exists.
        var tocFieldInstructions = body.Descendants<FieldCode>()
            .Select(f => f.Text)
            .Where(t => t != null && t.Contains("TOC", StringComparison.OrdinalIgnoreCase))
            .ToList();
        JsonHelpers.Set(section, "hasTableOfContentsField", tocFieldInstructions.Count > 0);
        JsonHelpers.Set(section, "tocInstruction", string.IsNullOrWhiteSpace(tocFieldInstructions.FirstOrDefault()) ? null : tocFieldInstructions.First()!.Trim());
        JsonHelpers.Set(section, "tocTypeHint", TocHint(tocFieldInstructions.FirstOrDefault()));
        JsonHelpers.Set(section, "tocStyleName",
            body.Elements<Paragraph>()
                .Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value)
                .FirstOrDefault(styleId => styleId != null && styleId.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)));

        JsonHelpers.Set(section, "hasRevisionHistoryHeading", JsonHelpers.ContainsHeading(sectionOrder, "revision"));
        JsonHelpers.Set(section, "hasPurposeHeading", JsonHelpers.ContainsHeading(sectionOrder, "purpose"));
        JsonHelpers.Set(section, "hasIntroductionHeading", JsonHelpers.ContainsHeading(sectionOrder, "intro"));
        JsonHelpers.Set(section, "hasScopeHeading", JsonHelpers.ContainsHeading(sectionOrder, "scope"));
        JsonHelpers.Set(section, "hasReferencesHeading", JsonHelpers.ContainsHeading(sectionOrder, "reference"));
    }

    private static string? TocHint(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction)) return null;
        if (instruction.Contains("\\o", StringComparison.OrdinalIgnoreCase)) return "outline-levels";
        if (instruction.Contains("\\t", StringComparison.OrdinalIgnoreCase)) return "style-map";
        if (instruction.Contains("\\f", StringComparison.OrdinalIgnoreCase)) return "tc-fields";
        return "toc-field";
    }
}
