using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "emphasis" section: what percentage of body text, heading text, and
/// table text is bold or italic. This is measured character-by-character from the
/// actual runs in the document — not from style names — because a run can be
/// manually bolded even when its paragraph style says otherwise (or vice versa),
/// and "how much of this document did someone manually emphasise" is a genuinely
/// different question from "what does the Normal style say."
/// </summary>
internal static class EmphasisExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var section = JsonHelpers.Obj(root, "emphasis");
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return;

        Measure(section, "body", body.Elements<Paragraph>().Where(p => !IsHeading(p) && p.Parent is Body));
        Measure(section, "headings", body.Elements<Paragraph>().Where(IsHeading));
        Measure(section, "tables", body.Descendants<Table>().SelectMany(t => t.Descendants<Paragraph>()));
    }

    private static void Measure(JsonObject emphasis, string key, IEnumerable<Paragraph> paragraphs)
    {
        var slot = JsonHelpers.Obj(emphasis, key);
        int totalChars = 0, boldChars = 0, italicChars = 0;

        foreach (var paragraph in paragraphs)
        {
            foreach (var run in paragraph.Elements<Run>())
            {
                var charCount = run.Descendants<Text>().Sum(t => t.Text.Length);
                if (charCount == 0) continue; // an empty or non-text run (e.g. just a line break) has nothing to weigh

                totalChars += charCount;
                if (run.RunProperties?.Bold is not null) boldChars += charCount;
                if (run.RunProperties?.Italic is not null) italicChars += charCount;
            }
        }

        JsonHelpers.Set(slot, "chars", totalChars);
        JsonHelpers.Set(slot, "boldPct", totalChars == 0 ? null : Math.Round(100.0 * boldChars / totalChars, 1));
        JsonHelpers.Set(slot, "italicPct", totalChars == 0 ? null : Math.Round(100.0 * italicChars / totalChars, 1));
    }

    private static bool IsHeading(Paragraph paragraph)
    {
        var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";
        return JsonHelpers.IsStyle(style, "Heading1", "Heading 1", "Heading2", "Heading 2", "Heading3", "Heading 3", "Heading4", "Heading 4");
    }
}
