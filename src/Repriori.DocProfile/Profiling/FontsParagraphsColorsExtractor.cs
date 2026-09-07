using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "fonts", "colors", and "paragraphs" sections. This must run *after*
/// <see cref="StylesExtractor"/>, because the "fonts" section is largely a
/// convenience copy of what the styles section already found (see
/// <see cref="JsonHelpers.CopyFont"/>) — this extractor's own job is the parts
/// styles alone cannot answer: every distinct typeface and colour actually
/// present anywhere in the body, regardless of which style (if any) put it there.
/// </summary>
internal static class FontsParagraphsColorsExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var fonts = JsonHelpers.Obj(root, "fonts");
        var colors = JsonHelpers.Obj(root, "colors");
        var paragraphs = JsonHelpers.Obj(root, "paragraphs");
        var styles = JsonHelpers.Obj(root, "styles");
        var body = word.MainDocumentPart?.Document?.Body;

        // SortedSet so the output is alphabetical and stable — otherwise the same
        // document could produce a differently-ordered list on every run, which
        // would make a diff between two profiles noisy for no reason.
        var typefaces = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        var nonBwColors = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);

        if (body is not null)
        {
            foreach (var runFonts in body.Descendants<RunFonts>())
            {
                if (!string.IsNullOrWhiteSpace(runFonts.Ascii?.Value)) typefaces.Add(runFonts.Ascii.Value);
                if (!string.IsNullOrWhiteSpace(runFonts.HighAnsi?.Value)) typefaces.Add(runFonts.HighAnsi.Value);
            }

            foreach (var color in body.Descendants<Color>())
            {
                var hex = color.Val?.Value;
                if (JsonHelpers.IsNonBw(hex)) nonBwColors.Add(hex!.ToUpperInvariant());
            }
        }

        fonts["typefacesUsed"] = JsonHelpers.ToArray(typefaces);
        colors["nonBlackWhiteHexUsed"] = JsonHelpers.ToArray(nonBwColors);
        JsonHelpers.Set(colors, "bodyUsesOnlyBlackOrWhite", nonBwColors.Count == 0);

        JsonHelpers.CopyFont(styles, "normal", fonts, "bodyTypeface", "bodySizePt", "bodyColorHex");
        JsonHelpers.CopyFont(styles, "heading1", fonts, "heading1Typeface", "heading1SizePt", "heading1ColorHex");
        JsonHelpers.CopyFont(styles, "heading2", fonts, "heading2Typeface", "heading2SizePt", "heading2ColorHex");
        JsonHelpers.CopyFont(styles, "heading3", fonts, "heading3Typeface", "heading3SizePt", "heading3ColorHex");
        JsonHelpers.CopyFont(styles, "heading4", fonts, "heading4Typeface", "heading4SizePt", "heading4ColorHex");

        // "What does normal body text look like" is answered by the *first*
        // paragraph that uses the Normal (or BodyText) style — good enough as a
        // representative sample without inspecting every paragraph in the file.
        var normalParagraph = body?.Elements<Paragraph>()
            .FirstOrDefault(p => (p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal") is "Normal" or "BodyText");

        JsonHelpers.Set(paragraphs, "bodyAlignment", normalParagraph?.ParagraphProperties?.Justification?.Val?.Value.ToString()?.ToLowerInvariant());
        var spacing = normalParagraph?.ParagraphProperties?.SpacingBetweenLines;
        JsonHelpers.Set(paragraphs, "spaceAfterPt", JsonHelpers.TwipToPt(spacing?.After));
        JsonHelpers.Set(paragraphs, "spaceBeforePt", JsonHelpers.TwipToPt(spacing?.Before));
        JsonHelpers.Set(paragraphs, "bodyLineSpacing", JsonHelpers.LineSpacingName(spacing));
    }
}
