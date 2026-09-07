using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "styles" section: the typeface, size, bold/italic, and colour defined
/// by each named style (Normal, Heading 1, Title, and so on). A "style" in Word is
/// a saved formatting preset — reading the style definitions answers "what is this
/// document's design meant to be," as distinct from FontsParagraphsColorsExtractor,
/// which measures what formatting is actually *used* run by run.
/// </summary>
internal static class StylesExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var section = JsonHelpers.Obj(root, "styles");
        var stylesPart = word.MainDocumentPart?.StyleDefinitionsPart?.Styles;

        var names = new JsonArray();
        if (stylesPart is not null)
        {
            foreach (var style in stylesPart.Elements<Style>())
            {
                var name = style.StyleName?.Val?.Value ?? style.StyleId?.Value;
                if (!string.IsNullOrWhiteSpace(name)) names.Add(name);
            }
        }
        section["styleNamesPresent"] = names;

        FillStyle(section, "normal", stylesPart, "Normal");
        FillStyle(section, "heading1", stylesPart, "Heading1", "Heading 1");
        FillStyle(section, "heading2", stylesPart, "Heading2", "Heading 2");
        FillStyle(section, "heading3", stylesPart, "Heading3", "Heading 3");
        FillStyle(section, "heading4", stylesPart, "Heading4", "Heading 4");
        FillStyle(section, "title", stylesPart, "Title");
        FillStyle(section, "caption", stylesPart, "Caption");
        FillStyle(section, "header", stylesPart, "Header");
        FillStyle(section, "footer", stylesPart, "Footer");
        FillStyle(section, "toc1", stylesPart, "TOC1", "toc 1");
        FillStyle(section, "toc2", stylesPart, "TOC2", "toc 2");
    }

    /// <summary>
    /// A style can be looked up either by its short internal id ("Heading1", no
    /// space) or its human-readable display name ("Heading 1", with a space) —
    /// different Word versions and templates are inconsistent about which one a
    /// paragraph actually references, so this checks both.
    /// </summary>
    private static void FillStyle(JsonObject stylesRoot, string key, Styles? styles, params string[] candidateIds)
    {
        var slot = stylesRoot[key] as JsonObject ?? new JsonObject();
        stylesRoot[key] = slot;

        var style = styles?.Elements<Style>().FirstOrDefault(s =>
            candidateIds.Any(id =>
                string.Equals(s.StyleId?.Value, id.Replace(" ", ""), StringComparison.OrdinalIgnoreCase)
                || string.Equals(s.StyleName?.Val?.Value, id, StringComparison.OrdinalIgnoreCase)));

        var runProperties = style?.StyleRunProperties;
        JsonHelpers.Set(slot, "typeface", runProperties?.RunFonts?.Ascii?.Value ?? runProperties?.RunFonts?.HighAnsi?.Value);
        JsonHelpers.Set(slot, "sizePt", JsonHelpers.HalfPointToPt(runProperties?.FontSize?.Val));
        // "style is null" (the style doesn't exist in this document at all) is a
        // genuinely different answer from "the style exists but isn't bold" — the
        // ternary keeps that distinction instead of collapsing both to false.
        JsonHelpers.Set(slot, "bold", style is null ? null : runProperties?.Bold is not null);
        JsonHelpers.Set(slot, "italic", style is null ? null : runProperties?.Italic is not null);
        JsonHelpers.Set(slot, "colorHex", runProperties?.Color?.Val?.Value);
    }
}
