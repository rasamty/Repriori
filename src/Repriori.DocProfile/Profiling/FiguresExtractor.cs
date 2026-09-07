using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "figures" section: how many images the document contains, how many
/// have a caption (a paragraph in the "Caption" style), and — separately, since
/// captions and accessibility alt text are unrelated features — how many images
/// have alt text set at all, which matters for accessibility compliance.
/// </summary>
internal static class FiguresExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var section = JsonHelpers.Obj(root, "figures");
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return;

        var drawings = body.Descendants<Drawing>().ToList();
        JsonHelpers.Set(section, "imageCount", drawings.Count);

        var captions = body.Elements<Paragraph>()
            .Where(p => JsonHelpers.IsStyle(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "", "Caption"))
            .Select(JsonHelpers.InnerText)
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .ToList();
        JsonHelpers.Set(section, "captionCount", captions.Count);
        section["captionTexts"] = JsonHelpers.ToArray(captions);

        var withAltText = 0;
        var missingAltText = 0;
        foreach (var drawing in drawings)
        {
            var description = drawing.Descendants<DW.DocProperties>().FirstOrDefault()?.Description?.Value
                ?? drawing.Descendants<A.NonVisualDrawingProperties>().FirstOrDefault()?.Description?.Value;
            if (string.IsNullOrWhiteSpace(description)) missingAltText++; else withAltText++;
        }

        // Null (not zero) when there are no images at all — "0 of 0 have alt text"
        // is a meaningless statistic, and a caller checking "is anything missing
        // alt text" should not have to special-case an empty document separately.
        JsonHelpers.Set(section, "imagesWithAltText", drawings.Count == 0 ? null : withAltText);
        JsonHelpers.Set(section, "imagesMissingAltText", drawings.Count == 0 ? null : missingAltText);
    }
}
