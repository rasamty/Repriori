using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "page" section: paper size, orientation, margins, and how many
/// distinct "sections" the document has (a Word document can mix page sizes or
/// orientations by splitting itself into sections — this is why some pages in a
/// report can be landscape while the rest are portrait).
/// </summary>
internal static class PageExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var section = JsonHelpers.Obj(root, "page");
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return;

        var sectionProperties = body.Descendants<SectionProperties>().ToList();
        JsonHelpers.Set(section, "sectionCount", sectionProperties.Count);
        // Word does not store a page count anywhere in the file — it's a property
        // of how the content happens to lay out on a given printer/font, computed
        // only at render time. There is no reliable way to answer this from the
        // XML alone, so it is left null on purpose rather than guessed.
        JsonHelpers.Set(section, "estimatedPageCount", null);

        // The *last* section's properties describe the page the document currently
        // ends on — the most representative single answer when a caller just wants
        // "what size is this document," even though earlier sections could differ.
        var lastSection = sectionProperties.LastOrDefault();
        if (lastSection is null) return;

        var size = lastSection.GetFirstChild<PageSize>();
        if (size?.Width?.HasValue == true && size.Height?.HasValue == true)
        {
            var widthMm = JsonHelpers.TwipToMm((int)size.Width.Value)!.Value;
            var heightMm = JsonHelpers.TwipToMm((int)size.Height.Value)!.Value;
            JsonHelpers.Set(section, "widthMm", widthMm);
            JsonHelpers.Set(section, "heightMm", heightMm);
            JsonHelpers.Set(section, "orientation", widthMm > heightMm ? "landscape" : "portrait");
            JsonHelpers.Set(section, "paperHint", JsonHelpers.PaperHint(widthMm, heightMm));
        }

        var margins = lastSection.GetFirstChild<PageMargin>();
        var marginsSection = JsonHelpers.Obj(section, "marginsMm");
        if (margins is not null)
        {
            JsonHelpers.Set(marginsSection, "top", JsonHelpers.TwipToMm(margins.Top?.Value));
            JsonHelpers.Set(marginsSection, "bottom", JsonHelpers.TwipToMm(margins.Bottom?.Value));
            JsonHelpers.Set(marginsSection, "left", JsonHelpers.TwipToMm((int?)margins.Left?.Value));
            JsonHelpers.Set(marginsSection, "right", JsonHelpers.TwipToMm((int?)margins.Right?.Value));
            JsonHelpers.Set(marginsSection, "header", JsonHelpers.TwipToMm((int?)margins.Header?.Value));
            JsonHelpers.Set(marginsSection, "footer", JsonHelpers.TwipToMm((int?)margins.Footer?.Value));
            JsonHelpers.Set(marginsSection, "gutter", JsonHelpers.TwipToMm((int?)margins.Gutter?.Value));
        }

        JsonHelpers.Set(section, "differentFirstPageHeaderFooter", lastSection.GetFirstChild<TitlePage>() is not null);
        JsonHelpers.Set(section, "differentOddEvenHeaderFooter",
            word.MainDocumentPart?.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>() is not null);
        JsonHelpers.Set(section, "columns", (int?)(lastSection.GetFirstChild<Columns>()?.ColumnCount?.Value ?? 1));
    }
}
