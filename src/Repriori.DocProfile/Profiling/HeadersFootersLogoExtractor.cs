using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "headersFooters" and "logo" sections. Headers and footers are stored
/// as their own small XML parts inside the package (not inline in the main body),
/// and Word can have up to three variants of each — default, first-page, and
/// even-page — which is why this extractor has to look in more than one place.
/// </summary>
internal static class HeadersFootersLogoExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var headerFooterSection = JsonHelpers.Obj(root, "headersFooters");
        var logoSection = JsonHelpers.Obj(root, "logo");
        var main = word.MainDocumentPart;
        var body = main?.Document?.Body;
        if (main is null || body is null) return;

        JsonHelpers.Set(headerFooterSection, "hasDefaultHeader", main.HeaderParts.Any());
        JsonHelpers.Set(headerFooterSection, "hasDefaultFooter", main.FooterParts.Any());

        // HeaderReference/FooterReference elements in the body are what actually
        // say *which* header/footer part applies to first/even/default pages —
        // the parts themselves don't know their own role.
        var headerRefs = body.Descendants<HeaderReference>().ToList();
        JsonHelpers.Set(headerFooterSection, "hasFirstHeader", headerRefs.Any(r => r.Type?.Value == HeaderFooterValues.First));
        JsonHelpers.Set(headerFooterSection, "hasEvenHeader", headerRefs.Any(r => r.Type?.Value == HeaderFooterValues.Even));

        var footerRefs = body.Descendants<FooterReference>().ToList();
        JsonHelpers.Set(headerFooterSection, "hasFirstFooter", footerRefs.Any(r => r.Type?.Value == HeaderFooterValues.First));
        JsonHelpers.Set(headerFooterSection, "hasEvenFooter", footerRefs.Any(r => r.Type?.Value == HeaderFooterValues.Even));

        JsonHelpers.Set(headerFooterSection, "headerTextSample", JsonHelpers.Clip(string.Join(" ", main.HeaderParts.Select(JsonHelpers.PartText))));
        JsonHelpers.Set(headerFooterSection, "footerTextSample", JsonHelpers.Clip(string.Join(" ", main.FooterParts.Select(JsonHelpers.PartText))));
        JsonHelpers.Set(headerFooterSection, "hasPageNumberField", main.HeaderParts.Any(JsonHelpers.HasPageField) || main.FooterParts.Any(JsonHelpers.HasPageField));

        var headerDrawings = main.HeaderParts
            .SelectMany(p => p.RootElement?.Descendants<Drawing>() ?? Enumerable.Empty<Drawing>())
            .ToList();
        JsonHelpers.Set(headerFooterSection, "headerImageCount", headerDrawings.Count);
        JsonHelpers.Set(logoSection, "headerContainsImage", headerDrawings.Count > 0);

        // A drawing's "name" or "description" is whatever the author typed when
        // inserting the image (or left as Word's default, e.g. "Picture 1"). If it
        // happens to mention "logo," that's a useful hint — but only a hint: there
        // is nothing in the file format that actually certifies an image as a
        // company logo, so this can never be more than "likely."
        var sample = headerDrawings
            .Select(d =>
                d.Descendants<DW.DocProperties>().FirstOrDefault()?.Name?.Value
                ?? d.Descendants<DW.DocProperties>().FirstOrDefault()?.Description?.Value
                ?? d.Descendants<A.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value)
            .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
        JsonHelpers.Set(logoSection, "headerImageNameOrAltSample", sample);

        var looksLikeLogo = sample?.Contains("logo", StringComparison.OrdinalIgnoreCase) == true;
        JsonHelpers.Set(logoSection, "likelyCompanyLogo", headerDrawings.Count > 0 && looksLikeLogo);
        JsonHelpers.Set(logoSection, "likelyCompanyLogoReason",
            headerDrawings.Count == 0 ? "no image in header"
            : looksLikeLogo ? "header image name/alt contains \"logo\""
            : "header has an image; cannot prove it is the company logo");
    }
}
