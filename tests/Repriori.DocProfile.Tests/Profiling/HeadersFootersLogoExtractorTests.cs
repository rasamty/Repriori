using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

public class HeadersFootersLogoExtractorTests
{
    [Fact]
    public void Extract_reports_no_headers_or_footers_for_a_plain_document()
    {
        using var docx = TestDocxBuilder.Minimal();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        HeadersFootersLogoExtractor.Extract(word, root);
        var section = root["headersFooters"]!;

        Assert.False(section["hasDefaultHeader"]!.GetValue<bool>());
        Assert.False(section["hasDefaultFooter"]!.GetValue<bool>());
        Assert.Equal(0, section["headerImageCount"]!.GetValue<int>());
    }

    [Fact]
    public void Extract_detects_default_first_and_even_headers_and_footers_and_a_real_page_field()
    {
        using var docx = TestDocxBuilder.WithHeadersFootersAndPageFeatures();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        HeadersFootersLogoExtractor.Extract(word, root);
        var section = root["headersFooters"]!;

        Assert.True(section["hasDefaultHeader"]!.GetValue<bool>());
        Assert.True(section["hasFirstHeader"]!.GetValue<bool>());
        Assert.True(section["hasEvenHeader"]!.GetValue<bool>());
        Assert.True(section["hasDefaultFooter"]!.GetValue<bool>());
        Assert.True(section["hasFirstFooter"]!.GetValue<bool>());
        Assert.True(section["hasEvenFooter"]!.GetValue<bool>());
        Assert.True(section["hasPageNumberField"]!.GetValue<bool>());
        Assert.Contains("Default header", section["headerTextSample"]!.GetValue<string>());
    }

    [Fact]
    public void Extract_flags_a_header_image_whose_name_mentions_logo_as_likely_a_company_logo()
    {
        using var docx = TestDocxBuilder.WithHeaderImage("Company Logo");
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        HeadersFootersLogoExtractor.Extract(word, root);
        var logo = root["logo"]!;

        Assert.True(logo["headerContainsImage"]!.GetValue<bool>());
        Assert.Equal("Company Logo", logo["headerImageNameOrAltSample"]!.GetValue<string>());
        Assert.True(logo["likelyCompanyLogo"]!.GetValue<bool>());
        Assert.Contains("logo", logo["likelyCompanyLogoReason"]!.GetValue<string>(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Extract_does_not_claim_an_unlabelled_header_image_is_a_logo()
    {
        using var docx = TestDocxBuilder.WithHeaderImage("Picture 1");
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        HeadersFootersLogoExtractor.Extract(word, root);
        var logo = root["logo"]!;

        Assert.True(logo["headerContainsImage"]!.GetValue<bool>());
        Assert.False(logo["likelyCompanyLogo"]!.GetValue<bool>());
        Assert.Contains("cannot prove", logo["likelyCompanyLogoReason"]!.GetValue<string>(), StringComparison.OrdinalIgnoreCase);
    }
}
