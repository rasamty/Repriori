using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

public class PageExtractorTests
{
    // A4 in twips: 210mm x 297mm, at 1440 twips per inch / 25.4mm per inch.
    private const int A4WidthTwips = 11906;
    private const int A4HeightTwips = 16838;

    [Fact]
    public void Extract_converts_A4_page_size_to_millimetres_correctly()
    {
        // This is a direct regression test for the real Phase 3 bug: converting
        // PageSize.Width/Height (OpenXml-typed values) to millimetres threw
        // InvalidCastException at runtime despite a clean build. If that
        // conversion ever breaks the same way again, this test fails immediately
        // instead of the bug surviving until someone happens to run the tool by
        // hand on a real file.
        using var docx = TestDocxBuilder.WithPageSize(A4WidthTwips, A4HeightTwips, marginTopTwips: 1440);
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        PageExtractor.Extract(word, root);
        var page = root["page"]!;

        Assert.Equal(210.0, page["widthMm"]!.GetValue<double>(), precision: 0);
        Assert.Equal(297.0, page["heightMm"]!.GetValue<double>(), precision: 0);
        Assert.Equal("portrait", page["orientation"]!.GetValue<string>());
        Assert.Equal("A4", page["paperHint"]!.GetValue<string>());
    }

    [Fact]
    public void Extract_reports_landscape_when_width_exceeds_height()
    {
        using var docx = TestDocxBuilder.WithPageSize(A4HeightTwips, A4WidthTwips, marginTopTwips: 1440);
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        PageExtractor.Extract(word, root);

        Assert.Equal("landscape", root["page"]!["orientation"]!.GetValue<string>());
    }

    [Fact]
    public void Extract_converts_margin_twips_to_millimetres()
    {
        // 1440 twips = 1 inch = 25.4mm exactly — a clean, easily-verified-by-hand number.
        using var docx = TestDocxBuilder.WithPageSize(A4WidthTwips, A4HeightTwips, marginTopTwips: 1440);
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        PageExtractor.Extract(word, root);

        Assert.Equal(25.4, root["page"]!["marginsMm"]!["top"]!.GetValue<double>(), precision: 1);
    }

    [Fact]
    public void Extract_reports_different_first_page_header_footer_odd_even_headers_and_multiple_columns()
    {
        using var docx = TestDocxBuilder.WithHeadersFootersAndPageFeatures();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        PageExtractor.Extract(word, root);
        var page = root["page"]!;

        Assert.True(page["differentFirstPageHeaderFooter"]!.GetValue<bool>());
        Assert.True(page["differentOddEvenHeaderFooter"]!.GetValue<bool>());
        Assert.Equal(2, page["columns"]!.GetValue<int>());
        Assert.Equal(1, page["sectionCount"]!.GetValue<int>());
    }

    [Fact]
    public void Extract_reports_a_single_column_and_no_odd_even_headers_for_a_plain_document()
    {
        using var docx = TestDocxBuilder.WithPageSize(A4WidthTwips, A4HeightTwips, marginTopTwips: 1440);
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        PageExtractor.Extract(word, root);
        var page = root["page"]!;

        Assert.False(page["differentFirstPageHeaderFooter"]!.GetValue<bool>());
        Assert.False(page["differentOddEvenHeaderFooter"]!.GetValue<bool>());
        Assert.Equal(1, page["columns"]!.GetValue<int>());
    }
}
