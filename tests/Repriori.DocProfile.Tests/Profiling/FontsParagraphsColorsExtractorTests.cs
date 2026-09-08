using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

public class FontsParagraphsColorsExtractorTests
{
    // FontsParagraphsColorsExtractor must run after StylesExtractor (see its own XML
    // doc comment) — every test here reproduces that same order, since fonts.bodyTypeface
    // etc. are copied from whatever styles already found.
    private static JsonObject ExtractBoth(System.IO.Stream docx)
    {
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();
        StylesExtractor.Extract(word, root);
        FontsParagraphsColorsExtractor.Extract(word, root);
        return root;
    }

    [Fact]
    public void Extract_collects_every_distinct_typeface_and_non_black_white_colour_used_directly()
    {
        using var docx = TestDocxBuilder.WithMultipleTypefacesAndColors();
        var root = ExtractBoth(docx);

        var typefaces = root["fonts"]!["typefacesUsed"]!.AsArray().Select(n => n!.GetValue<string>()).ToList();
        var colors = root["colors"]!["nonBlackWhiteHexUsed"]!.AsArray().Select(n => n!.GetValue<string>()).ToList();

        Assert.Contains("Georgia", typefaces);
        Assert.Contains("Verdana", typefaces);
        Assert.Contains("C00000", colors);
        Assert.Contains("1F4E5F", colors);
        Assert.False(root["colors"]!["bodyUsesOnlyBlackOrWhite"]!.GetValue<bool>());
    }

    [Fact]
    public void Extract_reports_no_non_black_white_colours_when_none_are_used()
    {
        using var docx = TestDocxBuilder.Minimal();
        var root = ExtractBoth(docx);

        Assert.Empty(root["colors"]!["nonBlackWhiteHexUsed"]!.AsArray());
        Assert.True(root["colors"]!["bodyUsesOnlyBlackOrWhite"]!.GetValue<bool>());
    }

    [Fact]
    public void Extract_reports_alignment_and_1point5_line_spacing()
    {
        using var docx = TestDocxBuilder.WithLineSpacing(360);
        var root = ExtractBoth(docx);
        var paragraphs = root["paragraphs"]!;

        Assert.Equal("center", paragraphs["bodyAlignment"]!.GetValue<string>());
        Assert.Equal("1.5", paragraphs["bodyLineSpacing"]!.GetValue<string>());
        Assert.Equal(12.0, paragraphs["spaceBeforePt"]!.GetValue<double>());
        Assert.Equal(6.0, paragraphs["spaceAfterPt"]!.GetValue<double>());
    }

    [Fact]
    public void Extract_reports_double_line_spacing()
    {
        using var docx = TestDocxBuilder.WithLineSpacing(480);
        var root = ExtractBoth(docx);

        Assert.Equal("double", root["paragraphs"]!["bodyLineSpacing"]!.GetValue<string>());
    }
}
