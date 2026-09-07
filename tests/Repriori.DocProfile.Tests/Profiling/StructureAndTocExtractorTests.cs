using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

public class StructureAndTocExtractorTests
{
    [Fact]
    public void Extract_reads_each_heading_level_into_its_own_array()
    {
        using var docx = TestDocxBuilder.WithHeadings();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        StructureAndTocExtractor.Extract(word, root);
        var structure = root["structure"]!;

        Assert.Equal("Purpose", structure["heading1Texts"]![0]!.GetValue<string>());
        Assert.Equal("Scope of this document", structure["heading2Texts"]![0]!.GetValue<string>());
        Assert.Equal("References", structure["heading3Texts"]![0]!.GetValue<string>());
        Assert.Equal("Revision History", structure["heading4Texts"]![0]!.GetValue<string>());
    }

    [Fact]
    public void Extract_detects_the_conventional_SaMD_headings_by_keyword_when_they_are_H1()
    {
        using var docx = TestDocxBuilder.WithH1Keywords();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        StructureAndTocExtractor.Extract(word, root);
        var structure = root["structure"]!;

        Assert.True(structure["hasPurposeHeading"]!.GetValue<bool>());
        Assert.True(structure["hasScopeHeading"]!.GetValue<bool>());
        Assert.True(structure["hasReferencesHeading"]!.GetValue<bool>());
        Assert.True(structure["hasRevisionHistoryHeading"]!.GetValue<bool>());
    }

    [Fact]
    public void Extract_does_not_detect_a_keyword_heading_below_H1()
    {
        // WithHeadings() puts "References" at Heading3 — sectionOrder (and every
        // hasXHeading flag built from it) only ever collects Heading1 text, by
        // design, so a lower-level heading with the exact same word is correctly
        // not detected. This is the boundary the previous test's H1 fixture does
        // not exercise, and it is worth asserting explicitly rather than leaving
        // it as an assumption.
        using var docx = TestDocxBuilder.WithHeadings();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        StructureAndTocExtractor.Extract(word, root);
        var structure = root["structure"]!;

        Assert.False(structure["hasScopeHeading"]!.GetValue<bool>());
        Assert.False(structure["hasReferencesHeading"]!.GetValue<bool>());
        Assert.False(structure["hasRevisionHistoryHeading"]!.GetValue<bool>());
    }

    [Fact]
    public void Extract_reports_no_table_of_contents_field_when_none_exists()
    {
        using var docx = TestDocxBuilder.WithHeadings();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        StructureAndTocExtractor.Extract(word, root);

        Assert.False(root["structure"]!["hasTableOfContentsField"]!.GetValue<bool>());
    }
}
