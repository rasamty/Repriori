using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

public class StylesExtractorTests
{
    [Fact]
    public void Extract_reports_null_not_false_for_every_style_slot_when_no_styles_part_exists_at_all()
    {
        using var docx = TestDocxBuilder.Minimal();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        StylesExtractor.Extract(word, root);
        var normal = root["styles"]!["normal"]!;

        Assert.Null(normal["typeface"]);
        Assert.Null(normal["bold"]);
        Assert.Null(normal["italic"]);
        Assert.Empty(root["styles"]!["styleNamesPresent"]!.AsArray());
    }

    [Fact]
    public void Extract_matches_a_style_by_its_display_name_not_only_its_bare_id()
    {
        // Regression-shaped test for the fallback branch in StylesExtractor.FillStyle:
        // a style whose id has nothing to do with "Heading1" is still found because its
        // StyleName ("Heading 1") matches — different Word templates are inconsistent
        // about which of the two a paragraph actually references.
        using var docx = TestDocxBuilder.WithDisplayNameStyledHeading();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        StylesExtractor.Extract(word, root);
        var heading1 = root["styles"]!["heading1"]!;

        Assert.Equal("Georgia", heading1["typeface"]!.GetValue<string>());
        Assert.True(heading1["bold"]!.GetValue<bool>());
        Assert.Equal("2E74B5", heading1["colorHex"]!.GetValue<string>());
    }

    [Fact]
    public void Extract_lists_every_style_name_present()
    {
        using var docx = TestDocxBuilder.WithHeadings();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        StylesExtractor.Extract(word, root);
        var names = root["styles"]!["styleNamesPresent"]!.AsArray().Select(n => n!.GetValue<string>()).ToList();

        Assert.Contains("heading 1", names);
        Assert.Contains("heading 4", names);
    }
}
