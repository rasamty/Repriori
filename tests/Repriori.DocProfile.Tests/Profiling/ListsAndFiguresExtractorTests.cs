using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

// Kept lighter than the other extractors on purpose. ListsExtractor mostly answers a
// question ("is there any numbering") that needs real numbering.xml wiring to test
// beyond this, and hasBullets is permanently null regardless (see its own code
// comment) — nothing more to prove there without a materially bigger fixture.
// FiguresExtractor's "an image genuinely exists" cases turned out not to need real
// embedded picture bytes after all (the extractor only reads a drawing's own declared
// name/description, never its content) — see FiguresExtractorTests.cs for those.
// What's tested here is the baseline this file was always meant to prove: the
// "nothing present" case is correct for both extractors.
public class ListsAndFiguresExtractorTests
{
    [Fact]
    public void Lists_Extract_reports_no_numbering_for_a_plain_document()
    {
        using var docx = TestDocxBuilder.Minimal();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        ListsExtractor.Extract(word, root);

        Assert.False(root["lists"]!["hasNumbering"]!.GetValue<bool>());
        Assert.Null(root["lists"]!["hasBullets"]); // deliberately always null — see the extractor's own comment
    }

    [Fact]
    public void Figures_Extract_reports_null_alt_text_stats_when_there_are_no_images()
    {
        using var docx = TestDocxBuilder.Minimal();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        FiguresExtractor.Extract(word, root);
        var figures = root["figures"]!;

        Assert.Equal(0, figures["imageCount"]!.GetValue<int>());
        // Null, not 0 — "0 of 0 images have alt text" is not a real statistic,
        // and this proves that distinction actually holds for a document that
        // genuinely has zero images, not just in the extractor's doc comment.
        Assert.Null(figures["imagesWithAltText"]);
        Assert.Null(figures["imagesMissingAltText"]);
    }
}
