using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

// The "nothing present" baseline lives in ListsAndFiguresExtractorTests.cs. This file
// covers the fuller, "an image genuinely exists" cases that baseline deliberately
// doesn't — see that file's own comment for why it started out lighter.
public class FiguresExtractorTests
{
    [Fact]
    public void Extract_counts_an_image_with_alt_text()
    {
        using var docx = TestDocxBuilder.WithFigure(altText: "A bar chart of quarterly results", captionText: null);
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        FiguresExtractor.Extract(word, root);
        var figures = root["figures"]!;

        Assert.Equal(1, figures["imageCount"]!.GetValue<int>());
        Assert.Equal(1, figures["imagesWithAltText"]!.GetValue<int>());
        Assert.Equal(0, figures["imagesMissingAltText"]!.GetValue<int>());
    }

    [Fact]
    public void Extract_counts_an_image_missing_alt_text()
    {
        using var docx = TestDocxBuilder.WithFigure(altText: null, captionText: null);
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        FiguresExtractor.Extract(word, root);
        var figures = root["figures"]!;

        Assert.Equal(0, figures["imagesWithAltText"]!.GetValue<int>());
        Assert.Equal(1, figures["imagesMissingAltText"]!.GetValue<int>());
    }

    [Fact]
    public void Extract_counts_a_caption_and_its_text()
    {
        using var docx = TestDocxBuilder.WithFigure(altText: "x", captionText: "Figure 1: Something interesting");
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        FiguresExtractor.Extract(word, root);
        var figures = root["figures"]!;

        Assert.Equal(1, figures["captionCount"]!.GetValue<int>());
        Assert.Contains("Figure 1", figures["captionTexts"]![0]!.GetValue<string>());
    }
}
