using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

public class EmphasisExtractorTests
{
    [Fact]
    public void Extract_computes_an_exact_bold_percentage_from_known_run_lengths()
    {
        // TestDocxBuilder.WithKnownEmphasis is built as 20 plain characters plus
        // 10 bold characters — 10 / 30 = 33.3%, an exact number worth asserting
        // rather than a vague "greater than zero."
        using var docx = TestDocxBuilder.WithKnownEmphasis();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        EmphasisExtractor.Extract(word, root);
        var body = root["emphasis"]!["body"]!;

        Assert.Equal(30, body["chars"]!.GetValue<int>());
        Assert.Equal(33.3, body["boldPct"]!.GetValue<double>(), precision: 1);
        Assert.Equal(0.0, body["italicPct"]!.GetValue<double>(), precision: 1);
    }

    [Fact]
    public void Extract_reports_null_percentages_for_a_paragraph_with_no_text()
    {
        using var docx = TestDocxBuilder.Minimal(text: "");
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        EmphasisExtractor.Extract(word, root);
        var body = root["emphasis"]!["body"]!;

        Assert.Equal(0, body["chars"]!.GetValue<int>());
        Assert.Null(body["boldPct"]);
        Assert.Null(body["italicPct"]);
    }
}
