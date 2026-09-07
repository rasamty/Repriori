using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

public class DocumentPropertiesExtractorTests
{
    [Fact]
    public void Extract_reads_title_and_creator_from_core_properties()
    {
        using var docx = TestDocxBuilder.WithCoreProperties(title: "Repriori Test Document", creator: "Rasam");
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        DocumentPropertiesExtractor.Extract(word, root);
        var document = root["document"]!;

        Assert.Equal("Repriori Test Document", document["coreTitle"]!.GetValue<string>());
        Assert.Equal("Rasam", document["coreCreator"]!.GetValue<string>());
    }

    [Fact]
    public void Extract_reports_null_title_rather_than_an_empty_string_when_unset()
    {
        // JsonHelpers.Set treats "" the same as null on the way in (see its own
        // comment) — this proves that rule actually applies to a real field, not
        // just in theory.
        using var docx = TestDocxBuilder.Minimal();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        DocumentPropertiesExtractor.Extract(word, root);

        Assert.Null(root["document"]!["coreTitle"]);
    }

    [Fact]
    public void Extract_reports_not_protected_for_an_ordinary_document()
    {
        using var docx = TestDocxBuilder.Minimal();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        DocumentPropertiesExtractor.Extract(word, root);

        Assert.False(root["document"]!["isProtected"]!.GetValue<bool>());
    }
}
