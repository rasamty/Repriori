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

    [Fact]
    public void Extract_reports_the_real_protection_type_not_the_wrapper_types_own_ToString()
    {
        // Regression test for the real Phase 5 bug: protection.Edit.Value.ToString()
        // printed "DocumentProtectionValues { }" against a real protected file —
        // .InnerText is what actually reads "readOnly" correctly. No Phase 4
        // fixture exercised this path at all, which is exactly why it took a
        // real file to find it.
        using var docx = TestDocxBuilder.WithReadOnlyProtection();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        DocumentPropertiesExtractor.Extract(word, root);
        var document = root["document"]!;

        Assert.True(document["isProtected"]!.GetValue<bool>());
        Assert.Equal("readOnly", document["protectionType"]!.GetValue<string>());
    }

    [Fact]
    public void Extract_reports_write_reservation_and_recommendation_when_write_protection_is_set()
    {
        // WriteProtection is a genuinely different Word feature from DocumentProtection
        // (the "recommend read-only" prompt on open, vs. enforced editing restrictions)
        // — no Phase 4 fixture set this at all, which is exactly why it was still an
        // untested branch going into this pass.
        using var docx = TestDocxBuilder.WithWriteProtection(recommended: true);
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        DocumentPropertiesExtractor.Extract(word, root);
        var document = root["document"]!;

        Assert.True(document["writeReservation"]!.GetValue<bool>());
        Assert.True(document["readOnlyRecommended"]!.GetValue<bool>());
    }

    [Fact]
    public void Extract_reports_write_reservation_without_a_recommendation_flag()
    {
        using var docx = TestDocxBuilder.WithWriteProtection(recommended: false);
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        DocumentPropertiesExtractor.Extract(word, root);
        var document = root["document"]!;

        Assert.True(document["writeReservation"]!.GetValue<bool>());
        Assert.Null(document["readOnlyRecommended"]);
    }

    [Fact]
    public void Extract_reports_the_compatibility_mode_when_set()
    {
        using var docx = TestDocxBuilder.WithCompatibilityMode("15");
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        DocumentPropertiesExtractor.Extract(word, root);

        Assert.Equal("15", root["document"]!["compatibilityMode"]!.GetValue<string>());
    }

    [Fact]
    public void Extract_reports_the_first_non_empty_language_found()
    {
        using var docx = TestDocxBuilder.WithLanguage("en-GB");
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        DocumentPropertiesExtractor.Extract(word, root);

        Assert.Equal("en-GB", root["document"]!["language"]!.GetValue<string>());
    }
}
