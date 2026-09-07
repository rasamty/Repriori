using System.Text.Json.Nodes;
using Repriori.DocProfile.Comparing;
using Xunit;

namespace Repriori.DocProfile.Tests.Comparing;

public class ProfileComparerTests
{
    [Fact]
    public void Compare_finds_no_diffs_for_identical_profiles()
    {
        var golden = JsonNode.Parse("""{"page":{"orientation":"portrait"}}""");
        var submission = JsonNode.Parse("""{"page":{"orientation":"portrait"}}""");

        var diffs = ProfileComparer.Compare(golden, submission);

        Assert.Empty(diffs);
    }

    [Fact]
    public void Compare_reports_a_nested_field_by_its_own_dotted_path()
    {
        var golden = JsonNode.Parse("""{"page":{"orientation":"portrait","paperHint":"A4"}}""");
        var submission = JsonNode.Parse("""{"page":{"orientation":"landscape","paperHint":"A4"}}""");

        var diffs = ProfileComparer.Compare(golden, submission);

        var diff = Assert.Single(diffs);
        Assert.Equal("page.orientation", diff.FieldPath);
        Assert.Equal("portrait", diff.GoldenValue!.GetValue<string>());
        Assert.Equal("landscape", diff.SubmissionValue!.GetValue<string>());
    }

    [Fact]
    public void Compare_ignores_extraction_metadata_fields_by_default()
    {
        var golden = JsonNode.Parse("""{"profileId":"a","sourceFile":"a.docx","extractedAt":"2026-01-01","page":{"orientation":"portrait"}}""");
        var submission = JsonNode.Parse("""{"profileId":"b","sourceFile":"b.docx","extractedAt":"2026-06-01","page":{"orientation":"portrait"}}""");

        var diffs = ProfileComparer.Compare(golden, submission);

        Assert.Empty(diffs);
    }

    [Fact]
    public void Compare_reports_a_changed_array_once_rather_than_element_by_element()
    {
        var golden = JsonNode.Parse("""{"structure":{"heading1Texts":["Purpose","Scope"]}}""");
        var submission = JsonNode.Parse("""{"structure":{"heading1Texts":["Purpose","Scope","Extra Section"]}}""");

        var diffs = ProfileComparer.Compare(golden, submission);

        var diff = Assert.Single(diffs);
        Assert.Equal("structure.heading1Texts", diff.FieldPath);
    }

    [Fact]
    public void ReadPath_reads_a_nested_scalar_by_dotted_path()
    {
        var root = JsonNode.Parse("""{"page":{"marginsMm":{"top":25.4}}}""");

        var value = ProfileComparer.ReadPath(root, "page.marginsMm.top");

        Assert.Equal(25.4, value!.GetValue<double>());
    }

    [Fact]
    public void ReadPath_returns_null_for_a_path_that_does_not_exist()
    {
        var root = JsonNode.Parse("""{"page":{"orientation":"portrait"}}""");

        var value = ProfileComparer.ReadPath(root, "page.doesNotExist");

        Assert.Null(value);
    }
}
