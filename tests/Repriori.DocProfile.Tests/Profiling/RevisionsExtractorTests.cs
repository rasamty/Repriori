using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

public class RevisionsExtractorTests
{
    [Fact]
    public void Extract_reports_no_tracked_changes_for_a_clean_document()
    {
        using var docx = TestDocxBuilder.Minimal();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        RevisionsExtractor.Extract(word, root);
        var revisions = root["revisions"]!;

        Assert.False(revisions["hasUnacceptedTrackedChanges"]!.GetValue<bool>());
        Assert.Equal(0, revisions["insertedRunCount"]!.GetValue<int>());
        Assert.Equal(0, revisions["deletedRunCount"]!.GetValue<int>());
    }

    [Fact]
    public void Extract_counts_inserted_and_deleted_runs_and_flags_them_as_unaccepted()
    {
        using var docx = TestDocxBuilder.WithTrackedChanges();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        RevisionsExtractor.Extract(word, root);
        var revisions = root["revisions"]!;

        Assert.Equal(1, revisions["insertedRunCount"]!.GetValue<int>());
        Assert.Equal(1, revisions["deletedRunCount"]!.GetValue<int>());
        Assert.True(revisions["hasUnacceptedTrackedChanges"]!.GetValue<bool>());
    }
}
