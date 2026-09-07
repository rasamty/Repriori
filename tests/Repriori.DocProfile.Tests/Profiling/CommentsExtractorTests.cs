using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

public class CommentsExtractorTests
{
    [Fact]
    public void Extract_reports_no_comments_for_a_document_with_none()
    {
        using var docx = TestDocxBuilder.Minimal();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        CommentsExtractor.Extract(word, root);
        var comments = root["comments"]!;

        Assert.False(comments["hasComments"]!.GetValue<bool>());
        Assert.Equal(0, comments["total"]!.GetValue<int>());
        Assert.Equal(0, comments["resolved"]!.GetValue<int>());
        Assert.Equal(0, comments["unresolved"]!.GetValue<int>());
    }

    [Fact]
    public void Extract_counts_an_unresolved_comment_and_leaves_resolved_split_null()
    {
        // No commentsEx part at all here (see TestDocxBuilder.WithComment) — the
        // extractor cannot know the resolved/unresolved split and must say so
        // with null rather than guessing "0 resolved."
        using var withComment = TestDocxBuilder.WithComment(resolved: false);
        using var word = WordprocessingDocument.Open(withComment, isEditable: false);
        var root = new JsonObject();

        CommentsExtractor.Extract(word, root);
        var comments = root["comments"]!;

        Assert.True(comments["hasComments"]!.GetValue<bool>());
        Assert.Equal(1, comments["total"]!.GetValue<int>());
        Assert.Null(comments["resolved"]);
        Assert.Null(comments["unresolved"]);
    }

    [Fact]
    public void Extract_counts_a_resolved_comment_correctly()
    {
        // This is a direct regression test for the real Phase 3 bug where
        // resolvedParaIds.Add(paraId.Value!) tried to call .Value on a string
        // that was already unwrapped — a build-clean bug that only a real
        // resolved-comment fixture like this one would ever exercise.
        using var docx = TestDocxBuilder.WithComment(resolved: true);
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        CommentsExtractor.Extract(word, root);
        var comments = root["comments"]!;

        Assert.Equal(1, comments["total"]!.GetValue<int>());
        Assert.Equal(1, comments["resolved"]!.GetValue<int>());
        Assert.Equal(0, comments["unresolved"]!.GetValue<int>());
    }
}
