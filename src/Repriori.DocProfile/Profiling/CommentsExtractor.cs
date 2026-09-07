using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "comments" section. Word stores comments in two separate parts of the
/// package: <c>WordprocessingCommentsPart</c> holds the comment text itself, and
/// the newer <c>WordprocessingCommentsExPart</c> ("comments extended") holds
/// whether each comment thread has been marked "resolved" in the Word UI — a
/// feature added years after basic comments, in its own separate XML part. A file
/// saved by an older Word version, or one where nobody ever clicked "Resolve," can
/// be entirely missing that second part, which is why "resolved"/"unresolved" can
/// legitimately come back as null rather than a count.
/// </summary>
internal static class CommentsExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var section = JsonHelpers.Obj(root, "comments");
        var commentsPart = word.MainDocumentPart?.WordprocessingCommentsPart;

        if (commentsPart is null)
        {
            JsonHelpers.Set(section, "hasComments", false);
            JsonHelpers.Set(section, "total", 0);
            JsonHelpers.Set(section, "resolved", 0);
            JsonHelpers.Set(section, "unresolved", 0);
            return;
        }

        var comments = commentsPart.Comments?.Elements<Comment>().ToList() ?? new List<Comment>();
        JsonHelpers.Set(section, "hasComments", comments.Count > 0);
        JsonHelpers.Set(section, "total", comments.Count);

        // Each entry in commentsEx.xml carries a "paraId" that matches back to a
        // paragraph inside the corresponding comment, plus a "done" flag. We collect
        // the paraIds marked done, then check which comments contain one of those
        // paragraphs — that's the only way the two parts connect to each other.
        var resolvedParaIds = new HashSet<string>(StringComparer.Ordinal);
        var commentsExPart = word.MainDocumentPart?.WordprocessingCommentsExPart;
        if (commentsExPart is not null)
        {
            foreach (var entry in commentsExPart.CommentsEx?.ChildElements ?? new DocumentFormat.OpenXml.OpenXmlElementList())
            {
                var done = entry.GetAttributes().FirstOrDefault(a => a.LocalName == "done").Value;
                var paraId = entry.GetAttributes().FirstOrDefault(a => a.LocalName == "paraId").Value;
                if (paraId is not null && (done == "1" || string.Equals(done, "true", StringComparison.OrdinalIgnoreCase)))
                {
                    resolvedParaIds.Add(paraId!);
                }
            }
        }

        if (resolvedParaIds.Count == 0)
        {
            // No commentsEx part, or nothing in it marked done — we genuinely don't
            // know the resolved/unresolved split, so we say so with null rather than
            // guessing "0 resolved."
            JsonHelpers.Set(section, "resolved", null);
            JsonHelpers.Set(section, "unresolved", null);
            return;
        }

        var resolvedCount = comments.Count(comment => comment.Descendants<Paragraph>()
            .Select(p => p.ParagraphProperties?.GetAttributes().FirstOrDefault(a => a.LocalName == "paraId").Value)
            .Any(v => v is not null && resolvedParaIds.Contains(v!)));

        JsonHelpers.Set(section, "resolved", resolvedCount);
        JsonHelpers.Set(section, "unresolved", comments.Count - resolvedCount);
    }
}
