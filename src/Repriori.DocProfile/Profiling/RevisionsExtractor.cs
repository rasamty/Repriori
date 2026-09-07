using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "revisions" section: whether the document still has "Track Changes"
/// edits sitting in it that nobody has accepted or rejected yet. This matters for
/// the format-check gate (Part 2.6 of the plan) — a document with unresolved
/// tracked changes generally should not be allowed to pass as "final."
/// </summary>
internal static class RevisionsExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var section = JsonHelpers.Obj(root, "revisions");
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return;

        // InsertedRun/DeletedRun/MoveFromRun/MoveToRun are the actual XML elements
        // Word wraps tracked-change content in. Counting them (rather than trying
        // to summarise *what* changed) is enough to answer "is anything still
        // pending a decision" without attempting a full diff of the edits.
        var inserted = body.Descendants<InsertedRun>().Count();
        var deleted = body.Descendants<DeletedRun>().Count();
        var moved = body.Descendants<MoveFromRun>().Count() + body.Descendants<MoveToRun>().Count();

        JsonHelpers.Set(section, "insertedRunCount", inserted);
        JsonHelpers.Set(section, "deletedRunCount", deleted);
        JsonHelpers.Set(section, "moveCount", moved);
        JsonHelpers.Set(section, "hasUnacceptedTrackedChanges", inserted + deleted + moved > 0);
    }
}
