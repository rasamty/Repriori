using System.Text.Json.Nodes;

namespace Repriori.DocProfile.Comparing;

/// <summary>
/// Walks two profile JSON trees side by side and lists every field where they
/// disagree. This class only ever reports facts ("these two values are
/// different") — it never decides whether a difference matters. That judgment
/// belongs to <see cref="ComparisonRuleSet"/>, kept deliberately separate so the
/// same diff can be reused by different rule sets (a strict one for a final
/// approval gate, a looser one for an early draft check) without recomputing it.
/// </summary>
public static class ProfileComparer
{
    /// <summary>
    /// Top-level fields describing the extraction itself, not the document's
    /// formatting — two profiles of the exact same file, extracted a second
    /// apart, will always disagree on these. Comparing them would report "the
    /// document changed" when nothing about the document actually did.
    /// </summary>
    private static readonly HashSet<string> DefaultIgnoredPaths = new(StringComparer.Ordinal)
    {
        "profileId", "sourceFile", "extractedAt",
    };

    public static IReadOnlyList<ProfileDiff> Compare(JsonNode? golden, JsonNode? submission, IReadOnlySet<string>? ignoredPaths = null)
    {
        var diffs = new List<ProfileDiff>();
        Walk(string.Empty, golden, submission, diffs, ignoredPaths ?? DefaultIgnoredPaths);
        return diffs;
    }

    private static void Walk(string path, JsonNode? golden, JsonNode? submission, List<ProfileDiff> diffs, IReadOnlySet<string> ignoredPaths)
    {
        if (ignoredPaths.Contains(path)) return;

        // Recurse only when both sides are genuinely comparable objects — this
        // is what lets a single differing field (say, fonts.bodySizePt) get
        // reported on its own, rather than the whole "fonts" section being
        // flagged as one big opaque difference.
        if (golden is JsonObject goldenObject && submission is JsonObject submissionObject)
        {
            var allKeys = goldenObject.Select(kv => kv.Key)
                .Union(submissionObject.Select(kv => kv.Key), StringComparer.Ordinal)
                .OrderBy(k => k, StringComparer.Ordinal);

            foreach (var key in allKeys)
            {
                var childPath = path.Length == 0 ? key : $"{path}.{key}";
                goldenObject.TryGetPropertyValue(key, out var goldenChild);
                submissionObject.TryGetPropertyValue(key, out var submissionChild);
                Walk(childPath, goldenChild, submissionChild, diffs, ignoredPaths);
            }
            return;
        }

        // Arrays and plain values are compared as a single unit rather than
        // element by element. Diffing an array properly (a heading added in the
        // middle, say) means deciding how to align old and new items with each
        // other — a genuinely harder problem than this phase needs to solve; an
        // array that changed at all is reported once, as one diff, which is
        // honest about what is and isn't being checked rather than pretending to
        // pinpoint exactly what moved.
        if (!JsonNode.DeepEquals(golden, submission))
        {
            diffs.Add(new ProfileDiff(path, golden?.DeepClone(), submission?.DeepClone()));
        }
    }

    /// <summary>
    /// Reads the value at a dotted path (e.g. "page.marginsMm.top") out of a
    /// profile — the same path notation ProfileDiff.FieldPath and
    /// ComparisonRule.Field both use. Only walks plain objects; a path that
    /// tries to reach into an array returns null rather than throwing, since
    /// rules are meant to target the kind of single, well-known scalar fields
    /// the default rule set uses, not arbitrary array contents.
    /// </summary>
    public static JsonNode? ReadPath(JsonNode? root, string path)
    {
        var current = root;
        foreach (var segment in path.Split('.'))
        {
            if (current is not JsonObject obj) return null;
            obj.TryGetPropertyValue(segment, out current);
        }
        return current;
    }
}
