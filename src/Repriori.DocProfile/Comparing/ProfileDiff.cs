using System.Text.Json.Nodes;

namespace Repriori.DocProfile.Comparing;

/// <summary>
/// One place where a golden profile and a submission profile disagree.
/// FieldPath uses dotted notation matching the profile's own JSON shape, e.g.
/// "page.orientation" or "fonts.bodyTypeface" — the same notation
/// <see cref="ComparisonRule"/> uses to name a field.
/// </summary>
public sealed record ProfileDiff(string FieldPath, JsonNode? GoldenValue, JsonNode? SubmissionValue)
{
    public override string ToString() =>
        $"{FieldPath}: golden={GoldenValue?.ToJsonString() ?? "null"}, submission={SubmissionValue?.ToJsonString() ?? "null"}";
}
