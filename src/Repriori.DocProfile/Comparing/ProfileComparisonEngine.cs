using System.Text.Json.Nodes;

namespace Repriori.DocProfile.Comparing;

/// <summary>
/// The public entry point for comparing two profiles — the Comparing-folder
/// equivalent of WordProfileExtractor. Like that class, this one does almost no
/// work itself: it parses the two profile strings, hands them to
/// <see cref="ProfileComparer"/> for the raw diff, then to a
/// <see cref="ComparisonRuleSet"/> to judge which differences actually matter.
/// </summary>
public static class ProfileComparisonEngine
{
    /// <summary>
    /// Compares a submission profile against a golden one. Both parameters are
    /// profile JSON strings — the exact output of
    /// <see cref="WordProfileExtractor.FromDocx(string)"/> — not file paths and
    /// not .docx files; extraction and comparison are two separate, independently
    /// testable steps.
    /// </summary>
    public static ComparisonResult Compare(string goldenProfileJson, string submissionProfileJson, ComparisonRuleSet? ruleSet = null)
    {
        ruleSet ??= ComparisonRuleSet.Default();

        var golden = JsonNode.Parse(goldenProfileJson) ?? throw new InvalidDataException("Golden profile is not valid JSON.");
        var submission = JsonNode.Parse(submissionProfileJson) ?? throw new InvalidDataException("Submission profile is not valid JSON.");

        var diffs = ProfileComparer.Compare(golden, submission);
        var violations = ruleSet.Evaluate(submission, diffs);

        return new ComparisonResult(Passed: violations.Count == 0, diffs, violations);
    }
}
