using System.Text.Json.Nodes;

namespace Repriori.DocProfile.Comparing;

/// <summary>
/// The two kinds of check a rule can make. Deliberately just two, rather than a
/// more general "expression language" — every example rule from this project's
/// own planning conversation (tracked changes must be accepted, comments must be
/// resolved, fonts must match the template) reduces to one of these two shapes,
/// and a small, fixed set of rule types is far easier to read a rules file and
/// trust than an embedded scripting language would be.
/// </summary>
public enum ComparisonRuleType
{
    /// <summary>The submission's value at Field must exactly equal ExpectedValue, regardless of what the golden profile says.</summary>
    MustEqual,

    /// <summary>The submission's value at Field must match the golden profile's value at that same field.</summary>
    MustMatchGolden,
}

/// <summary>
/// One configurable pass/fail check. A whole set of these is what
/// <see cref="ComparisonRuleSet"/> loads from JSON — see
/// default-comparison-rules.json for the default set, and its own comments for
/// why format-check rules live in JSON rather than code.
/// </summary>
public sealed record ComparisonRule(string Field, ComparisonRuleType Type, JsonNode? ExpectedValue, string Description);
