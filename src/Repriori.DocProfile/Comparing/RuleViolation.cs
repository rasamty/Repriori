namespace Repriori.DocProfile.Comparing;

/// <summary>One rule that a submission actually failed, with a human-readable reason.</summary>
public sealed record RuleViolation(ComparisonRule Rule, string Detail)
{
    public override string ToString() => $"{Rule.Field}: {Rule.Description} ({Detail})";
}
