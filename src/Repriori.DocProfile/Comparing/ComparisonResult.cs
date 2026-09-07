namespace Repriori.DocProfile.Comparing;

/// <summary>
/// The full outcome of comparing a submission against a golden profile: every
/// difference found (informational — see ProfileComparer), and which of those
/// differences actually violate a configured rule. Passed is true exactly when
/// Violations is empty — a submission can have plenty of Diffs (a changed
/// revision number, an extra sentence) and still Pass, as long as none of them
/// touch a field a rule actually cares about.
/// </summary>
public sealed record ComparisonResult(bool Passed, IReadOnlyList<ProfileDiff> Diffs, IReadOnlyList<RuleViolation> Violations);
