using System.Reflection;
using System.Text.Json.Nodes;
using Repriori.DocProfile.Validation;

namespace Repriori.DocProfile.Comparing;

/// <summary>
/// A named collection of <see cref="ComparisonRule"/>s, loaded from JSON rather
/// than written in C#. This matches the pattern the whole Repriori plan commits
/// to for gates and requirements generally (Part 2 of the design document):
/// policy — "what must be true before a document passes" — is authored as
/// reviewable JSON that a non-programmer can read and edit; only the mechanics
/// of checking it are code. Editing which fields matter, or what a passing value
/// looks like, never requires rebuilding this library.
/// </summary>
public sealed class ComparisonRuleSet
{
    private readonly IReadOnlyList<ComparisonRule> _rules;

    private ComparisonRuleSet(IReadOnlyList<ComparisonRule> rules) => _rules = rules;

    public IReadOnlyList<ComparisonRule> Rules => _rules;

    /// <summary>The rule set bundled with this library — see default-comparison-rules.json.</summary>
    public static ComparisonRuleSet Default()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = $"{assembly.GetName().Name}.default-comparison-rules.json";
        using var stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException($"Embedded resource '{resourceName}' was not found.");
        using var reader = new StreamReader(stream);
        return FromJson(reader.ReadToEnd());
    }

    /// <summary>Loads a custom rule set from a JSON file on disk — see the schema comment on <see cref="FromJson"/>.</summary>
    public static ComparisonRuleSet FromFile(string path)
    {
        var readable = FileValidation.EnsureFileIsReadable(path);
        if (!readable.IsOk) throw new IOException(readable.Reason);
        return FromJson(File.ReadAllText(path));
    }

    /// <summary>
    /// Parses a rule set from JSON shaped like:
    /// <code>
    /// { "rules": [
    ///   { "field": "revisions.hasUnacceptedTrackedChanges", "type": "mustEqual", "value": false, "description": "..." },
    ///   { "field": "page.orientation", "type": "mustMatchGolden", "description": "..." }
    /// ]}
    /// </code>
    /// "value" is required for mustEqual rules and ignored for mustMatchGolden ones (there is nothing fixed to compare against — the golden profile supplies it at comparison time).
    /// </summary>
    public static ComparisonRuleSet FromJson(string json)
    {
        var root = JsonNode.Parse(json)?.AsObject()
            ?? throw new InvalidDataException("Rule set JSON must be an object with a \"rules\" array.");
        var rulesArray = root["rules"]?.AsArray()
            ?? throw new InvalidDataException("Rule set JSON must contain a \"rules\" array.");

        var rules = new List<ComparisonRule>();
        foreach (var node in rulesArray)
        {
            var ruleObject = node!.AsObject();
            var field = ruleObject["field"]?.GetValue<string>()
                ?? throw new InvalidDataException("Every rule needs a \"field\".");
            var typeText = ruleObject["type"]?.GetValue<string>()
                ?? throw new InvalidDataException($"Rule for '{field}' needs a \"type\".");
            var type = typeText.Equals("mustEqual", StringComparison.OrdinalIgnoreCase) ? ComparisonRuleType.MustEqual
                : typeText.Equals("mustMatchGolden", StringComparison.OrdinalIgnoreCase) ? ComparisonRuleType.MustMatchGolden
                : throw new InvalidDataException($"Unknown rule type '{typeText}' for field '{field}'. Expected \"mustEqual\" or \"mustMatchGolden\".");
            var description = ruleObject["description"]?.GetValue<string>() ?? $"{field} must satisfy {typeText}.";
            var expectedValue = type == ComparisonRuleType.MustEqual
                ? ruleObject["value"] ?? throw new InvalidDataException($"mustEqual rule for '{field}' needs a \"value\".")
                : null;

            rules.Add(new ComparisonRule(field, type, expectedValue?.DeepClone(), description));
        }

        return new ComparisonRuleSet(rules);
    }

    /// <summary>
    /// Checks every rule and returns the ones the submission actually failed.
    /// diffs is the already-computed output of <see cref="ProfileComparer.Compare"/>
    /// — passed in rather than recomputed here, so the same diff pass serves both
    /// the informational diff list and the rule check.
    /// </summary>
    public IReadOnlyList<RuleViolation> Evaluate(JsonNode submissionRoot, IReadOnlyList<ProfileDiff> diffs)
    {
        var violations = new List<RuleViolation>();

        foreach (var rule in _rules)
        {
            switch (rule.Type)
            {
                case ComparisonRuleType.MustMatchGolden:
                    // A rule field that never showed up in the diff list is, by
                    // definition, identical between golden and submission —
                    // reusing the diff avoids walking the JSON tree a second time.
                    var diff = diffs.FirstOrDefault(d => d.FieldPath == rule.Field);
                    if (diff is not null)
                    {
                        violations.Add(new RuleViolation(rule,
                            $"golden={diff.GoldenValue?.ToJsonString() ?? "null"}, submission={diff.SubmissionValue?.ToJsonString() ?? "null"}"));
                    }
                    break;

                case ComparisonRuleType.MustEqual:
                    var actual = ProfileComparer.ReadPath(submissionRoot, rule.Field);
                    if (!JsonNode.DeepEquals(actual, rule.ExpectedValue))
                    {
                        violations.Add(new RuleViolation(rule,
                            $"expected {rule.ExpectedValue?.ToJsonString() ?? "null"}, got {actual?.ToJsonString() ?? "null"}"));
                    }
                    break;
            }
        }

        return violations;
    }
}
