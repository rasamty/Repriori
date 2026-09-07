using System.Text.Json.Nodes;
using Repriori.DocProfile.Comparing;
using Xunit;

namespace Repriori.DocProfile.Tests.Comparing;

public class ComparisonRuleSetTests
{
    [Fact]
    public void Default_loads_without_error_and_contains_at_least_one_rule_of_each_type()
    {
        var ruleSet = ComparisonRuleSet.Default();

        Assert.Contains(ruleSet.Rules, r => r.Type == ComparisonRuleType.MustEqual);
        Assert.Contains(ruleSet.Rules, r => r.Type == ComparisonRuleType.MustMatchGolden);
    }

    [Fact]
    public void FromJson_parses_a_mustEqual_rule_including_its_expected_value()
    {
        var ruleSet = ComparisonRuleSet.FromJson("""
            { "rules": [
              { "field": "comments.unresolved", "type": "mustEqual", "value": 0, "description": "must be zero" }
            ]}
            """);

        var rule = Assert.Single(ruleSet.Rules);
        Assert.Equal("comments.unresolved", rule.Field);
        Assert.Equal(ComparisonRuleType.MustEqual, rule.Type);
        Assert.Equal(0, rule.ExpectedValue!.GetValue<int>());
    }

    [Fact]
    public void FromJson_rejects_a_mustEqual_rule_with_no_value()
    {
        var ex = Assert.Throws<InvalidDataException>(() => ComparisonRuleSet.FromJson("""
            { "rules": [ { "field": "x", "type": "mustEqual", "description": "no value given" } ]}
            """));
        Assert.Contains("value", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FromJson_rejects_an_unknown_rule_type()
    {
        Assert.Throws<InvalidDataException>(() => ComparisonRuleSet.FromJson("""
            { "rules": [ { "field": "x", "type": "mustBeAwesome", "description": "not real" } ]}
            """));
    }

    [Fact]
    public void Evaluate_flags_a_mustEqual_violation_with_a_clear_reason()
    {
        var ruleSet = ComparisonRuleSet.FromJson("""
            { "rules": [ { "field": "comments.unresolved", "type": "mustEqual", "value": 0, "description": "all comments resolved" } ]}
            """);
        var submission = JsonNode.Parse("""{"comments":{"unresolved":3}}""")!;

        var violations = ruleSet.Evaluate(submission, diffs: Array.Empty<ProfileDiff>());

        var violation = Assert.Single(violations);
        Assert.Equal("comments.unresolved", violation.Rule.Field);
        Assert.Contains("3", violation.Detail);
    }

    [Fact]
    public void Evaluate_passes_a_mustEqual_rule_when_the_submission_already_matches()
    {
        var ruleSet = ComparisonRuleSet.FromJson("""
            { "rules": [ { "field": "comments.unresolved", "type": "mustEqual", "value": 0, "description": "all comments resolved" } ]}
            """);
        var submission = JsonNode.Parse("""{"comments":{"unresolved":0}}""")!;

        var violations = ruleSet.Evaluate(submission, diffs: Array.Empty<ProfileDiff>());

        Assert.Empty(violations);
    }

    [Fact]
    public void Evaluate_flags_a_mustMatchGolden_violation_when_the_field_appears_in_the_diff_list()
    {
        var ruleSet = ComparisonRuleSet.FromJson("""
            { "rules": [ { "field": "page.orientation", "type": "mustMatchGolden", "description": "orientation must match" } ]}
            """);
        var submission = JsonNode.Parse("""{"page":{"orientation":"landscape"}}""")!;
        var diffs = new[] { new ProfileDiff("page.orientation", JsonValue.Create("portrait"), JsonValue.Create("landscape")) };

        var violations = ruleSet.Evaluate(submission, diffs);

        Assert.Single(violations);
    }

    [Fact]
    public void Evaluate_passes_a_mustMatchGolden_rule_when_the_field_never_appears_in_the_diff_list()
    {
        var ruleSet = ComparisonRuleSet.FromJson("""
            { "rules": [ { "field": "page.orientation", "type": "mustMatchGolden", "description": "orientation must match" } ]}
            """);
        var submission = JsonNode.Parse("""{"page":{"orientation":"portrait"}}""")!;

        // No diff for page.orientation at all -- exactly what ProfileComparer
        // would produce for two profiles that agree on this field.
        var violations = ruleSet.Evaluate(submission, diffs: Array.Empty<ProfileDiff>());

        Assert.Empty(violations);
    }
}
