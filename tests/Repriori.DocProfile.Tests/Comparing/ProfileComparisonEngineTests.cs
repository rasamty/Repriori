using Repriori.DocProfile.Comparing;
using Xunit;

namespace Repriori.DocProfile.Tests.Comparing;

public class ProfileComparisonEngineTests
{
    [Fact]
    public void Compare_passes_when_golden_and_submission_are_identical()
    {
        const string profile = """{"revisions":{"hasUnacceptedTrackedChanges":false},"comments":{"unresolved":0},"page":{"orientation":"portrait","paperHint":"A4"},"fonts":{"bodyTypeface":"Calibri","bodySizePt":11,"heading1Typeface":"Calibri"}}""";

        var result = ProfileComparisonEngine.Compare(profile, profile);

        Assert.True(result.Passed);
        Assert.Empty(result.Violations);
    }

    [Fact]
    public void Compare_fails_when_a_default_rule_is_violated()
    {
        const string golden = """{"revisions":{"hasUnacceptedTrackedChanges":false},"comments":{"unresolved":0},"page":{"orientation":"portrait","paperHint":"A4"},"fonts":{"bodyTypeface":"Calibri","bodySizePt":11,"heading1Typeface":"Calibri"}}""";
        const string submission = """{"revisions":{"hasUnacceptedTrackedChanges":true},"comments":{"unresolved":0},"page":{"orientation":"portrait","paperHint":"A4"},"fonts":{"bodyTypeface":"Calibri","bodySizePt":11,"heading1Typeface":"Calibri"}}""";

        var result = ProfileComparisonEngine.Compare(golden, submission);

        Assert.False(result.Passed);
        Assert.Contains(result.Violations, v => v.Rule.Field == "revisions.hasUnacceptedTrackedChanges");
    }

    [Fact]
    public void Compare_reports_diffs_that_are_not_rule_violations_without_failing()
    {
        // extractorVersion differing is a real, reportable diff -- but no default
        // rule checks it, so it must show up in Diffs without affecting Passed.
        const string golden = """{"extractorVersion":"1.0.0","revisions":{"hasUnacceptedTrackedChanges":false},"comments":{"unresolved":0},"page":{"orientation":"portrait","paperHint":"A4"},"fonts":{"bodyTypeface":"Calibri","bodySizePt":11,"heading1Typeface":"Calibri"}}""";
        const string submission = """{"extractorVersion":"1.1.0","revisions":{"hasUnacceptedTrackedChanges":false},"comments":{"unresolved":0},"page":{"orientation":"portrait","paperHint":"A4"},"fonts":{"bodyTypeface":"Calibri","bodySizePt":11,"heading1Typeface":"Calibri"}}""";

        var result = ProfileComparisonEngine.Compare(golden, submission);

        Assert.True(result.Passed);
        Assert.Contains(result.Diffs, d => d.FieldPath == "extractorVersion");
    }

    [Fact]
    public void Compare_accepts_a_custom_rule_set_instead_of_the_default()
    {
        var customRules = ComparisonRuleSet.FromJson("""
            { "rules": [ { "field": "comments.unresolved", "type": "mustEqual", "value": 5, "description": "exactly five, for this test" } ]}
            """);
        const string profile = """{"comments":{"unresolved":5}}""";

        var result = ProfileComparisonEngine.Compare(profile, profile, customRules);

        Assert.True(result.Passed);
    }
}
