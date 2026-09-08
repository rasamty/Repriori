using Repriori.DocProfile.Comparing;
using Xunit;

namespace Repriori.DocProfile.Tests.Comparing;

public class RuleViolationTests
{
    [Fact]
    public void ToString_includes_the_field_the_description_and_the_detail()
    {
        var rule = new ComparisonRule("comments.unresolved", ComparisonRuleType.MustEqual, null, "all comments resolved");
        var violation = new RuleViolation(rule, "expected 0, got 3");

        var text = violation.ToString();

        Assert.Contains("comments.unresolved", text);
        Assert.Contains("all comments resolved", text);
        Assert.Contains("expected 0, got 3", text);
    }
}
