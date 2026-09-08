using System.Text.Json.Nodes;
using Repriori.DocProfile.Comparing;
using Xunit;

namespace Repriori.DocProfile.Tests.Comparing;

public class ProfileDiffTests
{
    [Fact]
    public void ToString_includes_the_field_path_and_both_values()
    {
        var diff = new ProfileDiff("page.orientation", JsonValue.Create("portrait"), JsonValue.Create("landscape"));

        var text = diff.ToString();

        Assert.Contains("page.orientation", text);
        Assert.Contains("portrait", text);
        Assert.Contains("landscape", text);
    }

    [Fact]
    public void ToString_shows_null_for_a_missing_golden_value()
    {
        var diff = new ProfileDiff("fonts.bodyTypeface", null, JsonValue.Create("Comic Sans MS"));

        Assert.Contains("golden=null", diff.ToString());
    }

    [Fact]
    public void ToString_shows_null_for_a_missing_submission_value()
    {
        var diff = new ProfileDiff("fonts.bodyTypeface", JsonValue.Create("Calibri"), null);

        Assert.Contains("submission=null", diff.ToString());
    }
}
