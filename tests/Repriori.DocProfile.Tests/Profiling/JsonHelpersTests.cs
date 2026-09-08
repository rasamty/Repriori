using DocumentFormat.OpenXml.Wordprocessing;
using Repriori.DocProfile.Profiling;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

// JsonHelpers is exercised indirectly by every other extractor's tests, but several of
// its own branches (a paper size that isn't A4, an unrecognised line-spacing value, the
// "auto" colour special case) are never actually hit that way — every extractor that
// calls PaperHint, say, only ever passes it A4 dimensions. These tests call the helper
// methods directly to prove the branches those indirect call sites never reach.
public class JsonHelpersTests
{
    [Fact]
    public void PaperHint_recognises_A4()
    {
        Assert.Equal("A4", JsonHelpers.PaperHint(210, 297));
    }

    [Fact]
    public void PaperHint_recognises_Letter()
    {
        Assert.Equal("Letter", JsonHelpers.PaperHint(215.9, 279.4));
    }

    [Fact]
    public void PaperHint_returns_null_for_an_unrecognised_size()
    {
        Assert.Null(JsonHelpers.PaperHint(100, 100));
    }

    [Theory]
    [InlineData(240, "single")]
    [InlineData(276, "1.15")]
    [InlineData(360, "1.5")]
    [InlineData(480, "double")]
    public void LineSpacingName_recognises_every_named_value(int line, string expected)
    {
        var spacing = new SpacingBetweenLines { Line = line.ToString() };
        Assert.Equal(expected, JsonHelpers.LineSpacingName(spacing));
    }

    [Fact]
    public void LineSpacingName_returns_null_for_an_unrecognised_value()
    {
        var spacing = new SpacingBetweenLines { Line = "999" };
        Assert.Null(JsonHelpers.LineSpacingName(spacing));
    }

    [Fact]
    public void LineSpacingName_returns_null_when_spacing_is_null()
    {
        Assert.Null(JsonHelpers.LineSpacingName(null));
    }

    [Fact]
    public void IsNonBw_returns_false_for_auto()
    {
        Assert.False(JsonHelpers.IsNonBw("auto"));
    }

    [Theory]
    [InlineData("000000")]
    [InlineData("FFFFFF")]
    public void IsNonBw_returns_false_for_black_or_white(string hex)
    {
        Assert.False(JsonHelpers.IsNonBw(hex));
    }

    [Fact]
    public void IsNonBw_returns_true_for_a_real_colour()
    {
        Assert.True(JsonHelpers.IsNonBw("C00000"));
    }

    [Fact]
    public void IsNonBw_returns_false_for_a_missing_value()
    {
        Assert.False(JsonHelpers.IsNonBw(null));
    }

    [Fact]
    public void TwipToMm_returns_null_for_null()
    {
        Assert.Null(JsonHelpers.TwipToMm(null));
    }

    [Fact]
    public void TwipToPt_returns_null_for_an_unparseable_string()
    {
        Assert.Null(JsonHelpers.TwipToPt("not-a-number"));
    }

    [Fact]
    public void HalfPointToPt_returns_null_for_an_unparseable_string()
    {
        Assert.Null(JsonHelpers.HalfPointToPt("not-a-number"));
    }

    [Fact]
    public void HalfPointToPt_halves_a_real_value()
    {
        Assert.Equal(11.0, JsonHelpers.HalfPointToPt("22"));
    }

    [Fact]
    public void Clip_collapses_whitespace_and_truncates_to_160_characters()
    {
        var input = new string('a', 200);
        var clipped = JsonHelpers.Clip(input);

        Assert.NotNull(clipped);
        Assert.Equal(160, clipped!.Length);
    }

    [Fact]
    public void Clip_returns_null_for_whitespace_only()
    {
        Assert.Null(JsonHelpers.Clip("   "));
    }

    [Fact]
    public void IsStyle_matches_regardless_of_spacing_or_case()
    {
        Assert.True(JsonHelpers.IsStyle("heading1", "Heading 1"));
        Assert.True(JsonHelpers.IsStyle("Heading 1", "heading1"));
        Assert.False(JsonHelpers.IsStyle("Heading2", "Heading 1"));
    }
}
