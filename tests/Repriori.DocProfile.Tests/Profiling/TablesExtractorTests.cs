using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Repriori.DocProfile.Profiling;
using Repriori.DocProfile.Tests.Fixtures;
using Xunit;

namespace Repriori.DocProfile.Tests.Profiling;

public class TablesExtractorTests
{
    [Fact]
    public void Extract_reports_zero_tables_when_there_are_none()
    {
        using var docx = TestDocxBuilder.Minimal();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        TablesExtractor.Extract(word, root);

        Assert.Equal(0, root["tables"]!["count"]!.GetValue<int>());
        Assert.Empty(root["tables"]!["items"]!.AsArray());
    }

    [Fact]
    public void Extract_reads_row_count_header_flag_and_header_formatting()
    {
        using var docx = TestDocxBuilder.WithTable();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        TablesExtractor.Extract(word, root);
        var table = root["tables"]!["items"]![0]!;

        Assert.Equal(2, table["rowCount"]!.GetValue<int>());
        Assert.True(table["hasHeaderRow"]!.GetValue<bool>());
    }

    [Fact]
    public void Extract_reports_no_header_row_when_no_header_signal_is_present()
    {
        using var docx = TestDocxBuilder.WithTableNoHeader();
        using var word = WordprocessingDocument.Open(docx, isEditable: false);
        var root = new JsonObject();

        TablesExtractor.Extract(word, root);
        var table = root["tables"]!["items"]![0]!;

        Assert.False(table["hasHeaderRow"]!.GetValue<bool>());
        Assert.Null(table["styleName"]);
    }
}
