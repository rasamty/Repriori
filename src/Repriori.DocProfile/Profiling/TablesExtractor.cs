using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "tables" section: one entry per table in the document, sampling its
/// header-row and body-row formatting. This deliberately samples rather than
/// inspecting every cell — a large document can have tables with hundreds of
/// rows, and "what font does the header use" is answered just as well by the
/// first header run as by checking all of them.
/// </summary>
internal static class TablesExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var section = JsonHelpers.Obj(root, "tables");
        var tables = word.MainDocumentPart?.Document?.Body?.Elements<Table>().ToList() ?? new List<Table>();
        JsonHelpers.Set(section, "count", tables.Count);

        var items = new JsonArray();
        var index = 0;
        foreach (var table in tables)
        {
            index++;
            var item = new JsonObject { ["index"] = index };

            JsonHelpers.Set(item, "styleName", table.GetFirstChild<TableProperties>()?.TableStyle?.Val?.Value);

            var rows = table.Elements<TableRow>().ToList();
            JsonHelpers.Set(item, "rowCount", rows.Count);

            var headerRow = rows.FirstOrDefault();
            // A header row can be marked two different ways in the XML: an explicit
            // <w:tblHeader/> on the row itself, or the table-wide "first row is a
            // header" look flag. Either one counts.
            JsonHelpers.Set(item, "hasHeaderRow",
                headerRow?.TableRowProperties?.GetFirstChild<TableHeader>() is not null
                || table.GetFirstChild<TableProperties>()?.TableLook?.FirstRow?.Value == true);

            var headerRun = headerRow?.Descendants<Run>().FirstOrDefault();
            JsonHelpers.Set(item, "headerTypeface", headerRun?.RunProperties?.RunFonts?.Ascii?.Value);
            JsonHelpers.Set(item, "headerSizePt", JsonHelpers.HalfPointToPt(headerRun?.RunProperties?.FontSize?.Val));
            JsonHelpers.Set(item, "headerColorHex", headerRun?.RunProperties?.Color?.Val?.Value);
            JsonHelpers.Set(item, "headerFillHex", headerRow?.Descendants<Shading>().FirstOrDefault()?.Fill?.Value);

            var bodyRun = rows.Skip(1).SelectMany(r => r.Descendants<Run>()).FirstOrDefault();
            JsonHelpers.Set(item, "contentTypeface", bodyRun?.RunProperties?.RunFonts?.Ascii?.Value);
            JsonHelpers.Set(item, "contentSizePt", JsonHelpers.HalfPointToPt(bodyRun?.RunProperties?.FontSize?.Val));
            JsonHelpers.Set(item, "contentColorHex", bodyRun?.RunProperties?.Color?.Val?.Value);

            JsonHelpers.Set(item, "gridColorHex", table.GetFirstChild<TableProperties>()?.TableBorders?.TopBorder?.Color?.Value);

            items.Add(item);
        }

        section["items"] = items;
    }
}
