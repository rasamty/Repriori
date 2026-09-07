using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "lists" section. This one is intentionally the smallest and least
/// complete extractor: Word's numbering system (numbering.xml) can express a
/// bulleted list and a numbered list using the exact same &lt;w:numPr&gt; element on
/// a paragraph — telling them apart means resolving that reference all the way
/// into the numbering definitions part and reading the format code there. That is
/// a genuinely separate, more involved piece of work, so "hasBullets" is left
/// null on purpose here rather than guessed — see the "still null on purpose"
/// notes in the schema file for the project's convention on this.
/// </summary>
internal static class ListsExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var section = JsonHelpers.Obj(root, "lists");
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return;

        JsonHelpers.Set(section, "hasNumbering", body.Descendants<NumberingProperties>().Any());
        JsonHelpers.Set(section, "hasBullets", null);
    }
}
