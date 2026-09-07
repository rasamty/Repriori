using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Tests.Fixtures;

/// <summary>
/// Builds small, in-memory .docx files for tests, using the same OpenXml SDK the
/// product code reads with — no binary fixture files are checked into the repo at
/// all. Two independent reasons for that: the .gitignore now excludes .docx
/// everywhere, so a checked-in binary fixture would silently never make it into
/// Git; and even before that rule existed, a fixture built here in code is
/// reviewable in a diff and self-documents exactly what it contains, where a
/// binary blob would just be an opaque file nobody can inspect without opening it.
///
/// Every method here returns a ready-to-read MemoryStream positioned at the start.
/// </summary>
internal static class TestDocxBuilder
{
    /// <summary>The smallest possible valid document — one plain paragraph, nothing else.</summary>
    public static MemoryStream Minimal(string text = "Hello, world.")
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text(text)))
            ));
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// One paragraph in each of the Heading1-4 styles, with distinct text per
    /// level — for proving each level lands in its own separate array.
    /// </summary>
    public static MemoryStream WithHeadings()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                HeadingParagraph("Heading1", "Purpose"),
                HeadingParagraph("Heading2", "Scope of this document"),
                HeadingParagraph("Heading3", "References"),
                HeadingParagraph("Heading4", "Revision History"),
                new Paragraph(new Run(new Text("Ordinary body text that is not a heading.")))
            ));
            AddStylePart(mainPart);
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Purpose/Scope/References/Revision History as four separate Heading1
    /// paragraphs — matching how these conventionally appear as top-level
    /// sections in a real SaMD document. This is what actually exercises
    /// StructureAndTocExtractor's keyword-matching flags: they only ever check
    /// sectionOrder, which the extractor deliberately populates from Heading1
    /// text alone (see its own comment) — a Heading3 called "References" is not
    /// detected, by design, which is exactly what <see cref="WithHeadings"/>
    /// above demonstrates instead.
    /// </summary>
    public static MemoryStream WithH1Keywords()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                HeadingParagraph("Heading1", "Purpose"),
                HeadingParagraph("Heading1", "Scope"),
                HeadingParagraph("Heading1", "References"),
                HeadingParagraph("Heading1", "Revision History")
            ));
            AddStylePart(mainPart);
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>One table: a bold header row ("Name") and one body row ("Value").</summary>
    public static MemoryStream WithTable()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();

            var headerRun = new Run(new RunProperties(new Bold()), new Text("Name"));
            var headerRow = new TableRow(
                new TableRowProperties(new TableHeader()),
                new TableCell(new Paragraph(headerRun)));

            var bodyRow = new TableRow(new TableCell(new Paragraph(new Run(new Text("Value")))));

            var table = new Table(
                new TableProperties(new TableStyle { Val = "TableGrid" }),
                headerRow,
                bodyRow);

            mainPart.Document = new Document(new Body(table));
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>One run inserted and one run deleted via real tracked-changes markup.</summary>
    public static MemoryStream WithTrackedChanges()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            var inserted = new InsertedRun { Author = "Reviewer", Id = "1" };
            inserted.Append(new Run(new Text("added text")));
            var deleted = new DeletedRun { Author = "Reviewer", Id = "2" };
            deleted.Append(new Run(new DeletedText("removed text")));

            mainPart.Document = new Document(new Body(new Paragraph(inserted, deleted)));
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// One comment attached to a run. When <paramref name="resolved"/> is true, a
    /// commentsEx part is added marking that comment thread "done" — this is the
    /// exact two-part relationship CommentsExtractor's XML-doc comment explains.
    /// </summary>
    public static MemoryStream WithComment(bool resolved)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            const string paraId = "1A2B3C4D";

            var commentedParagraph = new Paragraph(
                new CommentRangeStart { Id = "0" },
                new Run(new Text("commented text")),
                new CommentRangeEnd { Id = "0" },
                new Run(new CommentReference { Id = "0" }));

            mainPart.Document = new Document(new Body(commentedParagraph));

            // The paraId that ties a comment to its "resolved" flag in
            // commentsEx.xml belongs to the comment's *own* paragraph (inside
            // comments.xml) — not to the paragraph in the main document the
            // comment is anchored to. Those are two different paragraphs; an
            // earlier version of this fixture set it on the wrong one, which is
            // exactly the kind of thing a real, wired-together end-to-end test
            // catches and a compile-only check cannot.
            //
            // The w14:paraId attribute isn't exposed as a typed property on
            // ParagraphProperties in this SDK version — CommentsExtractor itself
            // reads it generically via GetAttributes(), so it's set the same
            // generic way here rather than assuming a typed property exists.
            var commentParagraphProperties = new ParagraphProperties();
            commentParagraphProperties.SetAttribute(new OpenXmlAttribute(
                "w14", "paraId", "http://schemas.microsoft.com/office/word/2010/wordml", paraId));
            var commentParagraph = new Paragraph(commentParagraphProperties, new Run(new Text("A reviewer note.")));

            var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments(
                new Comment(commentParagraph) { Id = "0", Author = "Reviewer" });
            commentsPart.Comments.Save();

            if (resolved)
            {
                var commentsExPart = mainPart.AddNewPart<WordprocessingCommentsExPart>();
                var commentEx = new DocumentFormat.OpenXml.Office2013.Word.CommentEx { ParaId = paraId, Done = true };
                commentsExPart.CommentsEx = new DocumentFormat.OpenXml.Office2013.Word.CommentsEx(commentEx);
                commentsExPart.CommentsEx.Save();
            }

            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Deliberately uneven emphasis: exactly one third of the body's characters
    /// are bold, so the test asserting boldPct can check an exact, known number
    /// (33.3) instead of merely "not zero."
    /// </summary>
    public static MemoryStream WithKnownEmphasis()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            var plainRun = new Run(new Text("AAAAAAAAAAAAAAAAAAAA"));                       // 20 plain chars
            var boldRun = new Run(new RunProperties(new Bold()), new Text("BBBBBBBBBB"));    // 10 bold chars -> 10 / 30 = 33.3%

            mainPart.Document = new Document(new Body(new Paragraph(plainRun, boldRun)));
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// A document with read-only DocumentProtection set — a direct regression
    /// fixture for the real bug found running the CLI against a real protected
    /// file in Phase 5: DocumentProtection.Edit.Value.ToString() prints
    /// "DocumentProtectionValues { }" in this SDK version rather than the actual
    /// value, because these "*Values" types are lightweight wrapper structs here,
    /// not plain enums. Nothing in the synthetic Phase 4 fixtures happened to set
    /// DocumentProtection at all, which is exactly why this slipped through until
    /// a real file was tried.
    /// </summary>
    public static MemoryStream WithReadOnlyProtection()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Body text.")))));

            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(
                new DocumentProtection { Edit = DocumentProtectionValues.ReadOnly, Enforcement = true });
            settingsPart.Settings.Save();

            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>Sets the document's core Title/Creator properties (File &gt; Info fields).</summary>
    public static MemoryStream WithCoreProperties(string title, string creator)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Body text.")))));
            mainPart.Document.Save();
            doc.PackageProperties.Title = title;
            doc.PackageProperties.Creator = creator;
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// A document with an explicit page size and margins, in twips (1/20 pt).
    /// This is a direct regression fixture for the real bug found while smoke
    /// testing Phase 3: converting these OpenXml-typed values to millimetres threw
    /// at runtime despite compiling cleanly, because the wrapper types
    /// (Int16Value/UInt32Value) don't implement IConvertible the way a plain int
    /// does in this SDK version.
    /// </summary>
    public static MemoryStream WithPageSize(int widthTwips, int heightTwips, int marginTopTwips)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            var sectionProperties = new SectionProperties(
                new PageSize { Width = (uint)widthTwips, Height = (uint)heightTwips },
                new PageMargin { Top = marginTopTwips, Bottom = 1440, Left = 1440, Right = 1440 });

            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Body text."))),
                sectionProperties));
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>A file that exists but is not a ZIP archive at all — for the "not really a .docx" error path.</summary>
    public static MemoryStream NotAZipFile()
    {
        var bytes = System.Text.Encoding.UTF8.GetBytes("this is plain text, not a zip file");
        return new MemoryStream(bytes);
    }

    private static Paragraph HeadingParagraph(string styleId, string text) =>
        new(new ParagraphProperties(new ParagraphStyleId { Val = styleId }), new Run(new Text(text)));

    /// <summary>
    /// Heading1-4 paragraphs only apply visually in Word when a matching style
    /// exists in styles.xml — but StructureAndTocExtractor only reads the
    /// paragraph's styleId, not the rendered appearance, so the extractor works
    /// correctly without this. It's added anyway to keep fixtures realistic
    /// rather than relying on that shortcut.
    /// </summary>
    private static void AddStylePart(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();
        foreach (var id in new[] { "Heading1", "Heading2", "Heading3", "Heading4" })
        {
            styles.Append(new Style(new StyleName { Val = id.Replace("Heading", "heading ") })
            {
                Type = StyleValues.Paragraph,
                StyleId = id,
            });
        }
        stylesPart.Styles = styles;
        stylesPart.Styles.Save();
    }
}
