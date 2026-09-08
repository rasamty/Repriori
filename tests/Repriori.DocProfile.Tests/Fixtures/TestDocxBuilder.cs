using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

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

    /// <summary>
    /// A default header, a first-page header, and an even-page header — each with
    /// distinct text — plus matching footers (the default footer contains a real PAGE
    /// field via SimpleField, not just literal text that happens to say "page"), a
    /// two-column section, and TitlePage/EvenAndOddHeaders both set. Deliberately one
    /// large, realistic fixture rather than several tiny ones: every one of these
    /// features lives together on the same SectionProperties in a real document, so
    /// testing them together is more honest than pretending they're independent.
    /// </summary>
    public static MemoryStream WithHeadersFootersAndPageFeatures()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();

            var defaultHeader = mainPart.AddNewPart<HeaderPart>();
            defaultHeader.Header = new Header(new Paragraph(new Run(new Text("Default header"))));
            defaultHeader.Header.Save();

            var firstHeader = mainPart.AddNewPart<HeaderPart>();
            firstHeader.Header = new Header(new Paragraph(new Run(new Text("First-page header"))));
            firstHeader.Header.Save();

            var evenHeader = mainPart.AddNewPart<HeaderPart>();
            evenHeader.Header = new Header(new Paragraph(new Run(new Text("Even-page header"))));
            evenHeader.Header.Save();

            var defaultFooter = mainPart.AddNewPart<FooterPart>();
            defaultFooter.Footer = new Footer(new Paragraph(new Run(new SimpleField { Instruction = "PAGE" })));
            defaultFooter.Footer.Save();

            var firstFooter = mainPart.AddNewPart<FooterPart>();
            firstFooter.Footer = new Footer(new Paragraph(new Run(new Text("First-page footer"))));
            firstFooter.Footer.Save();

            var evenFooter = mainPart.AddNewPart<FooterPart>();
            evenFooter.Footer = new Footer(new Paragraph(new Run(new Text("Even-page footer"))));
            evenFooter.Footer.Save();

            var sectionProperties = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(defaultHeader) },
                new HeaderReference { Type = HeaderFooterValues.First, Id = mainPart.GetIdOfPart(firstHeader) },
                new HeaderReference { Type = HeaderFooterValues.Even, Id = mainPart.GetIdOfPart(evenHeader) },
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(defaultFooter) },
                new FooterReference { Type = HeaderFooterValues.First, Id = mainPart.GetIdOfPart(firstFooter) },
                new FooterReference { Type = HeaderFooterValues.Even, Id = mainPart.GetIdOfPart(evenFooter) },
                new TitlePage(),
                new Columns { ColumnCount = 2 },
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 });

            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Body text."))),
                sectionProperties));

            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(new EvenAndOddHeaders());
            settingsPart.Settings.Save();

            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// A header containing one image-like drawing, whose own declared name is the given
    /// string — the one signal HeadersFootersLogoExtractor's likelyCompanyLogo heuristic
    /// actually reads. No real embedded picture bytes are needed: the extractor only
    /// ever reads the drawing's own wp:docPr name/description metadata, never the image
    /// content itself, so a bare DW.DocProperties element is a faithful, minimal fixture.
    /// </summary>
    public static MemoryStream WithHeaderImage(string name)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(new Run(
                new Drawing(new DW.Inline(new DW.DocProperties { Id = 1, Name = name })))));
            headerPart.Header.Save();

            var sectionProperties = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) });

            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Body text."))), sectionProperties));
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>One inline image in the body (with or without alt text set), and optionally one Caption-styled paragraph.</summary>
    public static MemoryStream WithFigure(string? altText, string? captionText)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            var docProperties = new DW.DocProperties { Id = 1, Name = "Picture 1" };
            if (altText is not null) docProperties.Description = altText;

            var body = new Body(new Paragraph(new Run(new Drawing(new DW.Inline(docProperties)))));
            if (captionText is not null)
            {
                body.Append(new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
                    new Run(new Text(captionText))));
            }

            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// A document with WriteProtection set — a genuinely different Word feature from
    /// DocumentProtection (see WithReadOnlyProtection above): this is the "recommend
    /// read-only" checkbox shown when opening the file, not enforced editing
    /// restrictions.
    /// </summary>
    public static MemoryStream WithWriteProtection(bool recommended)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Body text.")))));

            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(
                new WriteProtection { Recommended = recommended ? OnOffValue.FromBoolean(true) : null });
            settingsPart.Settings.Save();

            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>A document whose Settings declare a Word compatibility-mode setting.</summary>
    public static MemoryStream WithCompatibilityMode(string value)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Body text.")))));

            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(
                new Compatibility(new CompatibilitySetting
                {
                    Name = new EnumValue<CompatSettingNameValues>(CompatSettingNameValues.CompatibilityMode),
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = value,
                }));
            settingsPart.Settings.Save();

            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>A document whose body run declares an explicit proofing language.</summary>
    public static MemoryStream WithLanguage(string langTag)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            var run = new Run(new RunProperties(new Languages { Val = langTag }), new Text("Body text."));
            mainPart.Document = new Document(new Body(new Paragraph(run)));
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>A Heading 1 style referenced by its display name with a space ("Heading 1"), not its bare id — the FillStyle fallback branch.</summary>
    public static MemoryStream WithDisplayNameStyledHeading()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(HeadingParagraph("MyCustomHeading1", "Purpose")));

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new Style(
                new StyleName { Val = "Heading 1" },
                new StyleRunProperties(new RunFonts { Ascii = "Georgia" }, new Bold(), new Color { Val = "2E74B5" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "MyCustomHeading1",
            });
            stylesPart.Styles.Save();

            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>Two distinct typefaces and two distinct non-black/white colours used directly on runs, not via any style.</summary>
    public static MemoryStream WithMultipleTypefacesAndColors()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            var run1 = new Run(new RunProperties(new RunFonts { Ascii = "Georgia" }, new Color { Val = "C00000" }), new Text("Red Georgia text."));
            var run2 = new Run(new RunProperties(new RunFonts { Ascii = "Verdana" }, new Color { Val = "1F4E5F" }), new Text("Blue Verdana text."));
            mainPart.Document = new Document(new Body(new Paragraph(run1), new Paragraph(run2)));
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>A Normal-style body paragraph with an explicit named line-spacing value, centred, with space before/after.</summary>
    public static MemoryStream WithLineSpacing(int lineValue)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            var paragraphProperties = new ParagraphProperties(
                new SpacingBetweenLines { Line = lineValue.ToString(), Before = "240", After = "120" },
                new Justification { Val = JustificationValues.Center });
            mainPart.Document = new Document(new Body(new Paragraph(paragraphProperties, new Run(new Text("Body text.")))));
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>A table with no header-row signal at all: no TableHeader flag, no TableLook first-row flag, and no TableStyle.</summary>
    public static MemoryStream WithTableNoHeader()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = doc.AddMainDocumentPart();
            var row1 = new TableRow(new TableCell(new Paragraph(new Run(new Text("A")))));
            var row2 = new TableRow(new TableCell(new Paragraph(new Run(new Text("B")))));
            mainPart.Document = new Document(new Body(new Table(row1, row2)));
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
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
