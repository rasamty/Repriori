using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Fills the "document" section: the file's own metadata (title, author, language,
/// protection state) — the kind of thing you'd see under File &gt; Info in Word,
/// rather than anything about the visible content.
/// </summary>
internal static class DocumentPropertiesExtractor
{
    public static void Extract(WordprocessingDocument word, JsonObject root)
    {
        var section = JsonHelpers.Obj(root, "document");

        // "PackageProperties" is the .docx file's own metadata block — the same
        // fields Windows Explorer shows under a file's Properties > Details tab.
        var core = word.PackageProperties;
        JsonHelpers.Set(section, "coreTitle", core.Title);
        JsonHelpers.Set(section, "coreSubject", core.Subject);
        JsonHelpers.Set(section, "coreCreator", core.Creator);
        JsonHelpers.Set(section, "coreLastModifiedBy", core.LastModifiedBy);
        JsonHelpers.Set(section, "coreCategory", core.Category);
        JsonHelpers.Set(section, "coreRevision", core.Revision);

        JsonHelpers.Set(section, "applicationName", word.ExtendedFilePropertiesPart?.Properties?.Application?.Text);

        // A document can declare a different proofing language per run of text
        // (e.g. a quoted phrase in French inside an English document). Taking the
        // first non-empty one gives a reasonable "what language is this document"
        // answer without trying to model per-run language mixing.
        JsonHelpers.Set(section, "language",
            word.MainDocumentPart?.Document?.Body?.Descendants<Languages>()
                .Select(l => l.Val?.Value)
                .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)));

        var settings = word.MainDocumentPart?.DocumentSettingsPart?.Settings;

        var protection = settings?.GetFirstChild<DocumentProtection>();
        JsonHelpers.Set(section, "isProtected", protection is not null);
        // Real-file bug, found by running this against an actual protected
        // document rather than a synthetic fixture: protection.Edit.Value's own
        // ToString() prints "DocumentProtectionValues { }" in this SDK version —
        // not the readable value ("readOnly", "forms", etc.). .InnerText reads
        // the real underlying XML attribute text directly and is correct.
        JsonHelpers.Set(section, "protectionType", protection?.Edit?.InnerText);

        var writeProtection = settings?.GetFirstChild<WriteProtection>();
        JsonHelpers.Set(section, "readOnlyRecommended", writeProtection?.Recommended?.Value);
        JsonHelpers.Set(section, "writeReservation", writeProtection is not null);

        JsonHelpers.Set(section, "compatibilityMode",
            settings?.GetFirstChild<Compatibility>()?.GetFirstChild<CompatibilitySetting>()?.Val?.Value);
    }
}
