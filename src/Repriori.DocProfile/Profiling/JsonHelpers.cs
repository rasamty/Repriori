using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Repriori.DocProfile.Profiling;

/// <summary>
/// Small, shared helpers used by every extractor in this folder. None of this is
/// business logic on its own — it exists so each extractor can stay focused on
/// "what does this section of the profile mean" instead of repeating the same
/// unit-conversion and null-safety code twelve times over.
/// </summary>
internal static class JsonHelpers
{
    /// <summary>
    /// Gets (or creates) a nested JSON object by name. The profile template
    /// (WordFormatComparer.json) already defines every section up front, so in
    /// practice this always finds an existing object — the "create if missing"
    /// branch is just a safety net if a section is ever renamed.
    /// </summary>
    public static JsonObject Obj(JsonObject parent, string name)
    {
        if (parent[name] is JsonObject existing) return existing;
        var created = new JsonObject();
        parent[name] = created;
        return created;
    }

    /// <summary>
    /// Writes a value into the JSON tree, normalising it on the way in so the
    /// output is predictable no matter which extractor produced it:
    /// - an empty/whitespace-only string becomes JSON null (never an empty string)
    /// - doubles are rounded to 2 decimal places (raw twip-to-mm maths produces
    ///   long floating-point tails that are not meaningful at this precision)
    /// This is the one place "what does null mean" is decided, instead of every
    /// extractor inventing its own rule.
    /// </summary>
    public static void Set(JsonObject o, string name, object? value)
    {
        o[name] = value switch
        {
            null => null,
            string s when string.IsNullOrWhiteSpace(s) => null,
            string s => s,
            bool b => b,
            int i => i,
            double d => Math.Round(d, 2),
            _ => JsonValue.Create(Convert.ToString(value, CultureInfo.InvariantCulture)),
        };
    }

    public static JsonArray ToArray(IEnumerable<string> items)
    {
        var array = new JsonArray();
        foreach (var item in items) array.Add(item);
        return array;
    }

    /// <summary>
    /// Reads the plain text out of a header/footer part (a separate small XML
    /// document inside the .docx package) and trims it to a short sample.
    /// Wrapped in try/catch deliberately: a header/footer part can theoretically
    /// contain malformed XML in a hand-edited file, and a sampling helper like
    /// this should degrade to "no sample" rather than take the whole extraction
    /// down with it.
    /// </summary>
    public static string? PartText(OpenXmlPart part)
    {
        try
        {
            using var reader = new StreamReader(part.GetStream());
            var xml = System.Xml.Linq.XDocument.Parse(reader.ReadToEnd());
            var text = string.Join(" ", xml.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value));
            return Clip(text);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>True if a header/footer part contains a PAGE field (an automatic page-number).</summary>
    public static bool HasPageField(OpenXmlPart part)
    {
        using var reader = new StreamReader(part.GetStream());
        return reader.ReadToEnd().Contains("PAGE", StringComparison.OrdinalIgnoreCase);
    }

    public static string InnerText(Paragraph paragraph) =>
        string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));

    /// <summary>
    /// Word style ids and style names disagree on whether "Heading 1" has a space
    /// in it depending on which tool last saved the file. Comparing with spaces
    /// stripped, case-insensitively, is what makes style matching reliable across
    /// files from different Word versions.
    /// </summary>
    public static bool IsStyle(string style, params string[] names) =>
        names.Any(n => string.Equals((style ?? "").Replace(" ", ""), n.Replace(" ", ""), StringComparison.OrdinalIgnoreCase));

    public static bool ContainsHeading(JsonArray headings, string token) =>
        headings.Select(n => n!.GetValue<string>()).Any(t => t.Contains(token, StringComparison.OrdinalIgnoreCase));

    /// <summary>Collapses whitespace and caps sample text at 160 characters, so long paragraphs don't bloat the profile.</summary>
    public static string? Clip(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = Regex.Replace(s, @"\s+", " ").Trim();
        return s.Length <= 160 ? s : s[..160];
    }

    // Word stores most measurements in "twips" (twentieths of a point) or half-points.
    // These three converters exist purely so the rest of the code can talk in the units
    // a human actually reads a profile in: millimetres and points.
    //
    // Takes a plain int?, not the OpenXml wrapper types (Int16Value, UInt32Value)
    // directly — callers pass "someValue?.Value" to unwrap them first. An earlier
    // version of this method took `object?` and used Convert.ToInt32 to handle
    // whatever wrapper type showed up, on the assumption that they all implement
    // IConvertible the way a plain boxed int does. They don't, in this version of
    // the OpenXml SDK — a real-file smoke test (not just a build) is what caught
    // this, which is exactly why one is worth doing before calling code "done."
    public static double? TwipToMm(int? twips)
    {
        if (twips is null) return null;
        return Math.Round(twips.Value * 25.4 / 1440.0, 2);
    }

    public static double? TwipToPt(string? twips) =>
        int.TryParse(twips, out var n) ? Math.Round(n / 20.0, 2) : null;

    public static double? HalfPointToPt(string? half) =>
        int.TryParse(half, out var n) ? n / 2.0 : null;

    /// <summary>Guesses a common paper name from its measured size, within a small tolerance for rounding.</summary>
    public static string? PaperHint(double widthMm, double heightMm) =>
        Math.Abs(widthMm - 210) < 3 && Math.Abs(heightMm - 297) < 3 ? "A4"
        : Math.Abs(widthMm - 215.9) < 3 && Math.Abs(heightMm - 279.4) < 3 ? "Letter"
        : null;

    public static string? LineSpacingName(SpacingBetweenLines? spacing) =>
        spacing?.Line is null || !int.TryParse(spacing.Line.Value, out var line) ? null
        : line switch { 240 => "single", 276 => "1.15", 360 => "1.5", 480 => "double", _ => null };

    /// <summary>True for any colour that is not black, white, or Word's "auto" — used to spot deliberate colour choices.</summary>
    public static bool IsNonBw(string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex)) return false;
        hex = hex.Trim('#');
        if (hex.Equals("auto", StringComparison.OrdinalIgnoreCase)) return false;
        return hex.Length == 6
            && !hex.Equals("000000", StringComparison.OrdinalIgnoreCase)
            && !hex.Equals("FFFFFF", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Copies a style's typeface/size/colour into the flatter "fonts" section, so callers who only care about fonts don't have to know the styles structure.</summary>
    public static void CopyFont(JsonObject styles, string styleKey, JsonObject fonts, string faceField, string sizeField, string colorField)
    {
        var slot = styles[styleKey] as JsonObject;
        Set(fonts, faceField, slot?["typeface"]?.GetValue<string>());
        Set(fonts, sizeField, slot?["sizePt"] is JsonValue v && v.TryGetValue<double>(out var n) ? n : null);
        Set(fonts, colorField, slot?["colorHex"]?.GetValue<string>());
    }
}
