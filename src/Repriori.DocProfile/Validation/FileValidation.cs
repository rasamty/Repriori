namespace Repriori.DocProfile.Validation;

/// <summary>Result of checking whether a file or folder can actually be used the way it's about to be.</summary>
public sealed record FileAccessCheck(bool IsOk, string? Reason)
{
    public static readonly FileAccessCheck Ok = new(true, null);
}

/// <summary>
/// Checks a file or folder is actually usable *before* attempting real work on it,
/// so callers (the CLI in particular) can show one clear sentence explaining what
/// is wrong, instead of a raw exception from deep inside a ZIP/XML reader that a
/// beginner would have no way to interpret.
/// </summary>
public static class FileValidation
{
    /// <summary>
    /// True only if the file exists and can actually be opened for reading right
    /// now. Existence alone is not enough to promise readability — the most common
    /// real-world failure here is the file being open, and locked, in Word.
    /// </summary>
    public static FileAccessCheck EnsureFileIsReadable(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return new FileAccessCheck(false, "No file path was given.");

        if (!File.Exists(path))
            return new FileAccessCheck(false, $"File not found: {path}");

        try
        {
            // FileShare.ReadWrite: we only want to prove the file is *readable*,
            // not claim exclusive access to it — another program having it open
            // for reading (or even writing) should not fail this check on its own.
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return FileAccessCheck.Ok;
        }
        catch (UnauthorizedAccessException)
        {
            return new FileAccessCheck(false, $"Permission denied reading: {path}");
        }
        catch (IOException ex)
        {
            // The single most common cause in practice: the file is currently open
            // in Word, which locks it exclusively against other writers.
            return new FileAccessCheck(false, $"File is in use by another program (is it open in Word?): {path} ({ex.Message})");
        }
    }

    /// <summary>
    /// True only if a real, throwaway file can actually be written into the given
    /// directory right now. If the directory does not exist yet, this creates it —
    /// callers writing output for the first time should not have to create the
    /// folder themselves before checking whether they're even allowed to.
    /// </summary>
    public static FileAccessCheck EnsureDirectoryIsWritable(string directoryPath)
    {
        if (string.IsNullOrWhiteSpace(directoryPath))
            return new FileAccessCheck(false, "No output folder was given.");

        try
        {
            Directory.CreateDirectory(directoryPath);
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or IOException)
        {
            return new FileAccessCheck(false, $"Cannot create output folder: {directoryPath} ({ex.Message})");
        }

        var probePath = Path.Combine(directoryPath, $".write-check-{Guid.NewGuid():N}.tmp");
        try
        {
            File.WriteAllText(probePath, string.Empty);
            return FileAccessCheck.Ok;
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or IOException)
        {
            return new FileAccessCheck(false, $"Folder is not writable: {directoryPath} ({ex.Message})");
        }
        finally
        {
            // Best-effort cleanup of the probe file — if this one delete fails for
            // some transient reason, it is a harmless leftover .tmp file, not a
            // reason to change the answer we already have.
            try { File.Delete(probePath); } catch { /* ignored on purpose, see above */ }
        }
    }
}
