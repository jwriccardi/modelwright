using System;

namespace Modelwright.Core.Trace;

/// <summary>
/// Compares where a formula says an external workbook is with where the open workbook of that name came from. Excel
/// cannot open two workbooks of the same name, so a reference to <c>C:\Deals\[Book.xlsx]Inputs!A1</c> while a
/// <c>Book.xlsx</c> from another folder is open still points at the closed file; Trace In must not show the open one's
/// cells for it.
/// </summary>
public static class WorkbookPaths
{
    /// <summary>
    /// Whether <paramref name="writtenFolder"/> (the folder or URL a formula wrote before the workbook's name, as in
    /// <see cref="FormulaReference.WorkbookPath"/>) is the folder of the open workbook <paramref name="openFullName"/>
    /// (<c>Workbook.FullName</c>; <paramref name="openName"/> is its <c>Name</c>). Compared ignoring case, trailing
    /// separators and <c>\</c> versus <c>/</c>; a URL also ignoring percent-encoding (<c>%20</c> is a space).
    /// </summary>
    /// <returns>
    /// True for the same folder; false for a different one; null when it cannot be told: either side is empty (an
    /// unsaved workbook has no folder), or one is a URL and the other a local path (a OneDrive or SharePoint file can be
    /// open from its URL while the formula has its synced local path, or the other way round).
    /// </returns>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public static bool? SameFolder(string writtenFolder, string openFullName, string openName)
    {
        if (writtenFolder is null)
        {
            throw new ArgumentNullException(nameof(writtenFolder));
        }

        if (openFullName is null)
        {
            throw new ArgumentNullException(nameof(openFullName));
        }

        if (openName is null)
        {
            throw new ArgumentNullException(nameof(openName));
        }

        var openFolder = FolderOf(openFullName, openName);
        if (writtenFolder.Trim().Length == 0 || openFolder.Length == 0)
        {
            return null;
        }

        var writtenIsUrl = IsUrl(writtenFolder);
        if (writtenIsUrl != IsUrl(openFolder))
        {
            return null;
        }

        return string.Equals(Normalize(writtenFolder, writtenIsUrl), Normalize(openFolder, writtenIsUrl), StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The folder part of a workbook's <c>FullName</c> (<c>C:\dir\</c> of <c>C:\dir\Book.xlsx</c>), or empty for an
    /// unsaved workbook (whose <c>FullName</c> is just its name).
    /// </summary>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public static string FolderOf(string fullName, string name)
    {
        if (fullName is null)
        {
            throw new ArgumentNullException(nameof(fullName));
        }

        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        if (name.Length > 0 && fullName.Length > name.Length &&
            fullName.EndsWith(name, StringComparison.OrdinalIgnoreCase) &&
            IsSeparator(fullName[fullName.Length - name.Length - 1]))
        {
            return fullName.Substring(0, fullName.Length - name.Length);
        }

        var last = fullName.LastIndexOfAny(new[] { '\\', '/' });
        return last < 0 ? string.Empty : fullName.Substring(0, last + 1);
    }

    private static bool IsUrl(string path)
    {
        var trimmed = path.Trim();
        return trimmed.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSeparator(char c) => c == '\\' || c == '/';

    private static string Normalize(string folder, bool isUrl)
    {
        var text = folder.Trim();
        if (isUrl)
        {
            try
            {
                text = Uri.UnescapeDataString(text);
            }
            catch (UriFormatException)
            {
                // Not valid percent-encoding: compare as written.
            }

            return text.Replace('\\', '/').TrimEnd('/');
        }

        return text.Replace('/', '\\').TrimEnd('\\');
    }
}
