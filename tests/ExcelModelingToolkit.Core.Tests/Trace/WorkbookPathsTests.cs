using System;
using ExcelModelingToolkit.Core.Trace;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Trace;

public class WorkbookPathsTests
{
    [Theory]
    [InlineData(@"C:\Deals\", @"C:\Deals\Book.xlsx")]
    [InlineData(@"c:\deals", @"C:\Deals\Book.xlsx")]
    [InlineData(@"C:/Deals/", @"C:\Deals\Book.xlsx")]
    [InlineData(@"\\server\share\Deals\", @"\\SERVER\share\Deals\Book.xlsx")]
    [InlineData("https://contoso.sharepoint.com/sites/x/Shared Documents/", "https://contoso.sharepoint.com/sites/x/Shared%20Documents/Book.xlsx")]
    [InlineData("https://d.docs.live.net/abc/Docs", "https://d.docs.live.net/abc/Docs/Book.xlsx")]
    public void The_same_folder_written_differently_is_the_same(string written, string fullName)
    {
        Assert.True(WorkbookPaths.SameFolder(written, fullName, "Book.xlsx"));
    }

    [Theory]
    [InlineData(@"C:\Deals\", @"C:\Other\Book.xlsx")]
    [InlineData(@"C:\Deals\2024\", @"C:\Deals\Book.xlsx")]
    [InlineData("https://d.docs.live.net/abc/Docs/", "https://d.docs.live.net/abc/Other/Book.xlsx")]
    public void A_different_folder_is_different(string written, string fullName)
    {
        Assert.False(WorkbookPaths.SameFolder(written, fullName, "Book.xlsx"));
    }

    [Theory]
    [InlineData(@"C:\Users\me\OneDrive\Docs\", "https://d.docs.live.net/abc/Docs/Book.xlsx")]
    [InlineData("https://d.docs.live.net/abc/Docs/", @"C:\Users\me\OneDrive\Docs\Book.xlsx")]
    [InlineData(@"C:\Deals\", "Book.xlsx")]
    [InlineData("  ", @"C:\Deals\Book.xlsx")]
    public void A_url_against_a_local_path_or_an_unsaved_workbook_cannot_be_told(string written, string fullName)
    {
        Assert.Null(WorkbookPaths.SameFolder(written, fullName, "Book.xlsx"));
    }

    [Theory]
    [InlineData(@"C:\Deals\Book.xlsx", "Book.xlsx", @"C:\Deals\")]
    [InlineData("https://x.com/a/Book.xlsx", "Book.xlsx", "https://x.com/a/")]
    [InlineData("Book1", "Book1", "")]
    [InlineData(@"C:\Deals\Renamed.xlsx", "Book.xlsx", @"C:\Deals\")]
    public void The_folder_of_a_full_name_is_everything_before_the_name(string fullName, string name, string folder)
    {
        Assert.Equal(folder, WorkbookPaths.FolderOf(fullName, name));
    }

    [Fact]
    public void Arguments_are_checked()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookPaths.SameFolder(null!, "a", "a"));
        Assert.Throws<ArgumentNullException>(() => WorkbookPaths.SameFolder("a", null!, "a"));
        Assert.Throws<ArgumentNullException>(() => WorkbookPaths.SameFolder("a", "a", null!));
        Assert.Throws<ArgumentNullException>(() => WorkbookPaths.FolderOf(null!, "a"));
        Assert.Throws<ArgumentNullException>(() => WorkbookPaths.FolderOf("a", null!));
    }
}
