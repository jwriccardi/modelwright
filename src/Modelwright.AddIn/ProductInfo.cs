using System.Linq;
using System.Reflection;

namespace Modelwright.AddIn;

/// <summary>Product name and build stamp (see the StampVersion target in the project file).</summary>
internal static class ProductInfo
{
    /// <summary>Product name. Change it here only.</summary>
    public const string Name = "Modelwright";

    /// <summary>
    /// Folder name for per-user files under <c>%APPDATA%</c> and <c>%LOCALAPPDATA%</c>: <see cref="Name"/>
    /// without spaces, e.g. <c>Modelwright</c>.
    /// </summary>
    public static string FolderName { get; } = Name.Replace(" ", string.Empty);

    private static readonly Assembly ThisAssembly = typeof(ProductInfo).Assembly;

    /// <summary>Informational version, e.g. <c>0.1.0+1a2b3c4</c> (or <c>0.1.0+local</c> outside git).</summary>
    public static string Version { get; } =
        ThisAssembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "unknown";

    /// <summary>Short git commit SHA the add-in was built from, or <c>local</c>.</summary>
    public static string Commit { get; } = Metadata("GitCommit");

    /// <summary>UTC build date, <c>yyyy-MM-dd</c>.</summary>
    public static string BuildDate { get; } = Metadata("BuildDate");

    private static string Metadata(string key) =>
        ThisAssembly.GetCustomAttributes<AssemblyMetadataAttribute>().FirstOrDefault(a => a.Key == key)?.Value ?? "unknown";
}
