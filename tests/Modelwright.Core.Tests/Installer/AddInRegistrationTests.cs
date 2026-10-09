using System;
using System.Collections.Generic;
using System.Linq;
using Modelwright.Installer;
using Xunit;

namespace Modelwright.Core.Tests.Installer;

/// <summary>
/// The MSI's registration rules (installer/msi/CustomActions/AddInRegistration.cs), with the same cases as
/// tests/install/install-scripts.Tests.ps1 runs against install.ps1 and uninstall.ps1, which follow the same rules.
/// A key's values are applied in memory, the way the custom action applies them to the registry.
/// </summary>
public class AddInRegistrationTests
{
    private const string Xll = @"C:\Users\u\AppData\Local\Programs\Modelwright\Modelwright64.xll";
    private static readonly string Value = AddInRegistration.OpenValue(Xll);

    // Registers (or unregisters) against a key holding these values; returns the key's values afterwards as sorted
    // "NAME=value" lines, or null when the plan said nothing needs to change.
    private static string? Run(bool register, params (string Name, string Value)[] values)
    {
        var key = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (name, value) in values)
        {
            key[name] = value;
        }

        var entries = AddInRegistration.GetOpenEntries(key.ToList());
        var plan = register ? AddInRegistration.PlanRegister(entries, Value) : AddInRegistration.PlanUnregister(entries);
        if (plan is null)
        {
            return null;
        }

        foreach (var change in AddInRegistration.Diff(entries, plan))
        {
            if (change.IsDelete)
            {
                Assert.True(key.Remove(change.Name), "deleted a value that is not there: " + change.Name);
            }
            else
            {
                key.Remove(change.Name);
                key[change.Name] = change.Value!;
            }
        }

        return Snapshot(key.Select(p => (p.Key, p.Value)).ToArray());
    }

    private static string Snapshot(params (string Name, string Value)[] values) =>
        string.Join("\n", values.Select(v => v.Name + "=" + v.Value).OrderBy(s => s, StringComparer.Ordinal));

    [Fact]
    public void Registers_as_OPEN_on_a_clean_profile()
    {
        Assert.Equal(Snapshot(("OPEN", Value)), Run(true));
        Assert.Equal("/R \"" + Xll + "\"", Value);
    }

    [Fact]
    public void Appends_after_other_add_ins()
    {
        Assert.Equal(
            Snapshot(("OPEN", "\"C:\\Other\\First.xlam\""), ("OPEN1", "\"C:\\Other\\Second.xlam\""), ("OPEN2", Value)),
            Run(true, ("OPEN", "\"C:\\Other\\First.xlam\""), ("OPEN1", "\"C:\\Other\\Second.xlam\"")));
    }

    [Fact]
    public void Replaces_an_old_ModelingToolkit_build_in_its_slot_and_keeps_other_values()
    {
        Assert.Equal(
            Snapshot(("OPEN", "\"C:\\First.xlam\""), ("OPEN1", Value), ("OPEN2", "\"C:\\Second.xlam\""), ("Options", "0")),
            Run(true,
                ("OPEN", "\"C:\\First.xlam\""),
                ("OPEN1", "/R \"C:\\dev\\bin\\Release\\net48\\publish\\ModelingToolkit64-packed.xll\""),
                ("OPEN2", "\"C:\\Second.xlam\""),
                ("Options", "0")));
    }

    [Fact]
    public void Exactly_the_right_entry_already_there_changes_nothing()
    {
        Assert.Null(Run(true, ("OPEN", "\"C:\\First.xlam\""), ("OPEN1", Value)));
        Assert.Null(Run(true, ("OPEN", Value.ToUpperInvariant())));
    }

    [Fact]
    public void Replaces_a_bare_entry_in_its_slot_drops_duplicates_keeps_look_alikes_and_renumbers()
    {
        Assert.Equal(
            Snapshot(
                ("OPEN", "\"C:\\Other\\First.xlam\""),
                ("OPEN1", Value),
                ("OPEN2", "\"C:\\Other\\Second.xlam\""),
                ("OPEN3", "/R \"C:\\Other\\MyModelwright64.xll\""),
                ("OPEN4", "\"C:\\Other\\Modelwright64.xll.bak\"")),
            Run(true,
                ("OPEN", "\"C:\\Other\\First.xlam\""),
                ("OPEN1", "/R \"Modelwright64.xll\""),
                ("OPEN2", "\"C:\\Other\\Second.xlam\""),
                ("OPEN3", "/R \"C:\\Users\\u\\AppData\\Roaming\\Microsoft\\AddIns\\Modelwright64.xll\""),
                ("OPEN4", "/R \"C:\\Other\\MyModelwright64.xll\""),
                ("OPEN5", "\"C:\\Other\\Modelwright64.xll.bak\"")));
    }

    [Fact]
    public void Replaces_a_32_bit_entry_from_an_earlier_build()
    {
        Assert.Equal(
            Snapshot(("OPEN", Value)),
            Run(true, ("OPEN", "/R \"C:\\Users\\u\\AppData\\Roaming\\Microsoft\\AddIns\\Modelwright32.xll\"")));
    }

    [Fact]
    public void Sorts_by_number_not_by_name_and_closes_gaps()
    {
        Assert.Equal(
            Snapshot(("OPEN", "\"C:\\a.xlam\""), ("OPEN1", "\"C:\\b.xlam\""), ("OPEN2", "\"C:\\c.xlam\""), ("OPEN3", Value)),
            Run(true, ("OPEN10", "\"C:\\c.xlam\""), ("OPEN", "\"C:\\a.xlam\""), ("OPEN2", "\"C:\\b.xlam\"")));
    }

    [Fact]
    public void A_lower_case_name_is_rewritten_in_upper_case()
    {
        Assert.Equal(
            Snapshot(("OPEN", "\"C:\\a.xlam\""), ("OPEN1", Value)),
            Run(true, ("OPEN", "\"C:\\a.xlam\""), ("open1", "/R \"Modelwright64.xll\"")));
    }

    [Fact]
    public void Unregister_removes_bare_and_full_path_entries_and_renumbers_the_rest()
    {
        Assert.Equal(
            Snapshot(
                ("OPEN", "\"C:\\Other\\First.xlam\""),
                ("OPEN1", "\"C:\\Other\\Second.xlam\""),
                ("OPEN2", "\"C:\\Other\\Third.xlam\""),
                ("OPEN3", "/R \"C:\\Other\\MyModelwright64.xll\"")),
            Run(false,
                ("OPEN", "\"C:\\Other\\First.xlam\""),
                ("OPEN1", "/R \"Modelwright64.xll\""),
                ("OPEN2", "\"C:\\Other\\Second.xlam\""),
                ("OPEN3", Value),
                ("OPEN4", "/R \"modelwright32.XLL\""),
                ("OPEN6", "\"C:\\Other\\Third.xlam\""),
                ("OPEN7", "/R \"C:\\Other\\MyModelwright64.xll\"")));
    }

    [Fact]
    public void Unregister_with_nothing_of_ours_changes_nothing()
    {
        Assert.Null(Run(false, ("OPEN", "\"C:\\Other\\First.xlam\""), ("OPEN2", "\"C:\\Other\\Second.xlam\"")));
        Assert.Null(Run(false));
    }

    [Theory]
    [InlineData("Modelwright64.xll", true)]
    [InlineData("MODELWRIGHT32.XLL", true)]
    [InlineData("ModelingToolkit64-packed.xll", true)]
    [InlineData(@"C:\Users\u\AppData\Local\Programs\Modelwright\Modelwright64.xll", true)]
    [InlineData(@"/R ""C:\dev\publish\Modelwright64.xll""", true)]
    [InlineData(@"C:\Other\Unrelated.xll", false)]
    [InlineData(@"C:\Other\MyModelwright64.xll", false)]
    [InlineData(@"""C:\Other\Modelwright64.xll.bak""", false)]
    [InlineData("Modelwright.xll", false)]
    public void Recognises_our_add_in_names(string value, bool ours)
    {
        Assert.Equal(ours, AddInRegistration.IsOurs(value));
    }
}
