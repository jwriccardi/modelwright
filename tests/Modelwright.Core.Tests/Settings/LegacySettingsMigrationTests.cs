using System;
using Modelwright.Core.Settings;
using Xunit;

namespace Modelwright.Core.Tests.Settings;

public class LegacySettingsMigrationTests
{
    [Fact]
    public void A_first_run_after_the_rename_copies_the_settings_and_the_ui_state()
    {
        var copy = LegacySettingsMigration.FilesToCopy(
            new[] { "settings.json", "ui-state.json", "settings.json.20261001-120000.bak" },
            Array.Empty<string>());

        Assert.Equal(new[] { "settings.json", "ui-state.json" }, copy);
    }

    [Fact]
    public void Without_a_legacy_ui_state_only_the_settings_are_copied()
    {
        var copy = LegacySettingsMigration.FilesToCopy(new[] { "settings.json" }, Array.Empty<string>());

        Assert.Equal(new[] { "settings.json" }, copy);
    }

    [Fact]
    public void Nothing_is_copied_without_legacy_settings()
    {
        Assert.Empty(LegacySettingsMigration.FilesToCopy(Array.Empty<string>(), Array.Empty<string>()));
        Assert.Empty(LegacySettingsMigration.FilesToCopy(new[] { "ui-state.json" }, Array.Empty<string>()));
    }

    [Theory]
    [InlineData("settings.json")]
    [InlineData("Settings.JSON")]
    [InlineData("settings.json.20261008-090000.bak")]
    [InlineData("settings.json.20261008-090000-2.bak")]
    public void Nothing_is_copied_when_the_current_folder_has_settings_or_a_backup_of_them(string currentFile)
    {
        Assert.Empty(LegacySettingsMigration.FilesToCopy(new[] { "settings.json", "ui-state.json" }, new[] { currentFile }));
    }

    [Fact]
    public void An_existing_current_ui_state_is_not_overwritten()
    {
        var copy = LegacySettingsMigration.FilesToCopy(new[] { "settings.json", "ui-state.json" }, new[] { "ui-state.json" });

        Assert.Equal(new[] { "settings.json" }, copy);
    }

    [Fact]
    public void Other_files_do_not_block_the_copy()
    {
        var copy = LegacySettingsMigration.FilesToCopy(
            new[] { "SETTINGS.json" },
            new[] { "settings.json.tmp", "ui-state.json.tmp", "notes.txt" });

        Assert.Equal(new[] { "settings.json" }, copy);
    }

    [Fact]
    public void Null_arguments_are_rejected()
    {
        Assert.Throws<ArgumentNullException>(() => LegacySettingsMigration.FilesToCopy(null!, Array.Empty<string>()));
        Assert.Throws<ArgumentNullException>(() => LegacySettingsMigration.FilesToCopy(Array.Empty<string>(), null!));
    }
}
