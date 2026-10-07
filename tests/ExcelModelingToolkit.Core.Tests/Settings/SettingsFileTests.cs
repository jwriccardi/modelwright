using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelModelingToolkit.Core.Settings;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Settings;

public sealed class SettingsFileTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "emt-settings-tests-" + Guid.NewGuid().ToString("N"));
    private DateTime _now = new DateTime(2026, 9, 29, 10, 0, 0, DateTimeKind.Utc);

    public SettingsFileTests()
    {
        Directory.CreateDirectory(_folder);
    }

    private string SettingsPath => Path.Combine(_folder, "settings.json");

    public void Dispose()
    {
        try
        {
            Directory.Delete(_folder, recursive: true);
        }
        catch (IOException)
        {
            // A leftover temporary folder is harmless.
        }
    }

    [Fact]
    public void Save_creates_a_new_file_without_a_backup()
    {
        var file = CreateFile();

        var result = file.Save(Settings(11));

        Assert.True(result.Succeeded);
        Assert.Null(result.Problem);
        Assert.Null(result.BackupPath);
        Assert.Equal(SettingsFile.ComputeHash(File.ReadAllBytes(SettingsPath)), result.Hash);
        Assert.Equal(Settings(11).ToJson(), File.ReadAllText(SettingsPath));
        Assert.Empty(file.GetBackups());
        Assert.Empty(TemporaryFiles());
    }

    [Fact]
    public void Save_replaces_the_file_and_keeps_a_timestamped_backup()
    {
        var file = CreateFile();
        File.WriteAllText(SettingsPath, Settings(1).ToJson());

        var result = file.Save(Settings(2));

        Assert.True(result.Succeeded);
        Assert.Equal(Settings(2).ToJson(), File.ReadAllText(SettingsPath));
        Assert.Equal(Path.Combine(_folder, "settings.json.20260929-100000.bak"), result.BackupPath);
        Assert.Equal(Settings(1).ToJson(), File.ReadAllText(result.BackupPath!));
        Assert.Equal(new[] { result.BackupPath! }, file.GetBackups());
        Assert.Equal(SettingsFile.ComputeHash(File.ReadAllBytes(SettingsPath)), result.Hash);
        Assert.Empty(TemporaryFiles());
    }

    [Fact]
    public void Backups_rotate_keeping_the_newest_three()
    {
        var file = CreateFile();
        file.Save(Settings(1));
        for (var cap = 2; cap <= 6; cap++)
        {
            _now = _now.AddSeconds(1);
            Assert.True(file.Save(Settings(cap)).Succeeded);
        }

        var backups = file.GetBackups();

        Assert.Equal(SettingsFile.MaxBackups, backups.Count);
        Assert.Equal(new[] { Settings(5).ToJson(), Settings(4).ToJson(), Settings(3).ToJson() }, backups.Select(File.ReadAllText));
        Assert.Equal(
            new[] { "settings.json.20260929-100005.bak", "settings.json.20260929-100004.bak", "settings.json.20260929-100003.bak" },
            backups.Select(Path.GetFileName));
        Assert.Equal(Settings(6).ToJson(), File.ReadAllText(SettingsPath));
    }

    [Fact]
    public void Backups_in_the_same_second_are_numbered_and_ordered()
    {
        var file = CreateFile();
        file.Save(Settings(1));
        file.Save(Settings(2));
        file.Save(Settings(3));
        file.Save(Settings(4));

        var backups = file.GetBackups();

        Assert.Equal(
            new[] { "settings.json.20260929-100000-3.bak", "settings.json.20260929-100000-2.bak", "settings.json.20260929-100000.bak" },
            backups.Select(Path.GetFileName));
        Assert.Equal(new[] { Settings(3).ToJson(), Settings(2).ToJson(), Settings(1).ToJson() }, backups.Select(File.ReadAllText));
    }

    [Fact]
    public void No_new_backup_when_the_newest_backup_holds_the_replaced_contents()
    {
        var file = CreateFile();
        File.WriteAllText(SettingsPath, Settings(1).ToJson());
        var first = file.Save(Settings(2)).BackupPath!;
        File.Copy(first, SettingsPath, overwrite: true); // edited back by hand to what the backup holds
        _now = _now.AddMinutes(1);

        var result = file.Save(Settings(3));

        Assert.True(result.Succeeded);
        Assert.Null(result.BackupPath);
        Assert.Equal(new[] { first }, file.GetBackups());
        Assert.Equal(Settings(3).ToJson(), File.ReadAllText(SettingsPath));
    }

    [Fact]
    public void Files_that_only_look_like_backups_are_ignored_and_kept()
    {
        var file = CreateFile();
        var others = new[] { "settings.json.bak", "settings.json.2026-09-29.bak", "settings.json.20260929-100000.bak.txt", "other.json.20260929-100000.bak" };
        foreach (var other in others)
        {
            File.WriteAllText(Path.Combine(_folder, other), "x");
        }

        for (var cap = 1; cap <= 5; cap++)
        {
            _now = _now.AddSeconds(1);
            file.Save(Settings(cap));
        }

        Assert.Equal(SettingsFile.MaxBackups, file.GetBackups().Count);
        Assert.All(others, other => Assert.True(File.Exists(Path.Combine(_folder, other))));
    }

    [Fact]
    public void Detects_a_file_changed_since_it_was_loaded_or_saved()
    {
        var file = CreateFile();
        Assert.False(file.HasChangedSince(null)); // no file: nothing to overwrite

        var saved = file.Save(Settings(1)).Hash;
        Assert.False(file.HasChangedSince(saved));
        Assert.True(file.HasChangedSince(null));
        Assert.True(file.HasChangedSince("0000"));

        File.WriteAllText(SettingsPath, Settings(1).ToJson() + " ");
        Assert.True(file.HasChangedSince(saved));

        var loaded = file.Load();
        Assert.Equal(SettingsLoadOutcome.Loaded, loaded.Result.Outcome);
        Assert.False(file.HasChangedSince(loaded.Hash));

        File.Delete(SettingsPath);
        Assert.False(file.HasChangedSince(loaded.Hash));
    }

    [Fact]
    public void Hash_is_the_sha256_of_the_bytes()
    {
        Assert.Equal(
            "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad",
            SettingsFile.ComputeHash(Encoding.ASCII.GetBytes("abc")));
        Assert.Throws<ArgumentNullException>(() => SettingsFile.ComputeHash(null!));
    }

    [Fact]
    public void Load_creates_the_defaults_when_there_is_no_file_and_no_backup()
    {
        var file = CreateFile();

        var load = file.Load();

        Assert.Equal(SettingsLoadOutcome.CreatedDefaults, load.Result.Outcome);
        Assert.Empty(load.Result.Problems);
        Assert.Null(load.RestoredFrom);
        Assert.Equal(ToolkitSettings.Defaults().ToJson(), File.ReadAllText(SettingsPath));
        Assert.Equal(SettingsFile.ComputeHash(File.ReadAllBytes(SettingsPath)), load.Hash);
    }

    [Fact]
    public void Load_restores_the_newest_backup_when_the_file_is_missing()
    {
        var file = CreateFile();
        file.Save(Settings(1));
        _now = _now.AddSeconds(1);
        file.Save(Settings(2));
        _now = _now.AddSeconds(1);
        var newest = file.Save(Settings(3)).BackupPath;
        File.Delete(SettingsPath);

        var load = file.Load();

        Assert.Equal(SettingsLoadOutcome.Loaded, load.Result.Outcome);
        Assert.Equal(newest, load.RestoredFrom);
        Assert.Equal(Settings(2).ToJson(), load.Result.Settings.ToJson());
        Assert.Equal(Settings(2).ToJson(), File.ReadAllText(SettingsPath));
        Assert.Equal(SettingsFile.ComputeHash(File.ReadAllBytes(SettingsPath)), load.Hash);
        Assert.True(File.Exists(newest)); // the backup is copied, not moved
    }

    [Fact]
    public void Load_rejects_a_bad_file_leaves_it_alone_and_names_it()
    {
        var path = Path.Combine(_folder, "team.json");
        var file = new SettingsFile(path) { UtcNow = () => _now };
        var bytes = new byte[] { 0x7B, 0xFF, 0x7D };
        File.WriteAllBytes(path, bytes);

        var load = file.Load();

        Assert.Equal(SettingsLoadOutcome.Rejected, load.Result.Outcome);
        Assert.Equal("team.json is not valid UTF-8; save it as UTF-8.", Assert.Single(load.Result.Problems));
        Assert.Equal(SettingsFile.ComputeHash(bytes), load.Hash);
        Assert.Equal(bytes, File.ReadAllBytes(path));
        Assert.Equal("team.json", file.FileName);
    }

    [Fact]
    public void Load_rejects_an_oversized_file_without_reading_it()
    {
        var file = CreateFile();
        File.WriteAllBytes(SettingsPath, new byte[ToolkitSettings.MaxFileBytes + 1]);

        var load = file.Load();

        Assert.Equal(SettingsLoadOutcome.Rejected, load.Result.Outcome);
        Assert.StartsWith("settings.json is 1,048,577 bytes;", Assert.Single(load.Result.Problems));
        Assert.Null(load.Hash);
    }

    [Fact]
    public void A_busy_file_is_replaced_on_the_second_try()
    {
        var file = CreateFile();
        file.Save(Settings(1));
        var calls = 0;
        var replace = file.ReplaceFile;
        file.ReplaceFile = (source, destination, backup) =>
        {
            if (++calls == 1)
            {
                throw new IOException("The process cannot access the file because it is being used by another process.");
            }

            replace(source, destination, backup);
        };

        var result = file.Save(Settings(2));

        Assert.True(result.Succeeded);
        Assert.Equal(2, calls);
        Assert.Equal(Settings(2).ToJson(), File.ReadAllText(SettingsPath));
        Assert.Equal(Settings(1).ToJson(), File.ReadAllText(Assert.Single(file.GetBackups())));
        Assert.Empty(TemporaryFiles());
    }

    [Fact]
    public void Where_replace_keeps_failing_the_file_is_copied_deleted_and_renamed()
    {
        var file = CreateFile();
        file.Save(Settings(1));
        var calls = 0;
        file.ReplaceFile = (source, destination, backup) =>
        {
            calls++;
            throw new PlatformNotSupportedException();
        };

        var result = file.Save(Settings(2));

        Assert.True(result.Succeeded);
        Assert.Equal(1, calls); // not retried: it will never work here
        Assert.Equal(Settings(2).ToJson(), File.ReadAllText(SettingsPath));
        Assert.Equal(Settings(1).ToJson(), File.ReadAllText(result.BackupPath!));
        Assert.Empty(TemporaryFiles());
    }

    [Fact]
    public void A_save_that_fails_part_way_restores_the_newest_backup()
    {
        var file = CreateFile();
        file.Save(Settings(1));
        file.ReplaceFile = (source, destination, backup) =>
        {
            // Replace moved the file to the backup, then failed; the temporary file is gone too.
            if (File.Exists(destination))
            {
                File.Move(destination, backup!);
            }

            File.Delete(source);
            throw new IOException("Unable to remove the file to be replaced.");
        };

        var result = file.Save(Settings(2));

        Assert.False(result.Succeeded);
        Assert.StartsWith($"could not save {SettingsPath}:", result.Problem);
        Assert.EndsWith("The previous settings were restored from settings.json.20260929-100000.bak.", result.Problem);
        Assert.Null(result.Hash);
        Assert.Equal(Path.Combine(_folder, "settings.json.20260929-100000.bak"), result.RestoredFrom);
        Assert.Equal(Settings(1).ToJson(), File.ReadAllText(SettingsPath));
        Assert.Empty(TemporaryFiles());
    }

    [Fact]
    public void A_failed_save_that_leaves_the_file_changes_nothing()
    {
        var file = CreateFile();
        file.Save(Settings(1));
        using (new FileStream(SettingsPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            // Open without delete sharing: neither replace nor delete can remove it.
            var result = file.Save(Settings(2));

            Assert.False(result.Succeeded);
            Assert.Null(result.RestoredFrom);
            Assert.DoesNotContain("restored", result.Problem);
        }

        Assert.Equal(Settings(1).ToJson(), File.ReadAllText(SettingsPath));
        Assert.Empty(TemporaryFiles());
    }

    [Fact]
    public void EnsureExists_restores_a_backup_or_writes_the_defaults()
    {
        var file = CreateFile();
        Assert.Null(file.EnsureExists());
        Assert.Equal(ToolkitSettings.Defaults().ToJson(), File.ReadAllText(SettingsPath));

        file.Save(Settings(7));
        File.Delete(SettingsPath);

        Assert.Null(file.EnsureExists());
        Assert.Equal(ToolkitSettings.Defaults().ToJson(), File.ReadAllText(SettingsPath)); // the backup of the defaults

        File.WriteAllText(SettingsPath, "edited");
        Assert.Null(file.EnsureExists());
        Assert.Equal("edited", File.ReadAllText(SettingsPath));
    }

    [Fact]
    public void Stale_temporary_files_are_deleted()
    {
        var file = CreateFile();
        var stale = Path.Combine(_folder, "settings.json." + Guid.NewGuid().ToString("N") + ".tmp");
        var fresh = Path.Combine(_folder, "settings.json." + Guid.NewGuid().ToString("N") + ".tmp");
        var other = Path.Combine(_folder, "settings.json.mine.tmp");
        foreach (var path in new[] { stale, fresh, other })
        {
            File.WriteAllText(path, "x");
        }

        File.SetLastWriteTimeUtc(stale, _now.AddDays(-2));
        File.SetLastWriteTimeUtc(fresh, _now.AddHours(-1));
        File.SetLastWriteTimeUtc(other, _now.AddDays(-2));

        var deleted = file.DeleteStaleTemporaryFiles(TimeSpan.FromDays(1));

        Assert.Equal(1, deleted);
        Assert.False(File.Exists(stale));
        Assert.True(File.Exists(fresh));
        Assert.True(File.Exists(other));
        Assert.Equal(0, new SettingsFile(Path.Combine(_folder, "missing", "settings.json")).DeleteStaleTemporaryFiles(TimeSpan.Zero));
    }

    [Fact]
    public void Constructor_checks_the_path()
    {
        Assert.Throws<ArgumentNullException>(() => new SettingsFile(null!));
        Assert.Throws<ArgumentException>(() => new SettingsFile(_folder + Path.DirectorySeparatorChar));
        Assert.Throws<ArgumentNullException>(() => CreateFile().Save(null!));

        var file = CreateFile();
        Assert.Equal(SettingsPath, file.FilePath);
        Assert.Equal(_folder, file.DirectoryPath);
    }

    private static ToolkitSettings Settings(int undoCellCap) =>
        new ToolkitSettings(ToolkitSettings.Defaults().Cycles, ToolkitSettings.Defaults().Keymap, undoCellCap);

    private SettingsFile CreateFile() =>
        new SettingsFile(SettingsPath) { UtcNow = () => _now, RetryDelay = TimeSpan.Zero };

    private IEnumerable<string> TemporaryFiles() => Directory.GetFiles(_folder, "*.tmp");
}
