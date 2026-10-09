using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using WixToolset.Dtf.WindowsInstaller;

namespace Modelwright.Installer;

/// <summary>
/// The MSI's custom actions. All run as the installing user (a per-user package needs no elevation), so
/// HKEY_CURRENT_USER is that user's hive.
/// <list type="bullet">
/// <item><see cref="CheckPrerequisites"/> (immediate, before InstallValidate): Excel must be closed (it rewrites its
/// add-in list when it closes, which would undo the change); on install, Excel must not be 32-bit and .NET
/// Framework 4.8 must be present.</item>
/// <item><see cref="ScheduleRegistration"/> (immediate, after InstallFiles): snapshots the OPEN values and the
/// Add-in Manager key and schedules <see cref="RollbackRegistration"/> with that snapshot and
/// <see cref="ApplyRegistration"/> with what to do (register on install and repair, unregister on uninstall,
/// nothing for the old version's uninstall during an upgrade, so the add-in keeps its slot).</item>
/// </list>
/// </summary>
public static class CustomActions
{
    private const string XllName = "Modelwright64.xll";

    // .NET Framework 4.8: NDP\v4\Full Release 528040 or later.
    private const int NetFramework48Release = 528040;

    private const string ClickToRunKey = @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration";
    private const string AppPathsKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe";

    [CustomAction("MwCheckPrerequisites")]
    public static ActionResult CheckPrerequisites(Session session)
    {
        // The old version's uninstall during an upgrade: the new version's install has already checked.
        if (!string.IsNullOrEmpty(session["UPGRADINGPRODUCTCODE"]))
        {
            return ActionResult.Success;
        }

        var removing = session["REMOVE"] == "ALL";

        while (IsExcelRunning())
        {
            var answer = Show(session, InstallMessage.Warning, MessageButtons.RetryCancel,
                "Excel is running. Close Excel completely (check the taskbar, and Task Manager for a background " +
                "EXCEL.EXE), then click Retry. Excel rewrites its add-in list when it closes, which would undo " +
                (removing ? "the removal." : "the install."));
            if (answer == MessageResult.Retry)
            {
                continue;
            }

            session.Log("Modelwright: Excel is running; stopping.");
            return answer == MessageResult.Cancel ? ActionResult.UserExit : ActionResult.Failure;
        }

        if (removing)
        {
            return ActionResult.Success;
        }

        var bitness = GetExcelBitness(session);
        if (bitness == "32")
        {
            Show(session, InstallMessage.Error, MessageButtons.OK,
                "Modelwright is 64-bit only; your Excel is 32-bit, so it cannot load Modelwright. 32-bit Excel, " +
                "Excel for Mac and Excel for the web are not supported. Nothing was installed.");
            session.Log("Modelwright: Excel is 32-bit; stopping.");
            return ActionResult.Failure;
        }

        if (bitness is null)
        {
            session.Log("Modelwright: could not find EXCEL.EXE or tell its bitness; installing anyway.");
        }

        if (GetNetFrameworkRelease() < NetFramework48Release)
        {
            Show(session, InstallMessage.Error, MessageButtons.OK,
                "Modelwright needs .NET Framework 4.8 or later, which is not installed. Install it from " +
                "https://dotnet.microsoft.com/download/dotnet-framework (or ask IT), then run this again.");
            session.Log("Modelwright: .NET Framework 4.8 is missing; stopping.");
            return ActionResult.Failure;
        }

        return ActionResult.Success;
    }

    [CustomAction("MwScheduleRegistration")]
    public static ActionResult ScheduleRegistration(Session session)
    {
        string mode;
        var value = string.Empty;
        if (session["REMOVE"] == "ALL")
        {
            if (!string.IsNullOrEmpty(session["UPGRADINGPRODUCTCODE"]))
            {
                session.Log("Modelwright: upgrade; the new version keeps the Excel registration.");
                return ActionResult.Success;
            }

            mode = "unregister";
        }
        else
        {
            mode = "register";
            value = AddInRegistration.OpenValue(Path.Combine(session["INSTALLFOLDER"], XllName));
        }

        var rollback = new CustomActionData { ["snapshot"] = TakeSnapshot() };
        var apply = new CustomActionData { ["mode"] = mode, ["value"] = value };
        session.DoAction("MwRollbackRegistration", rollback);
        session.DoAction("MwApplyRegistration", apply);
        return ActionResult.Success;
    }

    [CustomAction("MwApplyRegistration")]
    public static ActionResult ApplyRegistration(Session session)
    {
        var data = session.CustomActionData;
        var register = data["mode"] == "register";
        var value = data["value"];

        using (var options = register
                   ? Registry.CurrentUser.CreateSubKey(AddInRegistration.OptionsSubKey)
                   : Registry.CurrentUser.OpenSubKey(AddInRegistration.OptionsSubKey, writable: true))
        {
            if (options is not null)
            {
                var entries = AddInRegistration.GetOpenEntries(ReadValues(options));
                var plan = register ? AddInRegistration.PlanRegister(entries, value) : AddInRegistration.PlanUnregister(entries);
                if (plan is null)
                {
                    session.Log(register
                        ? "Modelwright: already registered with Excel; left as it is."
                        : "Modelwright: not registered with Excel; nothing to remove.");
                }
                else
                {
                    foreach (var change in AddInRegistration.Diff(entries, plan))
                    {
                        if (change.IsDelete)
                        {
                            session.Log("Modelwright: removing " + change.Name);
                            options.DeleteValue(change.Name, throwOnMissingValue: false);
                        }
                        else
                        {
                            session.Log("Modelwright: setting " + change.Name + " = " + change.Value);
                            options.SetValue(change.Name, change.Value!, RegistryValueKind.String);
                        }
                    }
                }
            }
        }

        // Excel's list of known-but-unticked add-ins: drop Modelwright entries (stale paths or bare names).
        using (var manager = Registry.CurrentUser.OpenSubKey(AddInRegistration.AddInManagerSubKey, writable: true))
        {
            if (manager is not null)
            {
                foreach (var name in manager.GetValueNames().Where(AddInRegistration.IsOurs))
                {
                    session.Log("Modelwright: removing Add-in Manager entry " + name);
                    manager.DeleteValue(name, throwOnMissingValue: false);
                }
            }
        }

        return ActionResult.Success;
    }

    [CustomAction("MwRollbackRegistration")]
    public static ActionResult RollbackRegistration(Session session)
    {
        var snapshot = ParseSnapshot(session.CustomActionData["snapshot"]);
        using (var options = Registry.CurrentUser.OpenSubKey(AddInRegistration.OptionsSubKey, writable: true))
        {
            if (options is not null)
            {
                foreach (var entry in AddInRegistration.GetOpenEntries(ReadValues(options)))
                {
                    options.DeleteValue(entry.Name, throwOnMissingValue: false);
                }

                foreach (var pair in snapshot.Where(s => s.Kind == 'O'))
                {
                    options.SetValue(pair.Name, pair.Value, RegistryValueKind.String);
                }
            }
        }

        var removed = snapshot.Where(s => s.Kind == 'M').ToList();
        if (removed.Count > 0)
        {
            using var manager = Registry.CurrentUser.CreateSubKey(AddInRegistration.AddInManagerSubKey);
            foreach (var pair in removed)
            {
                if (manager.GetValue(pair.Name) is null)
                {
                    manager.SetValue(pair.Name, pair.Value, RegistryValueKind.String);
                }
            }
        }

        session.Log("Modelwright: Excel registration restored.");
        return ActionResult.Success;
    }

    private static MessageResult Show(Session session, InstallMessage type, MessageButtons buttons, string text)
    {
        session.Log("Modelwright: " + text);
        using var record = new Record(0);
        record.FormatString = text;
        return session.Message(type | (InstallMessage)buttons, record);
    }

    private static bool IsExcelRunning()
    {
        var sessionId = Process.GetCurrentProcess().SessionId;
        var excels = Process.GetProcessesByName("EXCEL");
        try
        {
            return excels.Any(p => p.SessionId == sessionId);
        }
        finally
        {
            foreach (var process in excels)
            {
                process.Dispose();
            }
        }
    }

    // "64", "32", or null if Excel cannot be found or read. Same order as install.ps1: EXCEL.EXE's PE header
    // (path from App Paths or Click-to-Run), then the Click-to-Run Platform setting.
    private static string? GetExcelBitness(Session session)
    {
        var candidates = new List<string>();
        foreach (var (hive, view) in new[]
                 {
                     (RegistryHive.LocalMachine, RegistryView.Registry64),
                     (RegistryHive.LocalMachine, RegistryView.Registry32),
                     (RegistryHive.CurrentUser, RegistryView.Default),
                 })
        {
            var path = ReadString(hive, view, AppPathsKey, string.Empty);
            if (!string.IsNullOrWhiteSpace(path))
            {
                candidates.Add(path!.Trim().Trim('"'));
            }
        }

        var installPath = ReadString(RegistryHive.LocalMachine, RegistryView.Registry64, ClickToRunKey, "InstallationPath");
        if (!string.IsNullOrWhiteSpace(installPath))
        {
            candidates.Add(Path.Combine(installPath, @"root\Office16\EXCEL.EXE"));
        }

        foreach (var exe in candidates)
        {
            var bitness = ReadExeBitness(exe);
            session.Log("Modelwright: " + exe + " -> " + (bitness ?? "unreadable"));
            if (bitness is not null)
            {
                return bitness;
            }
        }

        return ReadString(RegistryHive.LocalMachine, RegistryView.Registry64, ClickToRunKey, "Platform") switch
        {
            "x64" => "64",
            "x86" => "32",
            _ => null,
        };
    }

    // The PE header's Machine field: x64 and Arm64 (64-bit Excel on Arm loads x64 add-ins) are 64, x86 is 32.
    private static string? ReadExeBitness(string path)
    {
        try
        {
            var bytes = new byte[4096];
            int read;
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                read = stream.Read(bytes, 0, bytes.Length);
            }

            if (read < 64 || bytes[0] != 0x4D || bytes[1] != 0x5A)
            {
                return null;
            }

            var peOffset = BitConverter.ToInt32(bytes, 0x3C);
            if (peOffset < 0 || peOffset + 6 > read ||
                bytes[peOffset] != 0x50 || bytes[peOffset + 1] != 0x45 || bytes[peOffset + 2] != 0 || bytes[peOffset + 3] != 0)
            {
                return null;
            }

            return BitConverter.ToUInt16(bytes, peOffset + 4) switch
            {
                0x8664 => "64",
                0xAA64 => "64",
                0x014C => "32",
                _ => null,
            };
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException || ex is NotSupportedException)
        {
            return null;
        }
    }

    private static int GetNetFrameworkRelease()
    {
        using var root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
        using var key = root.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full");
        return key?.GetValue("Release") is int release ? release : 0;
    }

    private static string? ReadString(RegistryHive hive, RegistryView view, string subKey, string name)
    {
        using var root = RegistryKey.OpenBaseKey(hive, view);
        using var key = root.OpenSubKey(subKey);
        return key?.GetValue(name) as string;
    }

    private static IEnumerable<KeyValuePair<string, string>> ReadValues(RegistryKey key)
    {
        return key.GetValueNames()
            .Select(name => new KeyValuePair<string, string>(
                name,
                Convert.ToString(key.GetValue(name, string.Empty, RegistryValueOptions.DoNotExpandEnvironmentNames),
                    System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty))
            .ToList();
    }

    // The OPEN values and our Add-in Manager entries as "O:<name>:<value>" / "M:<name>:<value>" items (name and
    // value Base64-encoded UTF-8) joined by '|', so the text needs no escaping in CustomActionData.
    private static string TakeSnapshot()
    {
        var items = new List<string>();
        using (var options = Registry.CurrentUser.OpenSubKey(AddInRegistration.OptionsSubKey))
        {
            if (options is not null)
            {
                items.AddRange(AddInRegistration.GetOpenEntries(ReadValues(options)).Select(e => Encode('O', e.Name, e.Value)));
            }
        }

        using (var manager = Registry.CurrentUser.OpenSubKey(AddInRegistration.AddInManagerSubKey))
        {
            if (manager is not null)
            {
                items.AddRange(ReadValues(manager).Where(p => AddInRegistration.IsOurs(p.Key)).Select(p => Encode('M', p.Key, p.Value)));
            }
        }

        return string.Join("|", items);
    }

    private static string Encode(char kind, string name, string value) =>
        kind + ":" + Convert.ToBase64String(Encoding.UTF8.GetBytes(name)) + ":" + Convert.ToBase64String(Encoding.UTF8.GetBytes(value));

    private static List<(char Kind, string Name, string Value)> ParseSnapshot(string snapshot)
    {
        var items = new List<(char Kind, string Name, string Value)>();
        foreach (var item in snapshot.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = item.Split(':');
            items.Add((parts[0][0],
                Encoding.UTF8.GetString(Convert.FromBase64String(parts[1])),
                Encoding.UTF8.GetString(Convert.FromBase64String(parts[2]))));
        }

        return items;
    }
}
