# Per-user MSI

`Modelwright-<version>-x64.msi` is an **optional** install path (README, Option B). The zip with `install.ps1` stays the recommended path, and the manual path needs neither scripts nor installers. Decided 2026-10-09 (docs/PLAN.md, Phase 5: D-msi).

| File | What it is |
|---|---|
| `Modelwright.wxs` | The package (WiX v5). |
| `CustomActions/` | C# custom actions (WixToolset.Dtf, .NET Framework 4.8): register and unregister the add-in with Excel, and the Excel/bitness/.NET checks. Not in `Modelwright.sln`, so the main build doesn't need WiX. |
| `CustomActions/AddInRegistration.cs` | The OPEN/OPENn rules as pure logic, unit-tested by `tests/Modelwright.Core.Tests/Installer/` (the file is linked into that project). |
| `build-msi.ps1` | Builds the custom actions and the MSI; `-Validate` also runs the ICE checks, lists the contents and extracts it with `msiexec /a`. |

Build it locally (see also CONTRIBUTING.md):

```powershell
dotnet tool install --global wix --version 5.0.2     # once
dotnet build Modelwright.sln -c Release
powershell -NoProfile -ExecutionPolicy Bypass -File installer/msi/build-msi.ps1 -Validate
# -> dist\Modelwright-<version>-x64.msi
```

Don't install a development MSI on your everyday profile unless you mean to: it registers the add-in in your own `HKCU` Excel settings. `-Validate` never installs anything.

## What the MSI does

- **Per user, no administrator rights** (`Scope="perUser"`). It writes only to the user's profile and `HKCU`.
- **Installs** `Modelwright64.xll`, `LICENSE.txt`, `THIRD_PARTY_NOTICES.md` and `INSTALL.txt` to `%LOCALAPPDATA%\Programs\Modelwright`.
- **Registers the add-in with Excel** as the next free `OPEN`/`OPENn` value of `HKCU\Software\Microsoft\Office\16.0\Excel\Options`, holding `/R "<path>"`. It follows the same rules as `install.ps1`:
  - an existing Modelwright entry (including an `install.ps1` one, a bare `/R "Modelwright64.xll"` that Excel rewrote, a `Modelwright32.xll` from an earlier build, or a pre-rename ModelingToolkit build) is replaced in its slot;
  - duplicates are removed and the numbering is compacted, with no gaps;
  - Modelwright entries in the `Add-in Manager` key are removed.
- **Uninstall** (Settings › Apps, or `msiexec /x`) removes the entry, renumbers the others, and deletes the files and the folder. Settings in `%APPDATA%\Modelwright` and the log in `%LOCALAPPDATA%\Modelwright` are kept, as with `uninstall.ps1`.
- **Upgrades in place.** Every version has the same `UpgradeCode`. The old version is removed first; its uninstall skips the unregister step (`UPGRADINGPRODUCTCODE`), so the add-in keeps its `OPENn` slot. A prerelease and its release share an MSI version (`0.1.0-rc.1` and `0.1.0` both become 0.1.0), so same-version upgrades are allowed.
- **Refuses**:
  - while Excel is running in the user's session (Retry / Cancel; a silent install fails with 1603). Excel rewrites its add-in list when it closes, which would undo the change.
  - on 32-bit Excel: "Modelwright is 64-bit only…". The check reads `EXCEL.EXE`'s PE header, found through App Paths or Click-to-Run, then the Click-to-Run `Platform` setting, the same order as `install.ps1`. If Excel can't be found, it installs anyway and says so in the log.
  - without .NET Framework 4.8.
  - on 32-bit Windows (x64 package).
- **Add/remove programs** entry: name, publisher (Pegasus Technology Group LLC), version, and links to the repository, the install guide and the releases. Modify is hidden (there is nothing to choose).
- **Rollback.** Before changing anything, the custom action snapshots the `OPEN` values and the Modelwright `Add-in Manager` entries. A failed install or uninstall restores them.

Silent install and logging: `msiexec /i Modelwright-<version>-x64.msi /qn /l*v install.log`. Exit codes are Windows Installer's: 0 success, 1602 cancelled, 1603 failure (the log says why, on lines starting `Modelwright:`).

## Decisions and trade-offs

**Why a custom action.** Excel loads Add-ins-dialog add-ins only from the numbered `OPEN`, `OPEN1`, … values, and the numbers must have no gaps. The `HKCU\Software\Microsoft\Office\Excel\Addins` key is for COM add-ins only, not `.xll`. A fixed `OPENn` written by the `Registry` table could overwrite another add-in's entry or leave a gap, and Windows Installer can't renumber values on uninstall. So a custom action is needed, as in Excel-DNA's own WiX template.

**Why C# (DTF), and not PowerShell.** Running `install.ps1` from the MSI would fail wherever Group Policy sets the execution policy to AllSigned or Restricted, which are the very environments that pick an MSI. A managed custom action runs inside Windows Installer and needs only .NET Framework 4.x. That's already a requirement, and it's checked first.

The C# keeps `install.ps1`'s rules: the same pattern for "our" entries, the same slot handling, and its cases are unit-tested on both test frameworks. If you change the rules, change them in both places. The custom actions run as the user (immediate and impersonated deferred actions), so `HKCU` is the installing user's hive.

**Why `%LOCALAPPDATA%\Programs\Modelwright`, not `%APPDATA%\Microsoft\AddIns`.** `install.ps1` uses Excel's own add-ins folder (research/08 §1.2). The MSI uses Windows Installer's per-user programs folder instead, for three reasons:
1. **The file belongs to the MSI alone.** The `.xll`'s file version is Excel-DNA's (1.9.0.14 for every Modelwright release), so Windows Installer's versioning rules would not replace a same-version copy left in `AddIns` by the script or by hand. Uninstalling the MSI would also delete a file the script put there.
2. **Excel rewrites entries for files in its own `AddIns` folder** to a bare `/R "Modelwright64.xll"`. A full path outside that folder stays a full path, which keeps uninstall exact.
3. **It's the conventional per-user install location** that Settings › Apps and IT tools expect.

The cost: `%LOCALAPPDATA%` doesn't roam. With roaming profiles (VDI/Citrix), the `HKCU` entry would follow the user to machines without the file, and Excel would say it can't find the add-in. Those environments should use the script, or IT's own deployment (README, Option D).

**Why WiX v5, not v6/v7.** From v6, WiX binaries come with the Open Source Maintenance Fee EULA. Organizations with ≥ US$10,000 annual revenue that use WiX "as part of revenue-generating activities" must sponsor it, and the EULA must be accepted when building (research/08 §4.1). WiX v5.0.2 has no such EULA, which keeps that legal question out of v0.1. WiX v5 and DTF are MS-RL; the custom-action DLL embeds `WixToolset.Dtf.WindowsInstaller.dll` unmodified (see THIRD_PARTY_NOTICES.md). Moving to a newer WiX is an owner decision.

**No UI dialogs.** Double-clicking shows only Windows Installer's progress window, plus our Retry/Cancel and error messages. There's nothing to choose. The license is in the repository and installed next to the add-in.

**Unsigned in v0.1** (owner decision, 2026-10-09). SmartScreen shows "Windows protected your PC", and the user chooses **More info › Run anyway**. AppLocker's default Windows Installer rules let standard users run only *signed* MSIs, so in AppLocker shops this MSI is blocked until a signed build exists. The release workflow signs the MSI (after signing the `.xll` it contains) once the `SIGNING_*` secrets are set.

## Not verified yet

An install, upgrade and uninstall on a real profile hasn't been run. CI and `-Validate` only validate the package, list it and extract it. Before the first release, on a test VM or a spare Windows user account, with Excel 64-bit:
- install, then check the `OPEN` entry and the ribbon tab;
- install again over an `install.ps1` install, and check that a single entry remains;
- upgrade with a higher version, and check that the slot is unchanged;
- uninstall, and check that the entry is removed and the others renumbered;
- try it with Excel open (Retry/Cancel) and silently (1603);
- check that a 32-bit Excel machine is refused.
