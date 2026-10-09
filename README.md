# Modelwright

Modelwright is a source-available Excel add-in for financial modelers:

1. **Number format cycling**: step the selected cells through your own list of formats with one shortcut.
2. **Font color cycling**
3. **Fill color cycling**
4. **Smart trace precedents**: a keyboard-driven tree of a cell's precedents, with values, that works across sheets.

These are the features people pay for in Macabacus, FactSet Spreadsheet Tools, TTS Turbo Macros and similar tools. This project aims to provide them for free, using **exactly the same keyboard shortcuts as Macabacus**. The first target is **Excel for Windows desktop**. Mac and web are deferred, because Office.js can't bind Macabacus's punctuation keys.

## Install

Modelwright runs in **64-bit Windows Excel**: Microsoft 365, Excel 2024 and Excel 2021. **32-bit Excel, Excel for Mac and Excel for the web are not supported**: the add-in is a 64-bit Windows file (`.xll`) that they cannot load. Windows on Arm is untested. You do **not** need administrator rights.

| Option | Use it when |
|---|---|
| **A. Install script** (recommended) | You can run PowerShell scripts. It checks your Excel, unblocks the add-in and registers it; the fewest mistakes. |
| **B. By hand** | Scripts are not allowed. Uses only File Explorer and Excel's own Add-ins dialog. |
| **C. IT, or build from source** | Your company requires signed add-ins, controls software centrally, or forbids downloaded binaries. |

An MSI installer is planned for a future release (#13).

<!--
Screenshots to add (docs/research/08-install-and-distribution.md section 10). Not captured yet; link them
only once the files exist:
  docs/images/releases-assets.png         GitHub Releases page, Assets expanded, zip highlighted
  docs/images/edge-keep.png               Edge download bubble: Keep > Show more > Keep anyway (if shown)
  docs/images/unblock-properties.png      zip > Properties > General, with the Unblock box
  docs/images/extract-all.png             Extract All dialog
  docs/images/addins-folder.png           File Explorer with %APPDATA%\Microsoft\AddIns in the address bar
  docs/images/about-excel-64bit.png       File > Account > About Excel showing "64-bit"
  docs/images/options-addins.png          File > Options > Add-ins, Manage: Excel Add-ins and Go... highlighted
  docs/images/addins-dialog.png           Add-ins dialog: Browse... highlighted, then Modelwright ticked
  docs/images/ribbon-tab.png              the Modelwright ribbon tab
  docs/images/motw-security-notice.png    "Microsoft Excel Security Notice" (quote its text verbatim below)
  docs/images/bitness-mismatch.png        "The file format and extension of ... don't match"
  docs/images/untrusted-publisher.png     signed-but-untrusted-publisher notice (once releases are signed)
  docs/images/run-with-powershell.png     right-click install.ps1 > Run with PowerShell, and its output
  docs/images/com-addins-macabacus.png    Manage: COM Add-ins dialog with Macabacus listed
  docs/images/disabled-items.png          Manage: Disabled Items dialog
-->

### Before you start

1. **Close Excel.**
2. **Check that your Excel is 64-bit:** in Excel, File › Account › About Excel. The first line ends in "64-bit" or "32-bit". Most computers have 64-bit; Modelwright can't load in 32-bit Excel. (Option A checks this for you.)
3. **Why you must "Unblock" the download.** Windows marks every file that comes from the internet. Excel refuses to load a marked add-in: it shows a "Microsoft Excel Security Notice" whose only button is **Leave this add-in disabled**. Unblocking tells Windows you trust this one file; it doesn't change any security setting. If you unblock the zip *before* extracting it, everything inside is clean. (Option A unblocks the add-in for you.)

### Why Windows warns you

Release v0.1 is **not code-signed** (signed builds are planned). So Excel refuses the downloaded add-in until it is unblocked (above); the install script does that for you. (Windows SmartScreen may also show **"Windows protected your PC"** when you open an unsigned download: click **More info**, check the file name, then **Run anyway**.)

**If your company requires signed add-ins** ("Require that application add-ins are signed by Trusted Publisher", part of Microsoft's Microsoft 365 Apps security baseline), an unsigned add-in can't load there, whatever you do. Ask IT to trust Modelwright's publisher once a signed build exists, or to build and deploy it themselves (Option C).

### Option A: download and run the install script (recommended)

1. On the [Releases](https://github.com/jwriccardi/modelwright/releases) page, under **Assets**, download `Modelwright-<version>.zip`. If your browser warns that the file isn't commonly downloaded, choose **Keep**.
2. In your Downloads folder, right-click the zip › **Properties** › tick **Unblock** › **OK**. (No Unblock box means the file wasn't marked, which is fine.)
3. Right-click the zip › **Extract All…** › **Extract**.
4. In the extracted folder, right-click `install.ps1` › **Run with PowerShell** (on Windows 11 it may be under **Show more options**).
   If that doesn't work, open PowerShell and paste this line (change the folder if you extracted somewhere else):
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\Downloads\Modelwright\install.ps1"
   ```
5. Start Excel. The **Modelwright** tab appears on the ribbon.

The script refuses to run while Excel is open, and stops with a clear message (exit code 8) on 32-bit Excel. It copies the add-in to your own add-ins folder (`%APPDATA%\Microsoft\AddIns`), unblocks it, and adds it to Excel's add-in list (the same registry entry Excel's Add-ins dialog writes). Running it again is safe. Add `-WhatIf` to see what it would do without changing anything. `INSTALL.txt` in the zip lists every exit code.

### Option B: add it to Excel by hand (no scripts)

1. Do steps 1–3 of Option A.
2. Copy `Modelwright64.xll` to your add-ins folder: paste `%APPDATA%\Microsoft\AddIns` into the File Explorer address bar and press Enter. (You can instead browse to the extracted file in step 4, but then don't move or delete it.)
3. Start Excel and go to **File › Options › Add-ins**. At the bottom, set **Manage: Excel Add-ins** and click **Go…**.
4. Click **Browse…**, select the `.xll`, and click **OK**.
5. Make sure **Modelwright** is ticked, then click **OK**. The **Modelwright** tab appears.

The zip also has these steps in `INSTALL.txt`.

### Option C: ask IT, or build from source

<details>
<summary>For IT staff, and for anyone who must not run downloaded binaries</summary>

- **Verify the files.** Each release has a `SHA256SUMS.txt`, and GitHub shows a SHA-256 digest next to each asset: `Get-FileHash .\Modelwright64.xll -Algorithm SHA256` (or `certutil -hashfile Modelwright64.xll SHA256`). Releases made while the repository is public also have a build-provenance attestation that ties each asset to the workflow run and commit that built it: `gh attestation verify .\Modelwright64.xll -R jwriccardi/modelwright`. The release notes say whether that release is code-signed.
- **Deploy.** Copy the `.xll` to a local folder (not a network share; Program Files is fine for a per-machine copy). With Excel closed, set the user's next free `HKCU\Software\Microsoft\Office\16.0\Excel\Options` value (`OPEN`, `OPEN1`, `OPEN2`… with no gaps) to the string `/R "C:\path\to\Modelwright64.xll"`, for example with a Group Policy Preferences registry item or Intune. `install.ps1` does exactly this, and runs in Constrained Language mode.
- **Policies that block it.** Excel's block on `.xll` files from the internet (remove the mark of the web from the file); "Require that application add-ins are signed by Trusted Publisher", which Microsoft's Microsoft 365 Apps security baseline turns on (an unsigned add-in can never load there); AppLocker or App Control DLL rules (they need a path or publisher rule). See [research/08](docs/research/08-install-and-distribution.md) §1.5 and §2.
- **Build from source.** Windows 10/11, the .NET SDK 10 (`global.json` pins 10.0.100 with `rollForward: latestFeature`; it installs without admin rights with [dotnet-install.ps1](https://learn.microsoft.com/dotnet/core/tools/dotnet-install-script)), and access to nuget.org or your NuGet proxy. From the source of a release tag, run `dotnet build Modelwright.sln -c Release`. The add-in is written to `src\Modelwright.AddIn\bin\Release\net48\publish\Modelwright64.xll`; install it with Option B, or with `install.ps1 -SourceFolder <that folder>`. The build embeds its date, so a rebuilt file's hash differs from the release file's.

</details>

### Troubleshooting

| What you see | What to do |
|---|---|
| "Microsoft Excel Security Notice" with only **Leave this add-in disabled** | The file is still marked as downloaded. Close Excel, right-click the `.xll` in `%APPDATA%\Microsoft\AddIns` › Properties › tick **Unblock** (or run `install.ps1` again), then start Excel. Don't turn the block off in the Trust Center: that lowers security for every add-in. |
| "The file format and extension of '…xll' don't match" | Your Excel is 32-bit, which Modelwright doesn't support (it is 64-bit only). |
| The Modelwright tab is missing | File › Options › Add-ins › Manage: **Disabled Items** › Go… › select Modelwright › **Enable**. |
| "Sorry, we couldn't find …xll" | The file was moved or deleted. Run `install.ps1` again, or untick it in the Add-ins dialog. |
| Shortcuts do nothing, or do something else | Another add-in (often Macabacus) took the keys. See below. |

### Using Modelwright alongside Macabacus

Both add-ins use the same keyboard shortcuts. Whichever add-in registered its shortcuts last gets the key, so run only one of them at a time.
- **Turn Macabacus off:** File › Options › Add-ins › Manage: **COM Add-ins** › **Go…** › untick Macabacus › OK. It stays off until you tick it again.
- **Turn Modelwright off:** File › Options › Add-ins › Manage: **Excel Add-ins** › **Go…** › untick **Modelwright** › OK.
- **For this session only:** click Modelwright › **Re-register shortcuts** to let Modelwright win, or Macabacus's **Override** to let Macabacus win.

### Where your settings live

- Settings: `%APPDATA%\Modelwright\settings.json` (change them with Modelwright › Settings…). Settings from the earlier ModelingToolkit builds are copied over on first start.
- Log: `%LOCALAPPDATA%\Modelwright\log.txt`.

Uninstalling never deletes them. Delete those folders yourself if you want them gone.

### Uninstall

Close Excel first, and remove Modelwright the way you installed it.
- **Script:** right-click `uninstall.ps1` › **Run with PowerShell**, or paste:
  ```powershell
  powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\Downloads\Modelwright\uninstall.ps1" -RemoveFile
  ```
  It removes Modelwright from Excel's add-in list; `-RemoveFile` also deletes the `.xll` from `%APPDATA%\Microsoft\AddIns`.
- **By hand:** File › Options › Add-ins › Manage: Excel Add-ins › Go… › untick **Modelwright** › OK. Then close Excel and delete the `.xll`.

## Status: all four features built; release work in progress

The architecture is decided: an **Excel-DNA (C#) add-in for Windows desktop Excel** ([ADR-0002](docs/decisions/0002-excel-dna-windows-first.md)), validated by decision spikes ([results](docs/spike-results.md)). Working today, on Macabacus's exact keys:
- **Number-format cycles:** Ctrl+Shift+1 / 2 / 4 / 5 / 8, Ctrl+Shift+Y and Alt+Shift+;.
- **Color cycles:** Ctrl+' (font), Ctrl+Shift+K (fill) and Ctrl+; (blue/black toggle).
- **Undo:** Ctrl+Z / Ctrl+Y work alongside Excel's own undo.
- **Settings dialog** (Modelwright › Settings…).
- **Trace In** (Ctrl+Shift+[): a tree of the formula's precedents across sheets and workbooks (closed ones are opened read-only, OneDrive/SharePoint links included); Up/Down go there, Right/Left expand and collapse, Enter stays, Esc returns, **F2** edits the traced reference in Point mode, **Ctrl+E** evaluates functions and groups with Excel's argument names, **Ctrl+Shift+\** returns to the last audited cell.

Left before a first release ([PLAN](docs/PLAN.md)): the recorded manual test run, VirusTotal checks and screenshots. v0.1 ships unsigned and 64-bit only, with no MSI (owner decisions, 2026-10-09; the MSI is deferred, #13).

### Build and test

Requires the .NET SDK 10 on Windows; no Visual Studio needed. See [CONTRIBUTING.md](CONTRIBUTING.md) for building, testing, loading a development build in Excel and the Excel smoke tests. Contributions need a DCO sign-off.

### Docs

| Doc | What it covers |
|---|---|
| [docs/PLAN.md](docs/PLAN.md) | Work plan, design and acceptance criteria |
| [docs/spike-results.md](docs/spike-results.md) | Decision spike results (keys, undo, trace window, Office.js) |
| [docs/open-questions.md](docs/open-questions.md) | Decisions still needed, and points to check against Macabacus |
| [docs/decisions/0002-excel-dna-windows-first.md](docs/decisions/0002-excel-dna-windows-first.md) | ADR: Excel-DNA (C#), Windows first (accepted) |
| [docs/decisions/0001-platform-architecture.md](docs/decisions/0001-platform-architecture.md) | ADR: Office.js (superseded) |
| [docs/research/01-feature-survey.md](docs/research/01-feature-survey.md) | How Macabacus and competitors implement each feature |
| [docs/research/02-architecture-options.md](docs/research/02-architecture-options.md) | Office.js vs Excel-DNA vs VSTO vs VBA vs others |
| [docs/research/03-licensing.md](docs/research/03-licensing.md) | License options (MIT recommended) |
| [docs/research/04-xlerate-evaluation.md](docs/research/04-xlerate-evaluation.md) | Existing open-source prior art: whether to fork or build fresh |
| [docs/research/05-keys-and-undo.md](docs/research/05-keys-and-undo.md) | Which architectures can bind Macabacus's keys, and the options for undo |
| [docs/research/06-macabacus-observed-config.md](docs/research/06-macabacus-observed-config.md) | The owner's Macabacus settings: cycles, colors and the full keymap |
| [docs/research/07-macabacus-trace-in-spec.md](docs/research/07-macabacus-trace-in-spec.md) | How Macabacus's Trace In behaves, and what that means for our design |
| [docs/research/08-install-and-distribution.md](docs/research/08-install-and-distribution.md) | No-admin install, the internet block, signing, the MSI (deferred), release assets and bitness |

## License

[PolyForm Shield 1.0.0](LICENSE), Copyright (c) 2026 Pegasus Technology Group LLC. In short: you may use the add-in for any purpose, including inside your company, and change it privately with no obligation to share your changes; you may not offer a product that competes with it (selling it or a derivative, even free). This is a "source-available" license, not an [OSI open-source](https://opensource.org/osd) one. See [ADR 0003](docs/decisions/0003-polyform-shield-license.md). Third-party components are listed in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
