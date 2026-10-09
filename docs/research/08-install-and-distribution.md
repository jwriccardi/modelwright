# 08 — Install and distribution

*Researched 2026-10-08 for PLAN Phase 5 (Release). The product is being renamed **Modelwright**; this doc uses that name. File names (`Modelwright64.xll` and so on) are proposals. Today the build produces only `ModelingToolkit64-packed.xll`.*

**Owner's requirements**
- A non-technical user can install it by visiting the public GitHub repo.
- An MSI must **not** be the only path, because many users can't install compiled products in restricted corporate environments.
- Distributing the source is fine. Binaries and an MSI are fine as extra paths.
- 32-bit Excel must be planned for.

**Tags used below**
- **[V]**: verified against the cited source.
- **[O]**: observed on the owner's machine today (read-only registry and file queries; Excel was not launched).
- **[I]**: inferred from verified facts.
- **[U]**: unverified. A screenshot or spike should confirm it before the README relies on it.

---

## 0. Summary

1. **An .xll needs no admin rights.** Install it per user:
   - copy it into the user profile;
   - register it under `HKCU\Software\Microsoft\Office\16.0\Excel\Options` as the next free `OPEN`/`OPENn` value, with Excel closed.

   That is all Excel's own Add-ins dialog does. Excel-DNA documents both ways. [V]
2. **The biggest friction is Mark-of-the-Web (MOTW).** Since March 2023, Excel blocks an .xll that came from the internet. The block dialog has no "Enable" button, only **Leave this add-in disabled**. The fix is to **Unblock** the downloaded file before using it. If the user unblocks the .zip *before* extracting it, the extracted files are clean. [V]
3. **Signing doesn't get past the MOTW block.** It does decide whether the add-in can load *at all* in enterprises that follow Microsoft's own M365 Apps security baseline. That baseline enables "Require that application add-ins are signed by Trusted Publisher". Unsigned add-ins are dead there; signed ones need a one-time trust. Signing also helps with SmartScreen, Smart App Control, AppLocker publisher rules and antivirus. [V]/[I]
4. **The signing decision is the owner's** (§3):
   - **Azure Artifact Signing**: about $9.99/month, US/CA individuals or eligible orgs. Certificates are short-lived, which has a Trusted-Publisher catch.
   - **An OV certificate**: about $150–300/year, plus a hardware token or cloud HSM.
   - **SignPath Foundation**: free for OSS, but the publisher shown is "SignPath Foundation".
   - **Unsigned.**
5. **Ship both bitnesses.** It's a one-line build change (`ExcelDnaCreate32BitAddIn=true`), and net48 is fine for both. 32-bit needs its own test machine. [V]/[I]
6. **There are three install paths: manual, script, and build from source.** A per-user MSI is an optional fourth. winget is possible only through an MSI. Chocolatey, Scoop and MSIX aren't worth it for v0.1. [V]/[I]

---

## 1. No-admin, per-user install of an .xll

### 1.1 How Excel registers add-ins
- **The `OPEN` keys** [V]
  - Excel loads Add-ins-dialog add-ins from `HKCU\Software\Microsoft\Office\<ver>\Excel\Options`, in the values `OPEN`, `OPEN1`, `OPEN2` and so on.
  - Excel-DNA's "Installing your add-in" page lists this as an installer option: set the value to `/R "C:\...\MyAddIn.xll"`, which can be done from a `.bat` file using `reg`.
  - It warns: "Do not leave gaps if uninstalling … Eg OPEN, OPEN1, OPEN3."
  - Source: https://github.com/Excel-DNA/docs.excel-dna.net/blob/gh-pages/installing-your-add-in.md
- **Add-ins listed but not ticked** are kept as full paths under `HKCU\Software\Microsoft\Office\<ver>\Excel\Add-in Manager`. [V] https://jkp-ads.com/articles/AddinsAndSetupFactory.asp
  - [O] On the owner's machine, that key lists earlier dev builds by path, for example `...\bin\Release\net48\publish\ModelingToolkit64-packed.xll`.
- **Excel must be closed.** "The Keys are updated AFTER closing Excel … otherwise the addin will not be installed (because Excel will update the registry upon closing, it will remove the keys just added by the script)." [V] (jkp-ads, above)
  - The Excel-DNA WiX template also has a prompt-to-close-EXCEL custom action (`PromptToCloseProcesses=EXCEL`). [V] https://github.com/Excel-DNA/WiXInstaller
  - **Background `EXCEL.EXE` processes count too**, for example automation instances and preview handlers. [I]
- **`<ver>` is `16.0` for every version since Office 2016**: 2016, 2019, LTSC 2021/2024 and Microsoft 365. [O] This machine runs M365 Apps for enterprise, Click-to-Run, version 16.0.20430.20092, with keys under `16.0`.
  - Older keys (`14.0`, `15.0`) only matter for Office 2010/2013, which are long out of support.
- **The Excel-DNA alternatives** [V] (installing-your-add-in.md):
  - the Add-ins dialog;
  - VBA `Application.AddIns.Add(path).Installed = True`;
  - a VBScript that creates `Excel.Application` (the docs note "some sites ban vb script");
  - having the add-in **install itself on first open** from `AutoOpen` (`xlApp.AddIns.Add(xllPath, false).Installed = true`).
- **`AddIns.Add(FileName, CopyFile)`.** `CopyFile` "is ignored if the add-in file is on a hard disk". The method "does not install the new add-in. You must set the Installed property." [V] https://learn.microsoft.com/en-us/office/vba/api/excel.addins.add

### 1.2 Where to put the file
- **Default: `%APPDATA%\Microsoft\AddIns\`.**
  - It is Excel's own per-user add-ins folder, and the default Browse… location. [V] https://bettersolutions.com/excel/add-ins/local-copies.htm
  - It roams with the profile in VDI/Citrix setups, where `HKCU` also roams. [I]
  - Excel asks "Copy … to the Add-Ins folder?" only for add-ins browsed from non-local media. [V] (bettersolutions, above; Learn `AddIns.Add`)
- **It is NOT an Excel trusted location.**
  - [O] The owner's Excel trusted locations are `Office16\XLSTART`, `%APPDATA%\Microsoft\Excel\XLSTART`, `%APPDATA%\Microsoft\Templates`, `root\Templates`, `Office16\STARTUP` and `Office16\Library`.
  - [V] This matches Microsoft's published Excel defaults. The user `Addins` folder is a default trusted location only for PowerPoint. https://learn.microsoft.com/en-us/previous-versions/office/office-2010/cc179039(v=office.14)
  - So **MOTW must be removed from the file itself.** Installing into `AddIns` doesn't bypass the block.
- **`%LOCALAPPDATA%\Programs\Modelwright\`** is an alternative for the MSI and script paths. It doesn't roam: with a roaming profile, the `OPEN` key would follow the user to machines that don't have the file. Use it only as an `-InstallDir` option. [I]
- **Never install from a network share.**
  - Excel's block also covers files "opened from a network share". [V] support article in §2.
  - Shares reached by IP address or FQDN are often in the Internet zone. [V] https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked

### 1.3 Bitness, Office flavors and multiple versions
- **A 32-bit .xll can't load in 64-bit Excel, or the reverse.** Excel reports: "The file format and extension of '<name>.xll' don't match. The file could be corrupted or unsafe." [V] https://help.velixo.com/support/solutions/articles/153000016575-the-file-format-and-extension-of-velixoreportspro64-xll-don-t-match
  - The same message also appears when Software Restriction Policies or antivirus block the load. [V] (same page)
- **64-bit is the default install for Microsoft 365 / Office 2019+.** 32-bit remains for legacy 32-bit COM add-ins. [V] https://support.microsoft.com/en-us/office/choose-between-the-64-bit-or-32-bit-version-of-office-2dee7807-8f95-4d0c-b5fe-6c6f49b8d261
- **Detecting bitness** (best first):
  1. **Read the PE header of `EXCEL.EXE`.**
     - Take the path from `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe`.
     - The Machine field is `0x8664` for x64, `0x014C` for x86 and `0xAA64` for Arm64.
     - [O] This works on the owner's machine in Windows PowerShell 5.1 (`Get-Content -Encoding Byte -TotalCount`) and returns `8664`. PowerShell 7 needs `-AsByteStream` instead.
  2. **Click-to-Run:** `HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\Platform` = `x64`/`x86`. [O] Here it is `x64`. [V] https://help.pdq.com/hc/en-us/articles/360041142132-Determine-Office-365-2019-Bitness
  3. **MSI Office:** `HKLM\Software\Microsoft\Office\16.0\Outlook\Bitness`. This exists only if Outlook is installed.
     - The Excel-DNA WiX template relies on this key and **falls back to 32-bit when it is missing**. That fallback is a latent bug on 64-bit, Outlook-less machines. [V] (WiXInstaller `CustomAction.cs` `GetAddInName`)
- **Users can check by hand:** File › Account › About Excel. The first line ends in "32-bit" or "64-bit". [U, standard Excel UI; capture a screenshot]
- **Arm64.**
  - M365 Excel on Windows on Arm is Arm64EC, which is designed to load x64 add-ins. [V] https://office-watch.com/2021/how-will-microsoft-office-work-on-arm-devices/
  - 32-bit M365 on Arm gets security updates only until December 2026. [V] https://support.microsoft.com/en-us/office/lifecycle/end-of-support-for-32-bit-microsoft-365-apps-on-windows-arm-based-pcs
  - **Whether Excel-DNA + net48 loads correctly under Arm64EC is unverified.** [U] Document Arm as "untested".

### 1.4 Writing `install.ps1` / `uninstall.ps1` that work without admin
**Constraints that shape the script** [V]
- **Execution policy**
  - The Windows-client default is **Restricted**, meaning no scripts at all.
  - `powershell.exe -ExecutionPolicy Bypass -File …` sets the **Process** scope. That takes precedence over the CurrentUser and LocalMachine scopes, but **not over a Group Policy setting** (MachinePolicy/UserPolicy).
  - Under **RemoteSigned**, an unsigned script that carries MOTW doesn't run until it is unblocked.
  - `Invoke-WebRequest`, `Invoke-RestMethod` and `curl.exe` don't add MOTW.
  - https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies
- **AppLocker / App Control**
  - When AppLocker script rules block a `.ps1`, **PowerShell still runs it, in ConstrainedLanguage mode** (CLM).
  - `.bat`/`.cmd` files fall under the same script rule collection. By default, users may only run scripts from `%windir%` and `%programfiles%`. https://learn.microsoft.com/en-us/windows/security/application-security/application-control/app-control-for-business/applocker/script-rules-in-applocker
  - In CLM, every built-in cmdlet works. `Add-Type` with arbitrary C#, most .NET types, and all COM except three are blocked. https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_language_modes
- **Therefore**
  - Target **Windows PowerShell 5.1**, which every Windows 10/11 has.
  - Use **only cmdlets**: `Get-Process`, `Copy-Item`, `Unblock-File`, `Get-ItemProperty`, `New-ItemProperty`, `Remove-ItemProperty`, `Get-Content -Encoding Byte`, `Get-FileHash`.
  - Use **no** `[Microsoft.Win32.Registry]`, `[IO.File]`, `Add-Type` or `New-Object -ComObject Excel.Application`.

**`install.ps1` steps**
1. If any `EXCEL` process is running, stop with: "Close Excel completely (check the taskbar and Task Manager), then run this again." Optionally offer to wait and retry.
2. Detect Excel bitness (§1.3), with `-Bitness 32|64` as an override. Refuse Arm64 32-bit and unknown results, with a clear message.
3. Check that .NET Framework 4.8 is present: `HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full` `Release` ≥ 528040. [V] Release values: https://learn.microsoft.com/en-us/dotnet/framework/install/versions-and-dependencies
   - [O] The owner's machine reports 533509 (4.8.1).
4. Create the folder if needed, then `Copy-Item` `Modelwright64.xll` or `Modelwright32.xll` into `%APPDATA%\Microsoft\AddIns\`. Then **`Unblock-File`** on the copied file. Always unblock: this makes the script immune to the MOTW block.
5. Optionally `Get-FileHash -Algorithm SHA256` the source file and compare it with `SHA256SUMS.txt` in the same folder.
6. Register the add-in under `HKCU:\Software\Microsoft\Office\16.0\Excel\Options`:
   - Create the key only `if (-not (Test-Path …))`. **Never use `New-Item -Force` on an existing registry key: it replaces the key and wipes its values.** [U, widely reported PowerShell behavior; test it]
   - Read the existing `OPEN*` values and sort them by their numeric suffix.
   - Remove any value whose path ends in `\Modelwright32.xll` or `\Modelwright64.xll`, which covers re-installs and bitness switches. Compact the rest so there are no gaps.
   - Append `/R "<full path>"` as the next `OPENn`.
   - Remove any `Add-in Manager` value that names our file.
7. Print "Done. Start Excel; you should see the **Modelwright** tab." Use `Read-Host` unless `-Quiet`, so the window doesn't vanish.

**`uninstall.ps1` steps**
1. Require Excel to be closed.
2. Remove our `OPEN` values and compact the numbering. The Excel-DNA WiX custom action does the same: it deletes every `OPEN*` value and rewrites the survivors as `OPEN`, `OPEN1`…. [V] WiXInstaller `CaUnRegisterAddIn`
   - Sort numerically. The template keeps registry enumeration order, which isn't guaranteed. [I]
3. Remove `Add-in Manager` entries, then delete the .xll files.
4. Keep `%APPDATA%\<product>\settings.json` unless `-RemoveSettings` is passed.
   - Today the folder is `ModelingToolkit`, per `SettingsStore.cs`. The log is in `%LOCALAPPDATA%\ModelingToolkit\log.txt`.
   - The rename needs a settings-migration decision.

**How users launch it**
- **Recommended:** right-click `install.ps1` › **Run with PowerShell**.
  - On Windows 11 this is under "Show more options". [U]
  - That shell verb runs PowerShell with Process-scope Bypass unless the policy is AllSigned. [U, from memory of the default registry verb; verify on Windows 11 and capture a screenshot]
- **Fallback:** paste one line into PowerShell: `powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\Downloads\Modelwright\install.ps1"`.
- **Optional `Install.cmd` wrapper** (`powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"`).
  - Convenient, but AppLocker script rules block `.cmd` outright, not just constrain it. [V]
  - A `.cmd` downloaded with MOTW also shows Windows' "Open File – Security Warning". [U, capture it]

**Testing CLM safety**
- Run the scripts in a session with `$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'` (Learn notes this is "useful for experimenting"). Make that a CI check.

### 1.5 What happens under strict policy
| Control | Effect on Modelwright | Source |
|---|---|---|
| Execution policy set by **GPO** to Restricted or AllSigned | `-ExecutionPolicy Bypass` is ignored. Restricted: the script can't run, so use the manual path. AllSigned: a *signed* script can run after a publisher prompt. | [V] about_Execution_Policies |
| **AppLocker script rules** | A `.ps1` from Downloads runs, but in CLM, so the script must be CLM-safe (above). `.cmd` is blocked. | [V] AppLocker script rules |
| **AppLocker DLL rules** (off by default; when on, the defaults allow only `%windir%` and `%programfiles%`) | A per-user .xll in `%APPDATA%` would be blocked unless IT adds a path or **publisher** rule. Publisher rules need a signed file. Learn says the DLL collection covers ".dll and .ocx"; **whether AppLocker treats an `.xll` as a DLL by extension or by image load is unverified.** | [V] https://learn.microsoft.com/en-us/windows/security/application-security/application-control/app-control-for-business/applocker/dll-rules-in-applocker · [U] |
| **AppLocker Windows Installer rules** (defaults) | Everyone may run **all digitally signed** `.msi` files. An unsigned MSI is blocked for standard users. | [V] https://learn.microsoft.com/en-us/windows/security/application-security/application-control/app-control-for-business/applocker/windows-installer-rules-in-applocker |
| **WDAC / App Control for Business**, **Smart App Control** | Smart App Control checks "all code (DLLs, scripts, etc.) loaded by the Windows OS Loader". Unsigned code loads only with good cloud reputation. Smart App Control is on mainly for consumer clean installs, not managed enterprise devices. | [V] https://textslashplain.com/2026/04/28/smart-app-control/ · https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation |
| Office GPO **"Require that application add-ins are signed by Trusted Publisher"** | This is in Microsoft's **M365 Apps security baseline** (expected value: Enabled). Excel checks the signature before loading. Unsigned or untrusted add-ins are disabled with a notification. | [V] https://www.tenable.com/audits/items/MSCT_Microsoft_365_Apps_for_enterprise_2312_v1.0.0.audit:fd41beaecb14e039edf64b0d0045a6de |
| Office GPO **"Disable all application add-ins"**, or "Allow mix of policy and user locations" disabled | Nothing a user can do; IT must allow it. | [V] cc179039 (above) |
| Registry editing blocked | Only `regedit`/`reg.exe` are blocked. The script's cmdlets and Excel's own Add-ins dialog still write HKCU. [U] | — |

**Upshot:** in the most locked-down shops, no self-service path works. The deliverable for them is an **IT pack**: a signed .xll, its SHA-256, a publisher certificate or thumbprint to add to Trusted Publishers or AppLocker, and build-from-source instructions.

---

## 2. Mark-of-the-Web: "Excel is blocking untrusted XLL add-ins"

### 2.1 The policy
- **Timeline** [V]
  - Announced to Insiders on 2023-02-28 ("Block untrusted XLL add-ins by default"), Version 2302 (Build 16130.20128) or later.
  - Message Center MC524212 rolled it out to Current, Monthly Enterprise and Semi-Annual Enterprise channels in March 2023.
  - https://techcommunity.microsoft.com/blog/microsoft365insiderblog/block-untrusted-xll-add-ins-by-default/4218353 · https://app.cloudscout.one/evergreen-item/mc524212/
- **Products affected:** Excel for Microsoft 365, 2024, 2021, 2019 and 2016. [V] https://support.microsoft.com/en-us/excel/excel-is-blocking-untrusted-xll-add-ins-by-default
- **Scope:** "This setting change only blocks XLL add-ins - not all add-ins. For example, add-ins from the Store, are not impacted." [V] Insider blog.
- **What triggers it:** MOTW, the `Zone.Identifier` alternate data stream with `ZoneId=3` (Internet) or `4` (Restricted). [V] https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked
  - Browsers and email add MOTW.
  - MOTW exists only on NTFS.
  - Files brought down by the OneDrive sync client or "Open in Desktop App" have none.
  - **Extracting a MOTW-tagged .zip with Windows Explorer copies MOTW onto the extracted files.** [U, long-standing Windows behavior, not found in an official doc in this session; other archivers vary]
  - Excel-DNA's docs: "As of March 2023, Microsoft is testing an option whereby downloaded addins are blocked by default." https://github.com/Excel-DNA/Excel-DNA.github.io/commit/531372f51dae26cc25842aea56a23c201423f268

### 2.2 What the user sees
- **The dialog** [V] (Insider blog)
  - Title: **"Microsoft Excel Security Notice"**.
  - One button: **"Leave this add-in disabled"**.
  - A **"More Information"** link that opens the support article.
  - There is no Enable button.
- **The body text is not quoted in any source found.** [U] Capture the dialog verbatim for the README.
- **Related messages to capture as screenshots:**
  - Bitness mismatch: "The file format and extension of '…xll' don't match. The file could be corrupted or unsafe." [V]
  - A signed add-in when signatures are required: "The digital signature is valid, but the signature is from a publisher whom you have not yet chosen to trust." / "The application add-in has been disabled." with **"Enable all code published by this publisher"**. [V, third-party docs] https://help.velixo.com/doc/main/classic/notice-the-signature-is-from-a-publisher-whom-you-
  - Older Office security notices (outside the MOTW block) offered "Enable this add-in for this session only" / "Do not enable". [V, third-party] https://support.sou.edu/kb/articles/pdf/warnings-in-microsoft-office-apps-about-macros-and-add-ins
  - When an add-in is opened by double-click or File › Open, rather than through the Add-ins dialog, Excel may show a security notice even for an unblocked, unsigned .xll. [U] Capture it.

### 2.3 Controls that govern it
- **Per-file:** Properties › General › tick **Unblock** › OK, or PowerShell `Unblock-File`. Both do the same thing. [V] support article; internet-macros-blocked.
- **Trusted Location:** File › Options › Trust Center › Trust Center Settings… › Trusted Locations › Add new location…. [V] support article.
- **Trusted Sites** (Internet Options › Security › Trusted Sites › Sites) for a site or file server. [V] support article.
- **Trusted Publisher.** "add them as a trusted publisher to unblock this add-in and any future add-ins the publisher signs with the same certificate. First, you'll need to unblock this particular file." [V] support article.
  - **Signing alone does NOT bypass the MOTW block.** Only a publisher already in the Trusted Publishers store does.
  - Contrast: for `.xla`/`.xlam`, Microsoft says signing plus trusting the publisher "doesn't work for Excel Add-in files that have Mark of the Web" (since MS16-088). Whether the XLL path really honors Trusted Publisher for a MOTW file is **unverified**. [U] Test it once there is a signature.
- **Warn instead of block** (meant to be temporary): `HKCU\SOFTWARE\Microsoft\Office\16.0\Excel\Security`, DWORD **`BlockXLLFromInternet` = 0**, then fully restart Excel. [V] support article.
  - **GPO:** Administrative Templates › Microsoft Excel 2016 › Excel Options › Security › Trust Center › **"Block Excel XLL Add-ins that come from an untrusted source"**. [V] MC524212.
  - **Don't tell users to use this.** It weakens their security for every .xll.
- **Default add-in signature checks:** "If you disable or do not configure this policy setting, Office … do not check the digital signature on application add-ins before opening them." So by default an *unblocked* unsigned .xll loads silently from the Add-ins dialog. [V] Tenable/MSCT description (§1.5).

### 2.4 SmartScreen and browsers
- **SmartScreen application reputation** [V] https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation
  - It gates *running* downloaded files.
  - Unsigned files get "Windows protected your PC", and the user must choose "Run anyway". Enterprise policy can remove that option.
  - Signed files show the verified publisher name until reputation builds.
  - EV no longer bypasses SmartScreen.
  - This applies to an MSI or EXE. An .xll isn't "run" by Explorer; it's loaded by Excel. [I]
- **Chromium download classification** (Chrome; Edge is Chromium-based) [V] https://chromium.googlesource.com/chromium/src/+/main/components/safe_browsing/content/resources/download_file_types.asciipb
  - `.xll`, `.msi` and `.ps1`: `danger_level: ALLOW_ON_USER_GESTURE`, `FULL_PING` on Windows.
  - `.zip`: an archive whose contents are inspected, with no default danger level.
  - So a zip is the gentlest download. Edge can still say "*isn't commonly downloaded*" for new, unknown files. [U, capture Edge's Keep › Show more › Keep anyway flow]

### 2.5 What a non-technical user clicks (manual path)
1. On the repo's **Releases** page, under **Assets**, click `Modelwright-<ver>.zip`.
2. If the browser warns, choose **Keep**. [U, capture]
3. In **Downloads**, right-click the zip › **Properties** › tick **Unblock** › **OK**. If there is no Unblock box, the file wasn't marked, which is fine.
4. Right-click › **Extract All…** › **Extract**.
5. Continue with the manual steps (§10) or the script.

**Recovering after a block:** close Excel, unblock the .xll in `%APPDATA%\Microsoft\AddIns` (or re-run `install.ps1`), and restart Excel.
- If Excel has put the add-in in **Disabled Items**: File › Options › Add-ins › Manage: **Disabled Items** › Go… › select it › **Enable**. [U, capture]

---

## 3. Code signing: options, cost, mechanics (owner decision)

### 3.1 What signing buys

| Effect | Unsigned | Signed (OV or Artifact Signing) |
|---|---|---|
| Excel MOTW XLL block | Blocked; unblock required | **Still blocked**; unblock required. Only a *pre-trusted* publisher skips it. [V] |
| Default Office settings, unblocked file | Loads silently | Loads silently [V] |
| M365 security baseline (`requireaddinsig` on) | **Never loads.** The user has no recourse. | Loads after a one-time "Enable all code published by this publisher", unless policy also blocks user-store trusted publishers. Or IT pushes the certificate. [V]/[I] |
| AppLocker / WDAC | Only a path or hash rule can allow it | A publisher rule can allow it [V] |
| Smart App Control | Needs cloud reputation | Signed passes [V] |
| SmartScreen (for MSI/EXE) | "Windows protected your PC" → Run anyway | Publisher shown; reputation builds over time [V] |
| Antivirus false positives | Higher risk | "Some indications that signing … helps" (Govert van Drimmelen, Excel-DNA) [V] §7 |
| Signed MSI under AppLocker default rules | Blocked for standard users | Allowed [V] |

### 3.2 Options

| Option | Cost | Eligibility | Notes |
|---|---|---|---|
| **Azure Artifact Signing** (formerly Trusted Signing; GA January 2026) | Basic **$9.99/month**: 5,000 signatures, 1 certificate profile. Premium $99.99/month: 100,000 signatures, 10 profiles. Overage $0.005 per signature. Billed for the full month, not pro-rated. Needs a **paid** Azure subscription. | **Public Trust:** organizations in the US, CA, EU, UK, AU, NZ, JP, KR, SG, CH, NO and IL. **Individuals: US/CA only.** Identity validation takes 1–20 business days. | Certificates are **short-lived** (3-day validity, renewed daily) and cannot be exported. Timestamping is mandatory. **The "3+ years of verifiable history" rule for orgs** appears in a 2025 Microsoft Q&A moderator answer, but **not** in today's Learn quickstart or FAQ. Treat it as possibly still enforced. **[U]** Individual validation (US/CA) needs no business history: it uses a government ID through Microsoft Verified ID. No EV, "no plan to issue EV". |
| **OV certificate** (DigiCert, Sectigo, GlobalSign…) | **$150–300/year** | Worldwide | Since June 2023, the key must sit on an HSM or hardware token (a cloud HSM option exists). The certificate is the same for 1–3 years. |
| **EV certificate** | $400+/year | Worldwide | **No SmartScreen advantage since 2024.** Not recommended. |
| **SignPath Foundation** | **Free** for OSS | OSI license with no commercial dual-licensing; no proprietary components; actively maintained; **already released**; reproducible CI build with origin verification; MFA; a published code-signing policy; an approver for each signing request. | **The certificate is issued to "SignPath Foundation"**, which appears as the publisher, not Pegasus Technology Group. MIT, Excel-DNA (zlib), XLParser (MPL-2.0) and Irony (MIT) are all OSI. |
| **Unsigned** | $0 | — | Works for most individuals and small firms. Dead on baseline-hardened enterprise desktops. |

**Sources:**
- https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/code-signing-options
- https://learn.microsoft.com/en-us/azure/artifact-signing/quickstart
- https://learn.microsoft.com/en-us/azure/artifact-signing/faq
- https://www.devclass.com/security/2026/01/14/code-signing-windows-apps-may-be-easier-and-more-secure-with-new-azure-artifact-service/4079554 (pricing, GA)
- https://learn.microsoft.com/en-us/answers/questions/5532414/trusted-signing-identity-verification-failed-provi (3-year claim)
- https://learn.microsoft.com/en-us/answers/questions/2202247/azure-trusted-signing-short-lived-certificate (3-day validity)
- https://signpath.org/terms

### 3.3 The Trusted Publisher catch with Artifact Signing [I]
- Office trusts a publisher by adding **the signing certificate** to the Trusted Publishers store. "All macros validly signed with the same certificate are recognized…". [V] https://learn.microsoft.com/en-us/microsoft-365-apps/security/trusted-publisher
- Artifact Signing issues a new short-lived leaf certificate every day. So a user or IT trust granted for v0.1.0 would **not** carry over to v0.2.0. Each release would need re-trusting.
- An OV certificate keeps the same leaf for its whole 1–3-year life, so trust carries across releases.
- This matters only in environments that enforce `requireaddinsig`. It is a real reason to prefer OV if enterprise users are a target. Confirm it with a two-release test before deciding. [U]

### 3.4 Mechanics
- **Signing order:** **pack, then sign.** Excel-DNA's `PackExcelAddIn` task runs `signtool sign <ExcelAddInSignOptions> "<packed.xll>"` *after* a successful pack, when `ExcelAddInSignTool` and `ExcelAddInSignOptions` are set. [V] https://github.com/Excel-DNA/ExcelDna/blob/master/Source/ExcelDna.AddIn.Tasks/PackExcelAddIn.cs · SDK properties: https://github.com/Excel-DNA/docs.excel-dna.net/blob/gh-pages/sdk-style-project-properties.md
  - Any later change to the file breaks the signature, so "Do not modify signed files". [V] SmartScreen doc.
- **Excel-DNA's own advice:** "sign the packed .xll (using the signtool utility) and incorporate the certificate into the Office Trust Center". [V] installing-your-add-in.md
- **In CI, sign as a separate step rather than through MSBuild.** Use `Azure/artifact-signing-action@v2` (current tag), with `files:` taking explicit paths so `.xll` works. Then `azure/login@v3` with OIDC.
  - Set `file-digest: SHA256`, `timestamp-rfc3161: http://timestamp.acs.microsoft.com` and `timestamp-digest: SHA256`. [V] https://github.com/Azure/artifact-signing-action
  - For an OV certificate on a cloud HSM, use the vendor's signtool KSP or CLI.
- **Full order:** build → pack (32 and 64) → **sign the .xll files** → zip → build the MSI around the *signed* .xll files → **sign the MSI** → (optionally sign `install.ps1`/`uninstall.ps1`) → SHA-256 → attest → publish.
- **Verify a signature:** `Get-AuthenticodeSignature .\Modelwright64.xll` (CLM-safe), or `signtool verify /pa /v`.

**Owner decision needed (D-sign):** (a) unsigned v0.1.0, then SignPath once released; (b) Artifact Signing; (c) OV certificate. Main facts:
- Cost, and Artifact Signing's per-release trust.
- Whether "SignPath Foundation" as the displayed publisher is acceptable.
- The LLC's age, if the 3-year rule applies.

---

## 4. MSI and how others ship

### 4.1 Tooling
- **Excel-DNA WiXInstaller template** [V] https://github.com/Excel-DNA/WiXInstaller
  - WiX v4 SDK-style, last pushed 2025-07, "fledgling".
  - `Scope="perUser"`; installs to `%APPDATA%\<Manufacturer>\<Product>`.
  - Carries both `…32.xll` and `…64.xll`, and a C# DTF custom action that writes and compacts `OPEN` values for Office 11.0–16.0.
  - Has a close-Excel prompt and a .NET 4.0 launch condition.
  - A per-machine variant uses Active Setup: https://github.com/bpatra/ExcelDNAWixInstallerLM
  - **Weaknesses:** Outlook-based bitness detection with a 32-bit fallback (§1.3), and a .NET 4.0 check where we need 4.8.
- **WiX today:** v7.0.0 (2026-04-06). [V] https://github.com/wixtoolset/wix/releases
  - **WiX now uses the Open Source Maintenance Fee.** Users of the *binary releases* who use WiX "as part of revenue-generating activities" with **annual gross revenue ≥ US$10,000** must sponsor. Self-compiling WiX from source is exempt. [V] https://github.com/wixtoolset/wix/blob/main/OSMFEULA.txt
  - Whether a free OSS add-in published by Pegasus counts as "revenue-generating" is a **legal and owner call**.
- **Inno Setup 6.5+.** `PrivilegesRequired=lowest` supports per-user installs. [U, standard Inno directive]
  - It now *requests* (does not require) a commercial license from for-profit organizations with revenue over US$5,000. [V] https://jrsoftware.org/isorder.php
  - An Inno EXE is not an MSI, so AppLocker's "signed MSI" default rule doesn't help it, and EXE rules apply instead. [I]
- **Advanced Installer** (commercial) has an Excel-DNA guide linked from Excel-DNA's docs. [V]
- **MSI policy pitfalls:**
  - The MSI must be **signed** to pass AppLocker's default Windows Installer rules for non-admins. [V]
  - A per-user MSI writes only to HKCU and the profile. It can't reach Program Files, so AppLocker DLL default rules (§1.5) still bite. [I]

### 4.2 How comparable add-ins ship
| Product | Technology | Install model | Evidence |
|---|---|---|---|
| **Macabacus** 9.9.5 | VSTO COM add-in (`Macabacus.Excel.vsto\|vstolocal`), plus `Macabacus.xlam` registered as `OPEN` | **Per-machine MSI.** Uninstall entry under `HKLM\…\WOW6432Node` (`MsiExec.exe /X{B6E1…}`); files in `C:\Program Files (x86)\Macabacus\Macabacus 2016`; COM registration in HKLM (both views); the `.xlam` in `C:\ProgramData\Macabacus\Macabacus 2016`. Its help center tells users without rights to get an IT "over-the-shoulder" install. | [O]; [V] https://help.macabacus.com/article/324-basic-installation (summary via search; the page now needs login) |
| **Daloopa** | Office.js | **AppSource / M365 admin center only.** "There's no separate program to download"; no admin rights for individuals. | [V] https://docs.daloopa.com/docs/install-daloopa-excel-add-in |
| **S&P Capital IQ** Office plug-in | Windows installer (EXE) | University IT guides mention admin passwords or elevation prompts during install. | [V, third-party] https://guides.library.duke.edu/capitaliq-guide/capiq-tools · https://library.answers.nyu.edu/faq/392635 |
| **FactSet** Office add-in | Windows installer | **Not verified.** The FactSet "Office Lite Add-in Installation Guide" PDF couldn't be read here. | [U] |

**Takeaway:** the incumbents ship **admin MSIs or EXEs**, or Office.js via the store. A no-admin path is a genuine differentiator, and it's what the owner asked for.

### 4.3 MSIX: not feasible for an .xll [V]/[I]
- MSIX copies HKCU writes "on write to a per-user, per-app private location". Its `registry.dat` covers only virtual `HKLM\Software`. https://learn.microsoft.com/en-us/windows/msix/desktop/desktop-to-uwp-behind-the-scenes
- So Excel, running outside the package, would never see an `OPEN` value created by the package.
- Microsoft has said Office extensibility in MSIX is future work: https://techcommunity.microsoft.com/discussions/msix-discussions/microsoft-office-add-in-deployment-using-msix/2079547
- **Drop MSIX.**

---

## 5. Package managers

| Manager | Feasible for an .xll? | Facts |
|---|---|---|
| **winget** (`microsoft/winget-pkgs`) | **Only through an MSI or EXE installer.** | **Allowed installer types:** msix, msi, appx, exe, inno, nullsoft, wix, burn and zip. zip+`portable` is for executables, which an .xll is not. **Requirements:** a public **HTTPS** `InstallerUrl` straight from the publisher (GitHub Releases qualifies), plus `InstallerSha256`. The installer must support silent mode, work for admins *and* non-admins, and pass antivirus scanning (`Binary-Validation-Error`, `Validation-Defender-Error`). **Signing is not listed as a requirement.** A package that "does not install an application" may hit `Validation-Executable-Error` and need a reviewer comment. [V] https://learn.microsoft.com/en-us/windows/package-manager/package/repository · https://learn.microsoft.com/en-us/windows/package-manager/package/manifest |
| **Chocolatey** | Possible, poor fit | Admin is the norm; "non-administrative installation should be a last resort". Needs a non-Restricted execution policy. Community packages are moderated. [V] https://docs.chocolatey.org/en-us/choco/setup/ |
| **Scoop** | Technically yes | Per-user, no UAC. A custom bucket manifest could `post_install` the `OPEN` registration and `pre_uninstall` its removal. 32/64-bit URLs are supported. Rarely allowed on corporate desktops. [V] https://github.com/ScoopInstaller/Scoop · https://github.com/ScoopInstaller/Scoop/wiki/App-Manifests |

**Recommendation:** none for v0.1.0. Add winget once a signed per-user MSI exists and download volume justifies it.

---

## 6. Build from source, release pipeline, and both bitnesses

### 6.1 Prerequisites (today's repo)
- **Windows 10/11.** Packing uses Windows resource APIs by default (`ExcelDnaPackManagedResourcePackingOnWindows=false`). [V] SDK props.
- **.NET SDK 10.0.100 or a later feature band**, per `global.json` (`rollForward: latestFeature`). [O] Building `net48` needs no targeting pack, because `Microsoft.NETFramework.ReferenceAssemblies` is referenced.
  - The *tests* also need the .NET 8 runtime, since they target net8.0 and net48.
  - The SDK installs **without admin** through `dotnet-install.ps1` into the user profile. [U, standard Microsoft script; cite https://learn.microsoft.com/en-us/dotnet/core/tools/dotnet-install-script when writing the README]
- **NuGet access:** nuget.org or the company's proxy (Artifactory and similar) for ExcelDna.AddIn 1.9.0, XLParser 1.7.5, Irony and the reference assemblies.
- **Command:** `dotnet build src\ExcelModelingToolkit.AddIn -c Release`. **`dotnet publish` is not needed**: the pack runs at build and writes `bin\Release\net48\publish\*.xll` (`RunExcelDnaPack=true`, `ExcelDnaPublishPath` defaults to `publish`). [V]/[O]

### 6.2 Both bitnesses
- **Settings** [V] SDK props:
  - `ExcelDnaCreate32BitAddIn` (default true; the repo sets it **false**) and `ExcelDnaCreate64BitAddIn`.
  - Suffixes: 32-bit has none by default; 64-bit defaults to `64`. Packed outputs get `-packed`.
  - Explicit names: `ExcelDnaPack32BitXllName` / `ExcelDnaPack64BitXllName`.
- **Change:** set `ExcelDnaCreate32BitAddIn=true`, and either name the outputs `Modelwright32.xll` / `Modelwright64.xll` or rename them in the release job.
- **Excel-DNA's old packing doc** describes the same pairing ("`MyAddInFinal.xll` … `MyAddInFinal64.xll`"). [V] exceldna-packing-tool.md
- **net48 is fine for both.**
  - Managed code is AnyCPU, and `Core` is netstandard2.0.
  - Excel-DNA 1.9 supports .NET Framework 4.6.2–4.8.1 and .NET 6+ (NuGet page). The 1.9 release notes say "4.7.2+". https://www.nuget.org/packages/ExcelDna.AddIn/1.9.0 · https://github.com/Excel-DNA/ExcelDna/releases/tag/v1.9.0-Release
  - **Risk to test:** P/Invoke code (`TraceKeyHook`, `UndoKeyHook`, `KeyboardLayout`, `SettingsDialog`) must be pointer-size-neutral. For example `new IntPtr(Convert.ToInt64(hwnd))` in `ExcelEvents.cs` throws on 32-bit if the value exceeds 2³¹. [I]
  - **32-bit testing needs a separate machine or VM**, because 32-bit and 64-bit Office can't be installed side by side. [I] https://support.microsoft.com/en-us/office/choose-between-the-64-bit-or-32-bit-version-of-office-2dee7807-8f95-4d0c-b5fe-6c6f49b8d261

### 6.3 Release workflow (sketch, not yet written)
`.github/workflows/release.yml`, triggered on tag `v*`, `runs-on: windows-latest`, with permissions `contents: write`, `id-token: write` and `attestations: write`.
1. `actions/checkout@v7`, then `actions/setup-dotnet@v6` with SDK 10.0.x and 8.0.x.
2. `dotnet test …` (the existing CI gates).
3. `dotnet build src\ExcelModelingToolkit.AddIn -c Release -p:ExcelDnaCreate32BitAddIn=true -p:ContinuousIntegrationBuild=true`.
4. Rename the outputs to `Modelwright32.xll` / `Modelwright64.xll` in `dist\`.
5. *(If D-sign allows it)* `azure/login@v3` (OIDC), then `Azure/artifact-signing-action@v2` with `files:` set to both .xll files, then `Get-AuthenticodeSignature` to assert `Valid`.
6. Stage the zip contents and run `Compress-Archive`.
7. *(Optional)* build the WiX per-user MSI from the signed .xll files, then sign the MSI.
8. `Get-FileHash -Algorithm SHA256 dist\* | …` → `SHA256SUMS.txt`, in `sha256sum` format: `<hex>  <file>`.
9. `actions/attest-build-provenance@v4` (current: v4.2.2; new projects may use `actions/attest`) on `dist\*`. [V] https://docs.github.com/en/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations
10. `gh release create $TAG dist\* --verify-tag --notes-file …`. Turn on **immutable releases** in the repo settings. [V] https://github.blog/changelog/2025-10-28-immutable-releases-are-now-generally-available/

GitHub also shows a **SHA-256 digest beside every release asset** automatically, in the UI, the API and `gh`. [V] https://github.blog/changelog/2025-06-03-releases-now-expose-digests-for-release-assets/

### 6.4 Verification commands for IT
- `Get-FileHash .\Modelwright64.xll -Algorithm SHA256`, or `certutil -hashfile Modelwright64.xll SHA256` where PowerShell is blocked. Compare with `SHA256SUMS.txt` or the digest shown on GitHub.
- `gh attestation verify .\Modelwright64.xll -R <owner>/<repo>` proves which workflow run and commit built the file. [V] GitHub docs.
- `Get-AuthenticodeSignature .\Modelwright64.xll` shows the signer and the timestamp.

### 6.5 Reproducibility: the build is not bit-for-bit today [O]/[I]
- `StampVersion` embeds `DateTime.UtcNow` as `BuildDate`, so any two builds on different days differ. **Fix:** use the commit date (`git log -1 --format=%cs`) or `SOURCE_DATE_EPOCH`.
- There are **no `packages.lock.json` files.** Enable `RestorePackagesWithLockFile` and use `--locked-mode` in CI.
- Excel-DNA packing is compressed and multithreaded (`ExcelDnaPackRunMultithreaded=true`). Whether its output is deterministic is **[U]**. Add a CI job that builds twice and compares hashes, and set `ExcelDnaPackRunMultithreaded=false` if needed.
- Signatures and timestamps make signed files differ from any rebuild. IT compares a rebuilt *unsigned* .xll with the release .xll after `signtool remove /s`, or relies on the attestation. [I]

---

## 7. Antivirus and false positives

- **History** [V] https://excel-dna.net/blog/2022/03/07/excel-dna-1.6-.net6-packagereference-anti-virus/ · https://github.com/Excel-DNA/ExcelDna/issues/403 (2021-08, open) · https://groups.google.com/g/exceldna/c/argkRiTFH_0
  - "Some malicious Excel add-ins … built using Excel-DNA" trained antivirus heuristics to flag all Excel-DNA add-ins. The main signal is "executable assemblies as resources in the .xll".
  - Examples: `MSIL/TrojanDropper.Agent.FGU`, and Defender flagging ExcelDna.Integration 1.5.
- **Excel-DNA's mitigations (1.6+)** [V]
  - Packed resources are encoded.
  - `ExcelDnaPackCompressResources=false` turns compression off.
  - "Signing the final add-in library helps reduce false positive detections."
  - "Report the false positives to [the] vendor."
  - Govert also suggested not packing at all when distribution doesn't need it (`RunExcelDnaPack=false`). https://groups.google.com/g/ExcelDna/c/rJOUusWj27U
- **Our extra risk signals** [I, hypotheses]:
  - **thread keyboard hooks** (`SetWindowsHookEx` in `UndoKeyHook` and `TraceKeyHook`);
  - **in-memory assembly loading**, which .NET Framework 4.8 passes to AMSI for scanning.
- **Practice**
  - **Before each release:** scan locally with Defender (`MpCmdRun -Scan -ScanType 3 -File …`), then upload each asset to **VirusTotal**. VirusTotal uploads are shared with vendors, which is fine for OSS. Record the report links in the release notes.
  - **Exit criterion** (PLAN Phase 5): "0 detections from major engines". Expect occasional ML or heuristic hits from small vendors.
  - **False positives:** submit at https://www.microsoft.com/wdsi/filesubmission (developer submission, "incorrectly detected"), and to each flagging vendor. That portal is also the route for SmartScreen and winget issues. [V] SmartScreen and winget docs.
  - **Mitigations, in order:** sign; keep the publisher identity consistent; avoid unusual packing options; compare packed and unpacked detection rates if a hit appears.

---

## 8. Macabacus coexistence

- **Who owns a key** [V] research/05; https://learn.microsoft.com/en-us/office/vba/api/excel.application.onkey
  - Both add-ins bind through `Application.OnKey` / `xlcOnKey`. **Whoever registers last wins.**
  - Macabacus's **Override** button just re-registers its keys.
  - Our ribbon **Re-register shortcuts** (`EmtReregisterKeys`) does the same in the other direction. [O] `Commands.cs`, `ToolkitRibbon.cs`
- **Startup order**
  - Add-ins-dialog add-ins (`OPEN`, `OPEN1`… in numeric order) are reported to load **before** COM add-ins, which load alphabetically. https://bettersolutions.com/vba/add-ins/user-faqs.htm (search summary) **[U, not Microsoft-documented]**
  - Macabacus is a VSTO **COM** add-in, plus an `.xlam` in `OPEN` [O]. So **Macabacus most likely binds last and wins at startup** on shared keys. [I]
- **Unloading Modelwright mid-session** (PLAN §6) restores *Excel's* defaults, which also strips Macabacus's bindings until Macabacus re-registers them.
- **Ctrl+Z / Ctrl+Y** use our thread keyboard hook, not `OnKey`. Whether Macabacus hooks the same keys, and how the two hooks chain, is **[U]**. It belongs on the Phase 5 test list.

**Documentation to ship (README "Using Modelwright alongside Macabacus")**
1. "Both add-ins use the same shortcuts. Whichever registered its shortcuts most recently gets the key. Run only one of them at a time."
2. **Turn Macabacus off:** File › Options › Add-ins › Manage: **COM Add-ins** › **Go…** › untick **Macabacus COM Add-in** › OK.
   - This persists until you tick it again. Macabacus is registered per machine; a standard user can untick it, and Office records that per user. [U, capture as a standard user]
3. **Turn Modelwright off:** File › Options › Add-ins › Manage: **Excel Add-ins** › **Go…** › untick **Modelwright** › OK. This persists.
4. **For this session only:**
   - To let Modelwright win: click Modelwright › **Re-register shortcuts**.
   - To let Macabacus win: click its **Override**.
5. **Start Excel with no add-ins at all:** hold **Ctrl** while starting Excel and confirm Safe Mode. [U, standard Excel behavior]

**Suggested product change (not research):** a ribbon toggle **"Shortcuts on/off (this session)"**, so users don't need the persistent Add-ins dialog.

---

## 9. Office versions and platforms

| Version | Status (2026-10-08) | Modelwright |
|---|---|---|
| Microsoft 365 Apps, Current / Monthly Enterprise | Supported | **Primary target.** Test on Current Channel (PLAN). |
| Office LTSC 2024 / Office 2024 | Mainstream support to **2029-10-10** [V] https://learn.microsoft.com/en-us/lifecycle/products/office-ltsc-2024 | **Test** (PLAN). |
| Office LTSC 2021 | Ends **2026-10-14** [V] https://learn.microsoft.com/en-us/lifecycle/products/office-ltsc-2021 | Expected to work; don't test. Say "best effort". |
| Office 2016 / 2019 | **Ended 2025-10-14** (Excel 2019 extended end 2025-10-15) [V] https://learn.microsoft.com/en-us/lifecycle/products/excel-2019 | Probably works (Excel-DNA supports 2007+); **unsupported**. |
| 32-bit Excel (any of the above) | Supported by Microsoft | **Ship the 32-bit .xll** (§6.2); test once. |
| Windows on Arm (M365 64-bit, Arm64EC) | x64 add-ins are designed to load | **Untested** [U]. |
| **Excel for Mac** | — | **Not supported.** An .xll is a Windows DLL, and Excel-DNA is Windows-only. Say so plainly. |
| **Excel for the web** | — | **Not supported.** It doesn't load .xll add-ins. A **colors-only Office.js edition is on the v2 roadmap** (PLAN). |

- **Excel-DNA 1.9.0 (2025-11-07)** [V] https://www.nuget.org/packages/ExcelDna.AddIn/1.9.0
  - "Excel versions 2007 through 2022 / Office 365 can be targeted with a single add-in."
  - .NET Framework 4.6.2–4.8.1 and .NET 6+.
- **.NET Framework 4.8** [V] https://learn.microsoft.com/en-us/dotnet/framework/install/versions-and-dependencies
  - Preinstalled on Windows 10 1903+ and every Windows 11.
  - 4.8.1 is preinstalled on Windows 11 22H2+.
  - **No end-user runtime install is needed** on any supported Windows.

---

## 10. Recommended README install section

**Structure**
1. **Before you start** (3 bullets):
   - Windows Excel only. Not Mac, not Excel for the web.
   - Find out whether your Excel is 32-bit or 64-bit.
   - Close Excel.
2. **Option A — Download and add to Excel** (manual, no scripts):
   1. Download `Modelwright-<ver>.zip` from Releases.
   2. **Unblock the zip:** right-click › Properties › tick Unblock › OK.
   3. Extract All.
   4. Copy `Modelwright64.xll` (or `Modelwright32.xll`) to `%APPDATA%\Microsoft\AddIns`.
      - Paste `%APPDATA%\Microsoft\AddIns` into the File Explorer address bar.
      - Optional: users can instead Browse to the extracted file.
   5. In Excel, add it through the Add-ins dialog:
      1. **File › Options › Add-ins**.
      2. At the bottom, **Manage: Excel Add-ins › Go…**.
      3. **Browse…**, select the .xll, then **OK**.
      4. Make sure **Modelwright** is ticked, then **OK**.
   6. Check that the **Modelwright** tab appears.
3. **Option B — One-click script**:
   - Steps 1–3 as in Option A, then right-click `install.ps1` › **Run with PowerShell** (Windows 11: Show more options).
   - Or paste the one-line `powershell -ExecutionPolicy Bypass -File …` command.
   - Uninstall: `uninstall.ps1`.
4. **Option C — Ask IT**, as a collapsible section for IT staff:
   - **Verify the binaries:** SHA-256, attestation and signature.
   - **Or build from source:** the prerequisites in §6.1, then one `dotnet build` line.
   - **Deploy:**
     - Copy the .xll to a local folder (Program Files is fine for per-machine use).
     - Set the user's `OPEN` value with a GPP registry item, Intune, or Active Setup.
     - Optionally add the publisher certificate to Trusted Publishers, and an AppLocker or WDAC publisher rule.
     - Point IT to the MOTW, `requireaddinsig` and AppLocker facts (§1.5, §2).
5. **Troubleshooting** (symptom → fix):
   - "Leave this add-in disabled" → unblock.
   - "file format and extension … don't match" → wrong bitness.
   - Shortcuts do nothing → Macabacus (§8).
   - Tab missing → Disabled Items.
   - "Sorry, we couldn't find …xll" → stale `OPEN` value: re-run the installer or untick it in Add-ins.
6. **Using alongside Macabacus** (§8 text).
7. **Uninstall:** script, or untick in Add-ins and delete the file.

**Screenshots to capture** (Windows 11 plus M365 Current Channel; redo the Excel ones on LTSC 2024 if they look different)
1. GitHub Releases page with the **Assets** list expanded and the zip highlighted.
2. Edge's download bubble, plus the "isn't commonly downloaded" › **Keep** › **Show more** › **Keep anyway** sequence, if it appears.
3. File › **Properties** › General tab of the zip, showing the **Unblock** checkbox (the "This file came from another computer…" security line).
4. **Extract All** dialog.
5. File Explorer with `%APPDATA%\Microsoft\AddIns` typed in the address bar.
6. Excel › **File › Account › About Excel** with "64-bit" visible.
7. **File › Options › Add-ins** page with **Manage: Excel Add-ins** and **Go…** highlighted.
8. The **Add-ins** dialog with **Browse…** highlighted, and again with **Modelwright** ticked.
9. The ribbon with the **Modelwright** tab.
10. The **"Microsoft Excel Security Notice"** MOTW block dialog. Quote its text verbatim in the README.
11. The bitness-mismatch error dialog.
12. The signed-but-untrusted-publisher notice. Once signing exists, capture it with `requireaddinsig` turned on through local policy on a test VM.
13. Explorer right-click on `install.ps1` showing **Run with PowerShell**, and the script's success output.
14. **File › Options › Add-ins › Manage: COM Add-ins** dialog with **Macabacus COM Add-in** visible.
15. **Manage: Disabled Items** dialog.
16. *(If built)* the MSI's SmartScreen prompt, unsigned and signed.

---

## Recommendations

### R1. Release assets (per tag `vX.Y.Z`)
| Asset | Purpose | v0.1.0 |
|---|---|---|
| `Modelwright-X.Y.Z.zip` | **Primary download.** Contains `Modelwright64.xll`, `Modelwright32.xll`, `install.ps1`, `uninstall.ps1`, `INSTALL.txt` (the 6-step manual path in plain text), `LICENSE`, `THIRD_PARTY_NOTICES.md` and `SHA256SUMS.txt`. | Yes |
| `Modelwright64.xll`, `Modelwright32.xll` | Bare files for IT and power users | Yes |
| `SHA256SUMS.txt` | Checksums of every asset (GitHub also shows per-asset digests) | Yes |
| Build-provenance attestation | Produced by `actions/attest-build-provenance`; verified with `gh attestation verify` | Yes (free on public repos) |
| `Modelwright-X.Y.Z-peruser.msi` | Optional per-user MSI. Picks bitness, registers `OPEN`, closes Excel. **Signed, or not shipped**: an unsigned MSI adds SmartScreen and AppLocker friction for no gain over the script. | Later (after D-sign and the WiX licensing decision) |
| Source | The tag and commit SHA. GitHub's automatic source archives aren't attestable, so point IT to the tag. | Yes (automatic) |

Also: turn on **immutable releases**; record **VirusTotal** links in the release notes; keep the existing license gate.

### R2. Install-path decision table
| User's situation | Path | Why |
|---|---|---|
| Personal or small-firm PC, can run scripts | **B: script** | Picks bitness, unblocks, needs Excel closed, registers. Fewest mistakes. |
| Scripts blocked (GPO Restricted, `.ps1` blocked by EDR), but Excel's Add-ins dialog is allowed | **A: manual** | No PowerShell needed; Unblock plus the Add-ins dialog are the whole job. |
| AppLocker script rules only (`.ps1` runs in CLM) | **B** (if the script is CLM-safe) or **A** | The script runs constrained; `.cmd` wrappers are blocked. |
| `requireaddinsig` enforced (M365 security baseline) | **C: IT**, signed build required | Unsigned can never load; signed needs a publisher trust (re-trusted per release with Artifact Signing). |
| AppLocker DLL rules or WDAC enforced | **C: IT** | Needs a path or publisher rule, or a per-machine copy in Program Files with `OPEN` set by GPP or Active Setup. |
| IT forbids internet binaries entirely | **C: build from source** | Pinned SDK, locked restore, `dotnet build`; compare with the attested release. |
| Managed fleet, admin-run software deployment | Signed **MSI** (when it exists) or GPP | Standard tooling; per-user MSI, or a per-machine file plus per-user `OPEN`. |
| Mac or Excel for the web | Not supported | Office.js colors-only edition on the v2 roadmap. |

### R3. Build changes for Phase 5 (code work, not done here)
1. Set `ExcelDnaCreate32BitAddIn=true` and name the outputs `Modelwright32/64.xll`. Test once on 32-bit Excel, reviewing pointer-size P/Invoke code first.
2. Write `install.ps1`/`uninstall.ps1` per §1.4 (CLM-safe, PS 5.1). Add a CI step that runs them under ConstrainedLanguage against a scratch `HKCU` key.
3. Make builds reproducible: commit-date `BuildDate`, lock files, and a check on whether packing is deterministic.
4. Add `release.yml` per §6.3, with signing steps gated on secrets.
5. Add a session-only "Shortcuts on/off" toggle for Macabacus users (§8).

### R4. Decisions for the owner
- **D-sign: signing route** (§3.2–3.4): unsigned → SignPath / Artifact Signing / OV. Key facts:
  - Artifact Signing's 3-day certificates break Trusted-Publisher continuity across releases [I].
  - SignPath shows "SignPath Foundation" as the publisher.
  - The org 3-year rule is [U].
- **D-msi: whether to ship an MSI, and with which tool.** WiX (OSMF fee if Pegasus counts as "revenue-generating" ≥ $10k), Inno (license requested at > $5k revenue), or no MSI for v0.1.0.
- **D-scope:** state "Windows Excel 2021/2024/M365, 32- and 64-bit; Arm untested; Mac and web not supported" on the README front page.

### R5. Not verified; check before relying on it
1. The exact body text of the MOTW "Microsoft Excel Security Notice", and whether a security notice appears for an unblocked, unsigned .xll opened by double-click.
2. Whether a Trusted-Publisher-signed .xll *with* MOTW loads (Microsoft says it does for .xll and doesn't for .xlam).
3. That Explorer's Extract All copies MOTW onto extracted files (behavior of the Windows version in use).
4. Whether AppLocker DLL rules apply to `.xll`.
5. That the "Run with PowerShell" verb sets Process Bypass, on Windows 11.
6. That `New-Item -Force` on an existing registry key wipes its values.
7. Startup order between `OPEN` add-ins and COM add-ins, and Ctrl+Z hook interplay with Macabacus.
8. That a standard user can untick an HKLM-registered Macabacus COM add-in.
9. Whether the Artifact Signing org 3-year rule is still enforced in 2026.
10. FactSet's install model.
11. Excel-DNA behavior on Arm64EC.
12. Whether Excel-DNA packing is deterministic.
