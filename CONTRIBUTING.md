# Contributing

Thanks for helping. This is a Windows desktop Excel add-in written in C# with [Excel-DNA](https://excel-dna.net/). See [docs/PLAN.md](docs/PLAN.md) for the design.

## Build and test

You need the .NET SDK 10 (see `global.json`) and the .NET 8 runtime (for the `net8.0` test run). Visual Studio is not required; the .NET Framework 4.8 reference assemblies come from NuGet.

```powershell
dotnet restore
dotnet build Modelwright.sln -c Release
dotnet test Modelwright.sln -c Release --no-build
powershell -NoProfile -ExecutionPolicy Bypass -File build/check-licenses.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/install/install-scripts.Tests.ps1
```

- Warnings are errors, so the build must report 0 warnings.
- The tests run twice: on `net8.0`, and on `net48`, which is the runtime inside Excel.
- `tests/install/install-scripts.Tests.ps1` runs `install/install.ps1` and `install/uninstall.ps1` in Windows PowerShell 5.1 under ConstrainedLanguage (as AppLocker would), against a scratch `HKCU:\Software\Modelwright-InstallTest-*` key and a scratch folder in `%TEMP%`. It never touches Excel's own registry keys and works with Excel open. The install scripts may use only built-in cmdlets: no .NET method calls (including `$PSCmdlet.ShouldProcess`), `Add-Type` or COM.
- `build/check-licenses.ps1` fails if a shipped project (`src/`) uses a NuGet package that is not in `build/allowed-packages.json`. If you add a dependency, review its license and add it there in the same pull request. LGPL, GPL, AGPL, SSPL and non-commercial licenses are not accepted for shipped code.
- Any new DLL that ships with the add-in must also be added to `ExcelAddInInclude` in `src/Modelwright.AddIn/Modelwright.AddIn.csproj`. Otherwise it is not packed into the `.xll`, and Excel fails at load with a `FileNotFoundException`. Add its license notice to `THIRD_PARTY_NOTICES.md` too (for MPL-2.0 code such as XLParser, with a link to its source).

## Load the development add-in in Excel

The build produces a packed add-in for each Excel bitness (use the one that matches Excel, usually 64-bit):

```
src\Modelwright.AddIn\bin\Release\net48\publish\Modelwright64.xll
src\Modelwright.AddIn\bin\Release\net48\publish\Modelwright32.xll
```

1. **Disable Macabacus first, or change its keys.** Whichever add-in registers a shortcut last owns it.
2. In Excel: **File > Options > Add-ins > Manage: Excel Add-ins > Go... > Browse...**, select the `.xll`, and click **OK**.3. You should see a **Modelwright** ribbon tab and the status-bar message "Modelwright *version* loaded".
4. Press **Ctrl+Alt+Shift+F12**, or click **Modelwright > About**, to show the version, commit, build date and add-in path.

### Rebuilding while Excel is open

Excel keeps a loaded `.xll` locked for as long as it runs, and keys its add-in list by **file name**: a `Modelwright64.xll` at any other path counts as the same add-in, so a new build is never loaded until Excel restarts. Use `tests/excel-smoke/restart-excel.ps1` (Windows PowerShell 5.1; it refuses if any workbook has unsaved changes, and never saves):

```powershell
powershell -ExecutionPolicy Bypass -File tests/excel-smoke/restart-excel.ps1 -Stop    # close Excel
dotnet build Modelwright.sln -c Release                                              # now the .xll can be replaced
powershell -ExecutionPolicy Bypass -File tests/excel-smoke/restart-excel.ps1 -Start   # start Excel; it loads the new build
```

Without a switch it closes and restarts Excel in one go. To build without closing Excel, build into another folder (`dotnet build Modelwright.sln -c Release -p:OutputPath=bin\scratch\`) to check that it compiles; Excel still runs the old build.

### Excel smoke tests

With Excel open, `tests/excel-smoke/undo-smoke.ps1` and `tests/excel-smoke/trace-smoke.ps1` drive Excel with real keystrokes in scratch workbooks (`powershell -ExecutionPolicy Bypass -File tests/excel-smoke/undo-smoke.ps1`). Don't touch the keyboard while they run. Each script's header says what it covers.

A downloaded `.xll` is blocked by Mark-of-the-Web. If Excel refuses to load one you did not build yourself, right-click the file > **Properties** > tick **Unblock**.

## Releases

Pushing a tag `vX.Y.Z` runs `.github/workflows/release.yml`. The tag must match the version in `Directory.Build.props`. The workflow builds and tests, then creates a **draft** GitHub Release with `Modelwright-X.Y.Z.zip` (both add-ins, `install/*`, `LICENSE`, `THIRD_PARTY_NOTICES.md` and a `SHA256SUMS.txt`), the bare `.xll` files, `SHA256SUMS.txt` and build-provenance attestations. The add-ins are code-signed only if the `SIGNING_*` secrets are set (see the comment at the top of the workflow); otherwise they ship unsigned. Before publishing the draft, scan the assets on VirusTotal and add the links to the notes.

## Coding standards

- **`Modelwright.Core` stays Excel-free.** It targets `netstandard2.0` and must not reference Excel-DNA, COM, Office interop or Windows-only APIs. Put logic there (cycles, colors, keys, undo model, formula parsing) and cover it with unit tests.
- **No COM in Core.** Excel access (C API, COM, ribbon, windows) lives only in `Modelwright.AddIn`, as thin adapters over Core.
- **Every Core change comes with tests** in `tests/Modelwright.Core.Tests`.
- Commands must never throw into Excel. Report failures in the status bar.
- **`settings.json` has a versioned schema** (`schemaVersion`). The loader rejects unknown properties, so any change to the file's shape (a new, renamed or removed property, or a new cycle `kind`) must bump `ToolkitSettings.CurrentSchemaVersion` and still read the previous version, with tests. Otherwise users' existing files fall back to the defaults.
- Follow `.editorconfig`: 4-space indents, file-scoped namespaces, nullable reference types enabled.

## Sign your commits (DCO)

Every commit must carry a `Signed-off-by` line. It certifies the [Developer Certificate of Origin](https://developercertificate.org/) (DCO): you wrote the change, or otherwise have the right to submit it under this project's license (PolyForm Shield 1.0.0, see LICENSE). Add it with `-s`:

```powershell
git commit -s -m "Describe the change"
```

This appends `Signed-off-by: Your Name <you@example.com>`, using your `git config user.name` and `user.email`. To sign off commits you already made on a branch, run `git rebase --signoff main` and force-push the branch.

Pull requests with unsigned commits cannot be merged.

## License

By contributing, you agree that your contributions are licensed under the [PolyForm Shield License 1.0.0](LICENSE), with Pegasus Technology Group LLC as the licensor.
