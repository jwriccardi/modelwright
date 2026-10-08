# Contributing

Thanks for helping. This is a Windows desktop Excel add-in written in C# with [Excel-DNA](https://excel-dna.net/). See [docs/PLAN.md](docs/PLAN.md) for the design.

## Build and test

You need the .NET SDK 10 (see `global.json`) and the .NET 8 runtime (for the `net8.0` test run). Visual Studio is not required; the .NET Framework 4.8 reference assemblies come from NuGet.

```powershell
dotnet restore
dotnet build ExcelModelingToolkit.sln -c Release
dotnet test ExcelModelingToolkit.sln -c Release --no-build
powershell -NoProfile -ExecutionPolicy Bypass -File build/check-licenses.ps1
```

- Warnings are errors, so the build must report 0 warnings.
- The tests run twice: on `net8.0`, and on `net48`, which is the runtime inside Excel.
- `build/check-licenses.ps1` fails if a shipped project (`src/`) uses a NuGet package that is not in `build/allowed-packages.json`. If you add a dependency, review its license and add it there in the same pull request. LGPL, GPL, AGPL, SSPL and non-commercial licenses are not accepted for shipped code.
- Any new DLL that ships with the add-in must also be added to `ExcelAddInInclude` in `src/ExcelModelingToolkit.AddIn/ExcelModelingToolkit.AddIn.csproj`. Otherwise it is not packed into the `.xll`, and Excel fails at load with a `FileNotFoundException`. Add its license notice to `THIRD_PARTY_NOTICES.md` too (for MPL-2.0 code such as XLParser, with a link to its source).

## Load the development add-in in Excel

The build produces a single packed 64-bit add-in:

```
src\ExcelModelingToolkit.AddIn\bin\Release\net48\publish\ModelingToolkit64-packed.xll
```

1. **Disable Macabacus first, or change its keys.** Whichever add-in registers a shortcut last owns it.
2. In Excel: **File > Options > Add-ins > Manage: Excel Add-ins > Go... > Browse...**, select the `.xll`, and click **OK**.
3. You should see a **Modeling Toolkit** ribbon tab and the status-bar message "Modeling Toolkit *version* loaded".
4. Press **Ctrl+Alt+Shift+F12**, or click **Modeling Toolkit > About**, to show the version, commit, build date and add-in path.
5. To unload, untick the add-in in the same dialog. Excel keeps the `.xll` file open while the add-in is loaded, so unload it (or close Excel) before rebuilding.

A downloaded `.xll` is blocked by Mark-of-the-Web. If Excel refuses to load one you did not build yourself, right-click the file > **Properties** > tick **Unblock**.

## Coding standards

- **`ExcelModelingToolkit.Core` stays Excel-free.** It targets `netstandard2.0` and must not reference Excel-DNA, COM, Office interop or Windows-only APIs. Put logic there (cycles, colors, keys, undo model, formula parsing) and cover it with unit tests.
- **No COM in Core.** Excel access (C API, COM, ribbon, windows) lives only in `ExcelModelingToolkit.AddIn`, as thin adapters over Core.
- **Every Core change comes with tests** in `tests/ExcelModelingToolkit.Core.Tests`.
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
