# Third-party notices

The Modelwright add-in (`Modelwright64.xll`) includes the third-party software listed below. Its own code is licensed under the [PolyForm Shield License 1.0.0](LICENSE). The optional Windows installer (`Modelwright-<version>-x64.msi`) contains the same add-in plus the WiX component listed at the end.

## Excel-DNA

- **Packages:** `ExcelDna.AddIn` 1.9.0 and `ExcelDna.Integration` 1.9.0. The Excel-DNA native loader and the `ExcelDna.Integration` runtime are embedded in the packed `.xll`.
- **Project:** https://excel-dna.net/ · https://github.com/Excel-DNA/ExcelDna
- **License:** zlib

```
zLib license

Copyright (C) 2005-2021 Govert van Drimmelen

This software is provided 'as-is', without any express or implied
warranty.  In no event will the authors be held liable for any damages
arising from the use of this software.

Permission is granted to anyone to use this software for any purpose,
including commercial applications, and to alter it and redistribute it
freely, subject to the following restrictions:

1. The origin of this software must not be misrepresented; you must not
   claim that you wrote the original software. If you use this software
   in a product, an acknowledgment in the product documentation would be
   appreciated but is not required.
2. Altered source versions must be plainly marked as such, and must not be
   misrepresented as being the original software.
3. This notice may not be removed or altered from any source distribution.
```

The license text above is copied from `LICENSE.txt` at the Excel-DNA `v1.9.0` tag. The 1.9.0 binaries themselves carry the copyright notice "Copyright (C) 2005-2024 Govert van Drimmelen".

## XLParser

- **Package:** `XLParser` 1.7.5, by the TU Delft Spreadsheet Lab and Infotron. Its `XLParser.dll` is embedded, unmodified, in the packed `.xll`, where Trace In uses it to parse formulas.
- **Project:** https://github.com/spreadsheetlab/XLParser
- **License:** Mozilla Public License, version 2.0 (MPL-2.0). Full text: https://mozilla.org/MPL/2.0/, also in the project's [`LICENSE.md` at tag v1.7.5](https://github.com/spreadsheetlab/XLParser/blob/v1.7.5/LICENSE.md).
- **Source code:** XLParser is distributed here in executable form under MPL-2.0. Its source code is available at https://github.com/spreadsheetlab/XLParser/tree/v1.7.5 (commit `fb9af413dd561d7dda6e03710ca9d3d9f2d03c02`, the commit the NuGet package was built from). We have not modified it. MPL-2.0 applies only to XLParser's own files; it does not extend to the rest of this add-in.

## Irony

- **Package:** `Irony` 1.5.3 (a dependency of XLParser). Its `Irony.dll` is embedded in the packed `.xll`.
- **Project:** https://github.com/IronyProject/Irony
- **License:** MIT

```
MIT License

Copyright (c) 2019 Irony Project (https://github.com/IronyProject)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software
and associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or
substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE
AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
```

The license text above is copied from `LICENSE` at the Irony commit the 1.5.3 package was built from (`1098ffb279a3bbb2c617f5cf5685be6afe836a7f`).

## WiX Toolset DTF (in the MSI only)

- **Package:** `WixToolset.Dtf.WindowsInstaller` 5.0.2 (via `WixToolset.Dtf.CustomAction` 5.0.2). Its `WixToolset.Dtf.WindowsInstaller.dll` is embedded, unmodified, in the MSI's custom-action DLL (`installer/msi/CustomActions`), which uses it to talk to Windows Installer during install and uninstall. It is not part of the `.xll`.
- **Project:** https://wixtoolset.org/ · https://github.com/wixtoolset/wix (tag `v5.0.2`)
- **License:** Microsoft Reciprocal License (MS-RL). Full text: https://licenses.nuget.org/MS-RL (the license expression in the package's .nuspec). Source code: https://github.com/wixtoolset/wix/tree/v5.0.2. We have not modified it. MS-RL's reciprocal terms apply only to files of the licensed software; they do not extend to Modelwright's own code.
- The WiX command-line tool (`wix` 5.0.2, MS-RL) is used to build the MSI and is not distributed.

## Build-time and test-only dependencies (not distributed)

These packages are used only to build or test the add-in. Nothing from them is included in the `.xll`.

| Package | License | Use |
|---|---|---|
| Microsoft.NETFramework.ReferenceAssemblies (+ `.net48`) | MIT | .NET Framework 4.8 reference assemblies for building without Visual Studio |
| NETStandard.Library, Microsoft.NETCore.Platforms | MIT; MS .NET Library license | Implicit build references of the `netstandard2.0` Core library |
| xunit, xunit.runner.visualstudio | Apache-2.0 | Unit tests |
| Microsoft.NET.Test.Sdk | MIT | Unit tests |
| coverlet.collector | MIT | Code coverage in tests |

`build/allowed-packages.json` lists every package allowed in the shipped projects, and `build/check-licenses.ps1` enforces it in CI.
