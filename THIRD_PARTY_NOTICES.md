# Third-party notices

The Modeling Toolkit add-in (`ModelingToolkit64-packed.xll`) includes the third-party software listed below. Its own code is licensed under the [MIT License](LICENSE).

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
