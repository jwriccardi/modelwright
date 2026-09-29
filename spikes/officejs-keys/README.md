# Spike K1 - Office.js keyboard-shortcut key strings

## Purpose

`docs/research/05-keys-and-undo.md` ruled Office.js out of contention for exact Macabacus
keyboard shortcuts because the extended-manifest schema restricts every key string to
`^[A-Za-z0-9-_+]+$` (letters, digits, hyphen, underscore, plus). That rules out literal
punctuation like `Ctrl+[`, `Ctrl+'`, `Ctrl+;`, `Ctrl+,`, `Ctrl+.`. But Microsoft's own sample
uses named keys like `"Up"` / `"Down"`, which fit that regex. Nobody had found evidence that a
named key for punctuation (e.g. `BracketLeft`, `Semicolon`, `Oem1`) actually works.

This spike empirically tests, in real desktop Excel (M365 x64, Current Channel):

1. Do **named** keys for punctuation (`BracketLeft`, `Oem1`, `OemComma`, ...) actually bind and
   fire, or does Office silently ignore/reject them?
2. Do **literal** punctuation characters (`Ctrl+[`, `Ctrl+'`, ...) work anyway, despite failing
   the schema regex - i.e. is the regex enforced at all?
3. When a shortcuts file has one bad/unrecognized entry, does Office reject the **whole file**
   (breaking even the valid entries in it), or does it drop only the bad entry?

The answer determines whether Office.js should stay ruled out (per ADR-0002, Windows desktop via
Excel-DNA is primary) or come back into contention.

## What's here

Three minimal, no-build-step Office Add-ins (Excel TaskPaneApp, add-in-only XML manifest,
`VersionOverrides` V1_0, `SharedRuntime` 1.1, `<Runtime lifetime="long">`, one ribbon button, and
an `ExtendedOverrides` keyboard-shortcuts JSON file), all served from a single static server on
`https://localhost:3100`.

```
spikes/officejs-keys/
  manifests/
    manifest-a.xml   Add-in A "EMT K1 Named Keys"   (Id 91c5a88b-1521-47fb-ac5a-8c381f0eb854)
    manifest-b.xml   Add-in B "EMT K1 Literal Keys" (Id 4d1cd616-e36b-42f3-ad1f-8467d78bb100)
    manifest-c.xml   Add-in C "EMT K1 Single"       (Id 181291ff-36c5-47e3-8980-2c68d40d950a)
  web/                          <- served as document root on https://localhost:3100
    taskpane-a.html / -b.html / -c.html   task pane pages (also the shared-runtime pages)
    taskpane-common.js                    shared logging/diagnostics/probe logic
    taskpane-style.css
    shortcuts-a.json / -b.json / -c.json  ExtendedOverrides keyboard-shortcut definitions
    assets/icon-16.png, icon-32.png, icon-80.png
  server.js       zero-dependency Node https static server (port 3100)
  setup.ps1       installs dev cert + registers the 3 manifests for sideload (owner-run)
  serve.ps1       starts server.js (owner-run, keep the terminal open)
  cleanup.ps1     removes the sideload registry entries (owner-run)
  README.md       this file
```

### Add-in A - EMT K1 Named Keys (`shortcuts-a.json`)

Baseline controls (expected to work; if these fail, something is wrong with the harness, not
with named punctuation keys):
- `Control1`: `Ctrl+Alt+Shift+K`
- `Control2`: `Ctrl+Alt+Up` (Microsoft's documented named-arrow-key sample style)

Candidates (17 named-key spellings for the 5 target chords `Ctrl+Shift+[`, `Ctrl+[`, `Ctrl+'`,
`Ctrl+;`, `Ctrl+,`, `Ctrl+.`, plus a couple of extra plausible spellings and a `Ctrl+F2` sanity
check):
`Ctrl+Shift+BracketLeft`, `Ctrl+BracketLeft`, `Ctrl+BracketRight`, `Ctrl+Quote`,
`Ctrl+Semicolon`, `Ctrl+Comma`, `Ctrl+Period`, `Ctrl+OemComma`, `Ctrl+OemPeriod`, `Ctrl+Oem1`,
`Ctrl+Oem4`, `Ctrl+Oem7`, `Ctrl+Shift+OpenBracket`, `Ctrl+F2`, `Ctrl+OemSemicolon`,
`Ctrl+OemQuotes`, `Ctrl+OemOpenBrackets`.

(`OemComma`/`OemPeriod`/`Oem1`/`Oem4`/`Oem7` and the `OemSemicolon`/`OemQuotes`/
`OemOpenBrackets` variants are .NET `System.Windows.Forms.Keys` enum spellings; `BracketLeft`/
`Quote`/`Semicolon`/`Comma`/`Period` are DOM `KeyboardEvent.code` spellings.)

### Add-in B - EMT K1 Literal Keys (`shortcuts-b.json`)

Baseline control: `ControlJ`: `Ctrl+Alt+Shift+J`.

Literal candidates: `Ctrl+Shift+[`, `Ctrl+[`, `Ctrl+'`, `Ctrl+;`, `Ctrl+,`, `Ctrl+.`.

**These deliberately fail the published extended-manifest schema regex.** That's the point:
Add-in B exists to see what Office actually does when the file contains schema-invalid entries
(reject file / drop entries / accept anyway), as opposed to A's control, which only uses entries
that pass the regex.

### Add-in C - EMT K1 Single (`shortcuts-c.json`)

Baseline control: `ControlL`: `Ctrl+Alt+Shift+L`, plus **exactly one** named-key candidate:
`Ctrl+Shift+BracketLeft`.

If Add-in A's larger file (18 candidates, some possibly unrecognized) ends up with *no* working
shortcuts at all, Add-in C isolates whether that's whole-file rejection (C would also fail) or a
problem specific to one of A's other entries (C would still work).

### What each task pane does

Each task pane (the shared-runtime page itself) shows, on load:
- add-in name
- `Office.context.diagnostics` (host, platform, version)
- whether `SharedRuntime` 1.1 and `KeyboardShortcuts` 1.1 are reported supported
  (`Office.context.requirements.isSetSupported`)
- if `KeyboardShortcuts` 1.1 is supported: the result of `Office.actions.getShortcuts()` - **this
  is the key evidence of what Office actually parsed** out of the shortcuts JSON file, as opposed
  to what we wrote in it
- the result of `Office.actions.areShortcutsInUse(...)` for that add-in's candidate key strings
- the result of trying `Office.actions.replaceShortcuts(...)` for a couple of candidates, in two
  different call shapes (the exact accepted shape is unverified/preview behavior, so both an
  array-of-`{key,action}` shape and an object-literal `{actionId: key}` shape are tried and
  logged independently)
- a table of every action id + key string registered from that add-in's shortcuts file
- any errors encountered, shown on screen (not just console)

Every registered action handler (`Office.actions.associate`) appends
`<timestamp> fired: <actionId> (<key>)` to a visible, newest-first log list in the pane, and also
`console.log`s it.

No build step, no npm dependencies for the add-in code itself - plain HTML/JS/CSS/JSON, loading
`office.js` from `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`.

## Owner steps

1. **Read this file first.** Excel should be closed for step 2 if possible (registry changes are
   picked up on next Excel launch); it does not need to be closed for step 3.
2. Run `.\setup.ps1` in a PowerShell terminal from this folder.
   - It runs `npx --yes office-addin-dev-certs install`, which will show a **Windows
     certificate-trust prompt** - accept it (this is the standard, well-known Office-Addin dev
     localhost cert, scoped to `localhost`).
   - It then writes 3 string values under
     `HKCU:\Software\Microsoft\Office\16.0\WEF\Developer` (value name = each add-in's manifest
     GUID, value data = full path to that manifest's `.xml`). This is Microsoft's documented
     manual sideload mechanism for Windows desktop Office (see comments in `setup.ps1` for
     citations) - no Excel launch, no other registry writes.
3. Run `.\serve.ps1` in its **own terminal window** and leave it running. It starts
   `server.js` (Node's built-in `https` module, no dependencies) on `https://localhost:3100`,
   using the cert installed in step 2.
4. (Re)start Excel, open (or create) any workbook.
5. **Home tab > Add-ins > My Add-ins > Developer Add-ins** (or the "..." / Developer tab,
   depending on your Excel build) - you should see all three: *EMT K1 Named Keys*,
   *EMT K1 Literal Keys*, *EMT K1 Single*. Open each one's task pane (ribbon button, or directly
   from that list).
6. For each pane: read the on-load diagnostics block (especially `getShortcuts()` output -
   that's the ground truth of what Office parsed), then click into the Excel grid (a cell must be
   selected, not the task pane) and press each candidate key combination one at a time.
7. **What to record**, per add-in, per key:
   - Did the pane's log show a `fired:` line? (shortcut works)
   - Did Excel show a **conflict/override dialog** (e.g. "this shortcut is already assigned")
     instead of firing, or alongside firing?
   - Did nothing happen (key passed through to Excel's own behavior, or silently swallowed)?
   - Did the pane show any on-screen errors, especially around `getShortcuts()`,
     `areShortcutsInUse()`, or `replaceShortcuts()`?
   - For Add-in B specifically: did `getShortcuts()` list the literal-key entries at all, list
     only the `ControlJ` baseline, or fail/error entirely?
   - For Add-in C: does the lone `Ctrl+Shift+BracketLeft` candidate fire, independent of what
     Add-in A does with its larger file?
8. When finished, run `.\cleanup.ps1` (optionally `.\cleanup.ps1 -UninstallCert` to also remove
   the dev cert) and restart Excel again to fully drop the three add-ins.

## Validation already performed (by the assistant, not the owner)

- All 3 `shortcuts-*.json` files parse as valid JSON (`node -e "JSON.parse(...)"`).
- All 3 `manifest-*.xml` files are well-formed XML (PowerShell `[xml]` cast succeeded for each).
- `server.js` and `web/taskpane-common.js` pass `node --check` (syntax only; not executed).
- Icon PNGs (`icon-16/32/80.png`) are valid PNGs (`file` reports correct dimensions/color type
  for each).
- Ran `npx --yes office-addin-manifest validate <manifest>` against all three manifests. It
  completed in well under 60s for each (no hang), but **reported the same generic errors for all
  three** ("Manifest product ID Not Valid... must be a plain GUID", "Package Type Not
  Identified", "Wrong Package"), even though each `<Id>` is a plain GUID and each manifest is a
  normal TaskPaneApp. This tool runs a remote AppSource **acceptance-test service** that expects
  a submission-style package and appears to need to actually fetch the manifest's referenced
  URLs (icons, task pane, shortcuts JSON) - which weren't reachable since the local server
  wasn't running during this check and the tool doesn't appear intended for bare sideload-only
  manifests. **Treat this result as inconclusive**, not as evidence of a manifest defect; the
  `[xml]` well-formedness check plus manual comparison against the OfficeApp/VersionOverrides
  schema (namespaces, `VersionOverridesV1_0`, `ExtendedOverrides` shape) is the more meaningful
  local check here. If the owner wants a truly authoritative validation, run it again with
  `serve.ps1` already running so the URLs resolve.
- Did **not** launch Excel, install certificates, write the registry, or start a long-running
  server - all of that is left to the owner via the three `.ps1` scripts, per the spike's hard
  safety rules.

## Uncertain / left for the owner to observe

- Whether any named punctuation key (`BracketLeft`, `Oem1`, etc.) actually fires in real Excel -
  this is the whole point of the spike and cannot be determined without a running Excel session.
- Whether literal punctuation keys in Add-in B cause whole-file rejection, partial rejection, or
  work despite the schema.
- The exact accepted call shape for `Office.actions.replaceShortcuts` - both plausible shapes are
  tried and logged, but neither is confirmed against real behavior yet.
- Whether `Office.actions.getShortcuts()` / `areShortcutsInUse()` are even available in this
  Excel build (`KeyboardShortcuts` 1.1 requirement-set support is itself unverified) - the pane
  reports this and skips those probes gracefully if unsupported.
- Whether Excel shows a distinguishable conflict dialog for `Ctrl+Shift+[`-style combos that
  collide with something else in the same session (e.g. Windows OS-level chord capture) - only
  observable interactively.

## Cleanup

Run `.\cleanup.ps1` when done testing. It removes the 3 registry values under
`HKCU:\Software\Microsoft\Office\16.0\WEF\Developer`; pass `-UninstallCert` to also remove the
office-addin-dev-certs localhost certificate. Stop `serve.ps1` with Ctrl+C in its terminal. No
files under this spike folder need to be deleted for cleanup (they're just static files/scripts).
