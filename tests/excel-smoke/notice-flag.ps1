# Dot-sourced by the Excel smoke tests (trace-smoke.ps1, undo-smoke.ps1): keeps the add-in's one-time Macabacus notice
# away during a run. The notice is a modal dialog that would take the keyboard focus from the keys the test sends.
# The add-in does not show it once ui-state.json says it was shown ("macabacusNoticeShown": true), so a run sets that
# flag before it loads the add-in and puts the previous value back in `finally`. (The add-in also skips it with
# MODELWRIGHT_NO_NOTICES=1, but a script cannot set the environment of the Excel it drives.)
# Only that one property is changed; the rest of the file (the Trace In window's place) is left as it is.

$noticeFlagPattern = '"macabacusNoticeShown"\s*:\s*(true|false)'
$noticeFlagUtf8 = New-Object Text.UTF8Encoding $false

# Sets "macabacusNoticeShown": true in $Path. Returns what to pass to Restore-NoticeFlag: whether the file existed, and
# the flag's previous value ('true', 'false', or $null when it was not there).
function Set-NoticeFlag([string]$Path) {
  $before = New-Object PSObject -Property @{ Path = $Path; Existed = (Test-Path -LiteralPath $Path); Value = $null; Written = $null }
  if ($before.Existed) {
    $text = [IO.File]::ReadAllText($Path)
    if ($text -match $noticeFlagPattern) {
      $before.Value = $Matches[1]
      $text = $text -replace $noticeFlagPattern, '"macabacusNoticeShown": true'
    }
    elseif ($text -match '^\s*\{\s*\}\s*$') {
      $text = "{`r`n  `"macabacusNoticeShown`": true`r`n}`r`n"
    }
    elseif ($text -match '\}\s*$') {
      $text = $text -replace '\s*\}\s*$', ",`r`n  `"macabacusNoticeShown`": true`r`n}`r`n"
    }
    else {
      throw "ABORT: $Path is not a JSON object; can't keep the Macabacus notice away. Nothing was changed."
    }
  }
  else {
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Path) | Out-Null
    $text = "{`r`n  `"schemaVersion`": 1,`r`n  `"macabacusNoticeShown`": true`r`n}`r`n"
  }
  [IO.File]::WriteAllText($Path, $text, $noticeFlagUtf8)
  $before.Written = $text
  return $before
}

# Puts the flag back as Set-NoticeFlag found it: its old value, or no flag; a file the run created and nothing else
# changed since is deleted. Returns a line for the output.
function Restore-NoticeFlag($Before) {
  $path = $Before.Path
  if (-not (Test-Path -LiteralPath $path)) { return "ui-state.json is gone; nothing to put back" }
  $text = [IO.File]::ReadAllText($path)
  if (-not $Before.Existed -and $text -eq $Before.Written) {
    Remove-Item -LiteralPath $path
    return "Removed $path (the run created it)"
  }
  if ($null -ne $Before.Value) {
    $text = $text -replace $noticeFlagPattern, ('"macabacusNoticeShown": ' + $Before.Value)
  }
  else {
    $text = $text -replace (',\s*' + $noticeFlagPattern), ''
    $text = $text -replace ($noticeFlagPattern + '\s*,?'), ''
  }
  [IO.File]::WriteAllText($path, $text, $noticeFlagUtf8)
  return "Macabacus notice flag in $path put back to $(if ($null -ne $Before.Value) { $Before.Value } else { 'not set' })"
}
