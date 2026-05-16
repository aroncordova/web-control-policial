$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot
$Python = "C:\Users\acordova\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$Generator = Join-Path $PSScriptRoot "generate_control_data.py"
$Log = Join-Path $Root "watch-excel.log"
$SharePointExcel = "C:\Users\acordova\OneDrive - EMOVA MOVILIDAD S.A\COMPARTIDO CMI - COMPARTIDO CMI\VARIAS\02 - PLANILLA DE CONTROL DE PERSONAL POLICIAL\2026\05 - CONTROL POLICIAL - MAYO.xlsx"

function Get-ExcelState {
  $files = @(Get-ChildItem -Path $Root -File |
    Where-Object {
      ($_.Extension -in ".xlsx", ".xls") -and
      (-not $_.Name.StartsWith("~$"))
    } |
    Sort-Object Name)

  if (Test-Path -LiteralPath $SharePointExcel) {
    $files += Get-Item -LiteralPath $SharePointExcel
  }

  if (-not $files) {
    return ""
  }

  return ($files | ForEach-Object { "$($_.Name)|$($_.LastWriteTimeUtc.Ticks)|$($_.Length)" }) -join "`n"
}

function Invoke-Generate {
  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Add-Content -Path $Log -Value "[$timestamp] Regenerando datos desde Excel..."
  Add-Content -Path $Log -Value "[$timestamp] Fuente vigilada: $SharePointExcel"
  & $Python $Generator 2>&1 | Add-Content -Path $Log
}

Invoke-Generate
$lastState = Get-ExcelState

while ($true) {
  Start-Sleep -Seconds 5
  try {
    $currentState = Get-ExcelState
    if ($currentState -ne $lastState) {
      Start-Sleep -Seconds 2
      Invoke-Generate
      $lastState = Get-ExcelState
    }
  } catch {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $Log -Value "[$timestamp] Error: $($_.Exception.Message)"
  }
}
