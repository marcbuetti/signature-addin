<#
  Registers (or removes) the add-in manifest for desktop Outlook sideloading via the
  HKCU WEF\Developer key. Classic Outlook and recent new-Outlook builds read this for
  developer sideloading. After adding, restart Outlook.

  Usage:
    powershell -ExecutionPolicy Bypass -File .\sideload.ps1 -Action add
    powershell -ExecutionPolicy Bypass -File .\sideload.ps1 -Action remove
#>
param(
  [ValidateSet("add", "remove")]
  [string]$Action = "add",
  [string]$Manifest = "$PSScriptRoot\manifest.xml"
)

$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$manifestPath = (Resolve-Path $Manifest).Path

if (-not (Test-Path $key)) {
  New-Item -Path $key -Force | Out-Null
}

if ($Action -eq "add") {
  Set-ItemProperty -Path $key -Name $manifestPath -Value $manifestPath
  Write-Output "Sideloaded: $manifestPath"
  Write-Output "Restart Outlook, then start a new message to trigger the signature."
}
else {
  if (Get-ItemProperty -Path $key -Name $manifestPath -ErrorAction SilentlyContinue) {
    Remove-ItemProperty -Path $key -Name $manifestPath
    Write-Output "Removed sideload: $manifestPath"
  }
  else {
    Write-Output "Nothing to remove for: $manifestPath"
  }
}
