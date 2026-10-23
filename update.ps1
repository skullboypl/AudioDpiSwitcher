param(
    [Parameter(Mandatory = $true)][int]$ProcessId,
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Destination
)

$ErrorActionPreference = "Stop"

try {
    Wait-Process -Id $ProcessId -ErrorAction SilentlyContinue
} catch {
}

Start-Sleep -Milliseconds 800
Copy-Item -LiteralPath $Source -Destination $Destination -Force
Start-Process -FilePath $Destination
Remove-Item -LiteralPath $Source -Force -ErrorAction SilentlyContinue
