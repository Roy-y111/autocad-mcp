$ErrorActionPreference = 'Stop'

function Assert-AutoCADGuiClosed {
    $gui = Get-Process -Name acad -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($gui) {
        throw 'AutoCAD GUI is running. Close AutoCAD before this workflow so the Core Console cannot share or overwrite an active user profile.'
    }
}

function Get-RegistryValueSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $key = Get-Item -LiteralPath $Path
    $exists = $key.GetValueNames() -contains $Name
    [pscustomobject]@{
        Path = $Path
        Name = $Name
        Exists = $exists
        Kind = if ($exists) { [string]$key.GetValueKind($Name) } else { $null }
        Value = if ($exists) { $key.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) } else { $null }
    }
}

function Get-AutoCADProfileSettingSnapshots {
    $root = 'HKCU:\Software\Autodesk\AutoCAD'
    if (-not (Test-Path -LiteralPath $root)) { return @() }

    $snapshots = @()
    foreach ($release in Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue) {
        foreach ($product in Get-ChildItem -LiteralPath $release.PSPath -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like 'ACAD-*' }) {
            $fixedPath = Join-Path $product.PSPath 'FixedProfile\General Configuration'
            $fileDialog = Get-RegistryValueSnapshot -Path $fixedPath -Name 'FileDialog'
            if ($null -ne $fileDialog) { $snapshots += $fileDialog }

            $profilesPath = Join-Path $product.PSPath 'Profiles'
            if (Test-Path -LiteralPath $profilesPath) {
                $profilesKey = Get-Item -LiteralPath $profilesPath
                $profileName = [string]$profilesKey.GetValue('')
                if ($profileName) {
                    $variablesPath = Join-Path $profilesPath (Join-Path $profileName 'Variables')
                    $secureLoad = Get-RegistryValueSnapshot -Path $variablesPath -Name 'SECURELOAD'
                    if ($null -ne $secureLoad) { $snapshots += $secureLoad }
                }
            }
        }
    }
    return @($snapshots)
}

function Restore-RegistryValueSnapshot {
    param([Parameter(Mandatory = $true)]$Snapshot)

    if (-not (Test-Path -LiteralPath $Snapshot.Path)) {
        throw "AutoCAD registry path disappeared while restoring settings: $($Snapshot.Path)"
    }
    $subKey = $Snapshot.Path -replace '^Microsoft\.PowerShell\.Core\\Registry::HKEY_CURRENT_USER\\', ''
    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($subKey, $true)
    if ($null -eq $key) { throw "Could not open AutoCAD registry path for restoration: $($Snapshot.Path)" }
    try {
        if ($Snapshot.Exists) {
            $kind = [System.Enum]::Parse([Microsoft.Win32.RegistryValueKind], $Snapshot.Kind)
            $key.SetValue($Snapshot.Name, $Snapshot.Value, $kind)
        }
        else {
            $key.DeleteValue($Snapshot.Name, $false)
        }
    }
    finally {
        $key.Dispose()
    }
}

function Test-RegistryValueMatchesSnapshot {
    param([Parameter(Mandatory = $true)]$Snapshot)

    $current = Get-RegistryValueSnapshot -Path $Snapshot.Path -Name $Snapshot.Name
    if ($null -eq $current) { return $false }
    if ($current.Exists -ne $Snapshot.Exists) { return $false }
    if (-not $Snapshot.Exists) { return $true }
    return (($current.Kind -eq $Snapshot.Kind) -and ([string]$current.Value -eq [string]$Snapshot.Value))
}

function Invoke-AutoCADCoreGuarded {
    param(
        [Parameter(Mandatory = $true)][string]$CoreConsole,
        [Parameter(Mandatory = $true)][string]$Drawing,
        [Parameter(Mandatory = $true)][string]$Script,
        [Parameter(Mandatory = $true)][string]$Log
    )

    Assert-AutoCADGuiClosed
    $snapshots = @(Get-AutoCADProfileSettingSnapshots)
    if ($snapshots.Count -eq 0) {
        throw 'Could not snapshot the AutoCAD user-profile settings. Core Console was not started.'
    }

    $exitCode = $null
    $processError = $null
    try {
        & $CoreConsole /i $Drawing /s $Script /l en-US 2>&1 | Out-File -LiteralPath $Log -Encoding utf8
        $exitCode = $LASTEXITCODE
    }
    catch {
        $processError = $_
    }
    finally {
        $restoreErrors = @()
        foreach ($snapshot in $snapshots) {
            try { Restore-RegistryValueSnapshot -Snapshot $snapshot }
            catch { $restoreErrors += $_.Exception.Message }
        }
        foreach ($snapshot in $snapshots) {
            if (-not (Test-RegistryValueMatchesSnapshot -Snapshot $snapshot)) {
                $restoreErrors += "Verification failed for $($snapshot.Path)::$($snapshot.Name)"
            }
        }
        if ($restoreErrors.Count -gt 0) {
            throw "AutoCAD user-profile restoration failed: $($restoreErrors -join '; ')"
        }
    }

    if ($processError) { throw $processError }
    if ($exitCode -ne 0) { throw "AutoCAD Core Console failed with exit code $exitCode. See $Log" }
}
