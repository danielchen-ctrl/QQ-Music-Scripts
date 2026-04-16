param(
    [string]$AdbExe = "adb",
    [string]$DeviceId,
    [int]$TargetLeBi = 330,
    [string]$TargetTimeText = "12:00:00",
    [bool]$FastMode = $true,
    [bool]$EnableRehearsal = $true,
    [bool]$ManualStartMode = $true,
    [switch]$SkipTimeGate,
    [switch]$BurstOnly,
    [switch]$DryRun,
    [switch]$PreflightOnly,
    [switch]$RehearsalOnly,
    [switch]$ListDevicesJson,
    [string]$ArtifactRoot,
    [int]$RechargeToServiceTapDelayMs = 0,
    [int]$ServiceToPopupTapDelayMs = 0,
    [int]$PreBurstSettleMs = 0,
    [int]$RehearsalSafetyMarginMs = 80
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$QqMusicPackage = "com.tencent.qqmusic"
$BaseResolution = @{
    Width  = 1080
    Height = 2400
}

$BaseCoords = @{
    RechargeCard    = @{ X = 540; Y = 506 }
    ServiceExchange = @{ X = 980; Y = 724 }
    PopupMinus      = @{ X = 244; Y = 1198 }
    PopupPlus       = @{ X = 836; Y = 1198 }
    PopupCancel     = @{ X = 326; Y = 1538 }
    PopupExchange   = @{ X = 754; Y = 1538 }
}

$RelativeFallbacks = @{
    PopupPlus = @{
        Anchor = "PopupExchange"
        DeltaX = 82
        DeltaY = -340
    }
}

function New-TextFromCodePoints {
    param([int[]]$CodePoints)

    return (-join ($CodePoints | ForEach-Object { [char]$_ }))
}

$TextLeBiRecharge = New-TextFromCodePoints -CodePoints @(0x4E50, 0x5E01, 0x5145, 0x503C)
$TextQuickRechargeEntry = New-TextFromCodePoints -CodePoints @(0x4E50, 0x5E01, 0x5FEB, 0x6377, 0x5145, 0x503C, 0x5165, 0x53E3)
$TextLeBiService = New-TextFromCodePoints -CodePoints @(0x4E50, 0x5E01, 0x670D, 0x52A1)
$TextExchange = New-TextFromCodePoints -CodePoints @(0x5151, 0x6362)
$TextServiceExchange = New-TextFromCodePoints -CodePoints @(0x670D, 0x52A1, 0x5151, 0x6362)
$TextExchangeLeBi = New-TextFromCodePoints -CodePoints @(0x5151, 0x6362, 0x4E50, 0x5E01)
$TextIncrease = New-TextFromCodePoints -CodePoints @(0x589E, 0x52A0)
$TextDecrease = New-TextFromCodePoints -CodePoints @(0x51CF, 0x5C11)
$TextImmediateExchange = New-TextFromCodePoints -CodePoints @(0x7ACB, 0x5373, 0x5151, 0x6362)
$TextCoinWelfare = New-TextFromCodePoints -CodePoints @(0x91D1, 0x5E01, 0x798F, 0x5229)
$TextPer = New-TextFromCodePoints -CodePoints @(0x6BCF)
$TextCan = New-TextFromCodePoints -CodePoints @(0x53EF)
$TextCancel = New-TextFromCodePoints -CodePoints @(0x53D6, 0x6D88)
$TextExchangeRateHint = $TextPer + "10000" + $TextCan + $TextExchange + "10"

$LocatorProfiles = @{
    RechargeCard = @(
        @{ Label = "text equals lebi recharge"; TextEquals = $TextLeBiRecharge; PreferClickable = $true; Sort = "TopMost" }
        @{ Label = "text contains lebi recharge"; TextContains = $TextLeBiRecharge; PreferClickable = $true; Sort = "TopMost" }
        @{ Label = "desc contains lebi recharge"; DescContains = $TextLeBiRecharge; PreferClickable = $true; Sort = "TopMost" }
        @{ Label = "text contains quick recharge entry"; TextContains = $TextQuickRechargeEntry; PreferClickable = $true; Sort = "TopMost" }
        @{ Label = "desc contains quick recharge entry"; DescContains = $TextQuickRechargeEntry; PreferClickable = $true; Sort = "TopMost" }
    )
    ServicePageMarker = @(
        @{ Label = "text equals lebi service"; TextEquals = $TextLeBiService; PreferClickable = $false; Sort = "TopMost" }
        @{ Label = "desc equals lebi service"; DescEquals = $TextLeBiService; PreferClickable = $false; Sort = "TopMost" }
    )
    ServiceExchange = @(
        @{ Label = "text equals exchange"; TextEquals = $TextExchange; PreferClickable = $true; Sort = "RightMost" }
        @{ Label = "text contains service exchange"; TextContains = $TextServiceExchange; PreferClickable = $true; Sort = "RightMost" }
        @{ Label = "desc contains service exchange"; DescContains = $TextServiceExchange; PreferClickable = $true; Sort = "RightMost" }
        @{ Label = "text contains exchange"; TextContains = $TextExchange; PreferClickable = $true; Sort = "RightMost" }
    )
    PopupTitle = @(
        @{ Label = "text equals exchange lebi"; TextEquals = $TextExchangeLeBi; PreferClickable = $false; Sort = "TopMost" }
        @{ Label = "desc equals exchange lebi"; DescEquals = $TextExchangeLeBi; PreferClickable = $false; Sort = "TopMost" }
    )
    PopupPlus = @(
        @{ Label = "text equals plus"; TextEquals = "+"; PreferClickable = $true; Sort = "RightMost" }
        @{ Label = "desc equals plus"; DescEquals = "+"; PreferClickable = $true; Sort = "RightMost" }
        @{ Label = "resource-id contains increase"; ResourceIdContains = "increase"; PreferClickable = $true; Sort = "RightMost" }
        @{ Label = "text contains increase"; TextContains = $TextIncrease; PreferClickable = $true; Sort = "RightMost" }
        @{ Label = "desc contains increase"; DescContains = $TextIncrease; PreferClickable = $true; Sort = "RightMost" }
        @{ Label = "resource-id contains plus"; ResourceIdContains = "plus"; PreferClickable = $true; Sort = "RightMost" }
        @{ Label = "resource-id contains add"; ResourceIdContains = "add"; PreferClickable = $true; Sort = "RightMost" }
    )
    PopupMinus = @(
        @{ Label = "text equals minus"; TextEquals = "-"; PreferClickable = $true; Sort = "TopMost" }
        @{ Label = "desc equals minus"; DescEquals = "-"; PreferClickable = $true; Sort = "TopMost" }
        @{ Label = "resource-id contains decrease"; ResourceIdContains = "decrease"; PreferClickable = $true; Sort = "TopMost" }
        @{ Label = "resource-id contains minus"; ResourceIdContains = "minus"; PreferClickable = $true; Sort = "TopMost" }
        @{ Label = "resource-id contains subtract"; ResourceIdContains = "subtract"; PreferClickable = $true; Sort = "TopMost" }
        @{ Label = "text contains decrease"; TextContains = $TextDecrease; PreferClickable = $true; Sort = "TopMost" }
        @{ Label = "desc contains decrease"; DescContains = $TextDecrease; PreferClickable = $true; Sort = "TopMost" }
    )
    PopupCancel = @(
        @{ Label = "text equals cancel"; TextEquals = $TextCancel; PreferClickable = $true; Sort = "BottomMost" }
        @{ Label = "desc equals cancel"; DescEquals = $TextCancel; PreferClickable = $true; Sort = "BottomMost" }
    )
    PopupExchange = @(
        @{ Label = "text equals exchange"; TextEquals = $TextExchange; PreferClickable = $true; Sort = "BottomMost" }
        @{ Label = "text contains immediate exchange"; TextContains = $TextImmediateExchange; PreferClickable = $true; Sort = "BottomMost" }
        @{ Label = "text contains exchange"; TextContains = $TextExchange; PreferClickable = $true; Sort = "BottomMost" }
        @{ Label = "desc contains exchange"; DescContains = $TextExchange; PreferClickable = $true; Sort = "BottomMost" }
    )
}

$DefaultLeBi = 10
$StepLeBi = 10
$TargetLeBi = [int]$TargetLeBi
if ($TargetLeBi -lt $DefaultLeBi) {
    throw "TargetLeBi must be at least $DefaultLeBi."
}
if ((($TargetLeBi - $DefaultLeBi) % $StepLeBi) -ne 0) {
    throw "TargetLeBi must increase in steps of $StepLeBi from $DefaultLeBi."
}
$PlusTapCount = [int](($TargetLeBi - $DefaultLeBi) / $StepLeBi)

$Timing = @{
    TimeGatePollMs             = 50
    UiPollMs                   = 90
    RechargeOpenTimeoutMs      = 1200
    RechargeToServiceTimeoutMs = 900
    ServicePopupTimeoutMs      = 700
    PopupTitleTimeoutMs        = 350
    PostTapSettleMs            = 100
    RechargeOpenSettleMs       = 450
    PopupOpenSettleMs          = 260
    ServiceRetryGapMs          = 80
}

$FastTiming = @{
    InitialSettleMs          = 100
    RechargeToServiceMs      = 260
    ServiceToPopupMs         = 220
    PreBurstSettleMs         = 40
    QuickCheckBudgetMs       = 420
    QuickPollMs              = 60
}

$Rehearsal = @{
    LeadTimeMs         = 90000
    CacheTtlMinutes    = 30
    ReturnSettleMs     = 180
    PopupRollbackGapMs = 80
}

$KnownFastDeviceProfiles = @(
    [pscustomobject]@{
        Manufacturer = "Xiaomi"
        Model        = "24129PN74C"
        Width        = 1200
        Height       = 2670
        Density      = "520"
        Coords       = @{
            RechargeCard    = @{ X = 424; Y = 592 }
            ServiceExchange = @{ X = 1089; Y = 805 }
            PopupMinus      = @{ X = 322; Y = 1350 }
            PopupPlus       = @{ X = 936; Y = 1350 }
            PopupCancel     = @{ X = 382; Y = 1711 }
            PopupExchange   = @{ X = 838; Y = 1711 }
        }
        Delays       = @{
            RechargeToServiceTapDelayMs = 260
            ServiceToPopupTapDelayMs    = 220
            PreBurstSettleMs            = 40
        }
    }
)

$script:DeviceProfile = $null
$script:CalibrationProfile = $null
$script:FastDeviceProfile = $null
$script:FastModeActive = $false
$script:ExecutionMode = "AdaptiveMode"
$script:LastUiSnapshot = $null
$script:RehearsalCache = $null
$script:ResolvedTapPoints = @{}
$script:StateRoot = Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "state"
$script:ArtifactRoot = if ([string]::IsNullOrWhiteSpace($ArtifactRoot)) {
    Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "artifacts"
}
else {
    $ArtifactRoot
}
$script:RunStamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
$script:RunArtifactDir = $null

try {
    $TargetTime = [TimeSpan]::ParseExact(
        $TargetTimeText,
        "hh\:mm\:ss",
        [System.Globalization.CultureInfo]::InvariantCulture
    )
}
catch {
    throw "TargetTimeText must use HH:mm:ss, for example 12:00:00."
}

$TargetTimeCompact = "{0:D2}{1:D2}{2:D2}" -f $TargetTime.Hours, $TargetTime.Minutes, $TargetTime.Seconds
$TargetTimeText = "{0:D2}:{1:D2}:{2:D2}" -f $TargetTime.Hours, $TargetTime.Minutes, $TargetTime.Seconds

function Write-Log {
    param([string]$Message)

    $stamp = Get-Date -Format "HH:mm:ss"
    Write-Host "[$stamp] $Message"
}

function Ensure-ArtifactDirectory {
    if (-not $script:RunArtifactDir) {
        $script:RunArtifactDir = Join-Path -Path $script:ArtifactRoot -ChildPath ("qq_music_exchange_{0}" -f $script:RunStamp)
        [void](New-Item -ItemType Directory -Path $script:RunArtifactDir -Force)
    }

    return $script:RunArtifactDir
}

function Invoke-Adb {
    param(
        [string[]]$Arguments,
        [switch]$AllowFailure
    )

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()

    try {
        $process = Start-Process `
            -FilePath $script:AdbExe `
            -ArgumentList $Arguments `
            -NoNewWindow `
            -Wait `
            -PassThru `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath

        $stdout = if (Test-Path $stdoutPath) { Get-Content -LiteralPath $stdoutPath -Raw -Encoding UTF8 -ErrorAction SilentlyContinue } else { "" }
        $stderr = if (Test-Path $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw -Encoding UTF8 -ErrorAction SilentlyContinue } else { "" }
        $text = (($stdout, $stderr) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine
        $exitCode = $process.ExitCode
    }
    finally {
        Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    }

    if ($exitCode -ne 0 -and -not $AllowFailure) {
        if ([string]::IsNullOrWhiteSpace($text)) {
            throw "adb failed with exit code $exitCode"
        }

        throw "adb failed: $text"
    }

    return $text
}

function Invoke-AdbToFile {
    param(
        [string[]]$Arguments,
        [string]$OutputPath,
        [switch]$AllowFailure
    )

    $stderrPath = [System.IO.Path]::GetTempFileName()

    try {
        $process = Start-Process `
            -FilePath $script:AdbExe `
            -ArgumentList $Arguments `
            -NoNewWindow `
            -Wait `
            -PassThru `
            -RedirectStandardOutput $OutputPath `
            -RedirectStandardError $stderrPath

        $stderr = if (Test-Path $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue } else { "" }
        $exitCode = $process.ExitCode
    }
    finally {
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
    }

    if ($exitCode -ne 0 -and -not $AllowFailure) {
        if ([string]::IsNullOrWhiteSpace($stderr)) {
            throw "adb failed with exit code $exitCode"
        }

        throw "adb failed: $stderr"
    }
}

function Get-ConnectedDeviceInfos {
    $raw = Invoke-Adb -Arguments @("devices", "-l")
    $lines = $raw -split "`r?`n"
    $devices = New-Object System.Collections.Generic.List[object]

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line) -or $line -match "^List of devices attached") {
            continue
        }

        if ($line -match "^(?<id>\S+)\s+(?<state>\S+)(?<rest>.*)$") {
            $rest = $matches.rest
            $devices.Add([pscustomobject]@{
                Id          = $matches.id
                State       = $matches.state
                Product     = if ($rest -match "product:(\S+)") { $matches[1] } else { "" }
                Model       = if ($rest -match "model:(\S+)") { $matches[1] } else { "" }
                Device      = if ($rest -match "device:(\S+)") { $matches[1] } else { "" }
                TransportId = if ($rest -match "transport_id:(\d+)") { $matches[1] } else { "" }
            })
        }
    }

    return $devices.ToArray()
}

function Get-ConnectedDevices {
    return @((Get-ConnectedDeviceInfos | Where-Object { $_.State -eq "device" }).Id)
}

function Get-PreferredDeviceLabel {
    param([pscustomobject]$DeviceInfo)

    if (-not $DeviceInfo) {
        return ""
    }

    if (-not [string]::IsNullOrWhiteSpace($DeviceInfo.Model)) {
        return $DeviceInfo.Model
    }

    return $DeviceInfo.Device
}

function Resolve-Device {
    $deviceInfos = @(Get-ConnectedDeviceInfos)

    if ($script:DeviceId) {
        $selected = $deviceInfos | Where-Object { $_.Id -eq $script:DeviceId } | Select-Object -First 1
        if (-not $selected) {
            throw "Requested device '$script:DeviceId' is not connected."
        }
        if ($selected.State -ne "device") {
            throw "Requested device '$script:DeviceId' is in state '$($selected.State)'."
        }

        Write-Log ("Using requested device: {0} ({1})" -f $script:DeviceId, (Get-PreferredDeviceLabel -DeviceInfo $selected))
        return
    }

    $devices = @($deviceInfos | Where-Object { $_.State -eq "device" })
    if ($devices.Count -eq 0) {
        throw "No Android device is connected."
    }

    if ($devices.Count -gt 1) {
        throw "More than one device is connected. Re-run with -DeviceId."
    }

    $script:DeviceId = $devices[0].Id
    Write-Log ("Using detected device: {0} ({1})" -f $script:DeviceId, (Get-PreferredDeviceLabel -DeviceInfo $devices[0]))
}

function Get-AdbDeviceArgs {
    if ([string]::IsNullOrWhiteSpace($script:DeviceId)) {
        return @()
    }

    return @("-s", $script:DeviceId)
}

function Invoke-DeviceShell {
    param(
        [string]$Command,
        [switch]$AllowFailure
    )

    $args = (Get-AdbDeviceArgs) + @("shell", $Command)
    return Invoke-Adb -Arguments $args -AllowFailure:$AllowFailure
}

function Get-DeviceTimeText {
    $args = (Get-AdbDeviceArgs) + @("shell", "date", "+%H:%M:%S")
    return (Invoke-Adb -Arguments $args).Trim()
}

function Get-DeviceResolution {
    $args = (Get-AdbDeviceArgs) + @("shell", "wm", "size")
    $raw = Invoke-Adb -Arguments $args

    foreach ($line in ($raw -split "`r?`n")) {
        if ($line -match "(\d+)x(\d+)") {
            return [pscustomobject]@{
                Width  = [int]$matches[1]
                Height = [int]$matches[2]
                Text   = "$($matches[1])x$($matches[2])"
            }
        }
    }

    throw "Could not parse device resolution from: $raw"
}

function Get-DeviceDensityText {
    $args = (Get-AdbDeviceArgs) + @("shell", "wm", "density")
    $raw = Invoke-Adb -Arguments $args -AllowFailure
    $match = [regex]::Match($raw, "(\d+)")
    if ($match.Success) {
        return $match.Value
    }

    return ""
}

function Get-DeviceProp {
    param([string]$Name)

    $args = (Get-AdbDeviceArgs) + @("shell", "getprop", $Name)
    return (Invoke-Adb -Arguments $args -AllowFailure).Trim()
}

function Get-DeviceProfile {
    $profileText = Invoke-DeviceShell -Command (
        "wm size; " +
        "wm density; " +
        "echo __MANUFACTURER__; getprop ro.product.manufacturer; " +
        "echo __MODEL__; getprop ro.product.model; " +
        "echo __ANDROID__; getprop ro.build.version.release; " +
        "echo __SDK__; getprop ro.build.version.sdk"
    )

    $resolution = $null
    $density = ""
    $manufacturer = ""
    $model = ""
    $android = ""
    $sdk = ""

    $lines = @($profileText -split "`r?`n")
    for ($index = 0; $index -lt $lines.Count; $index++) {
        $line = $lines[$index].Trim()
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        if (-not $resolution -and $line -match "(\d+)x(\d+)") {
            $resolution = [pscustomobject]@{
                Width  = [int]$matches[1]
                Height = [int]$matches[2]
                Text   = "$($matches[1])x$($matches[2])"
            }
            continue
        }

        if ([string]::IsNullOrWhiteSpace($density) -and $line -match "(\d+)") {
            $density = $matches[1]
            continue
        }

        switch ($line) {
            "__MANUFACTURER__" {
                if (($index + 1) -lt $lines.Count) { $manufacturer = $lines[$index + 1].Trim() }
            }
            "__MODEL__" {
                if (($index + 1) -lt $lines.Count) { $model = $lines[$index + 1].Trim() }
            }
            "__ANDROID__" {
                if (($index + 1) -lt $lines.Count) { $android = $lines[$index + 1].Trim() }
            }
            "__SDK__" {
                if (($index + 1) -lt $lines.Count) { $sdk = $lines[$index + 1].Trim() }
            }
        }
    }

    if (-not $resolution) {
        throw "Could not parse device profile from adb output: $profileText"
    }

    $profile = [pscustomobject]@{
        Resolution   = $resolution.Text
        Width        = $resolution.Width
        Height       = $resolution.Height
        ScaleX       = [double]$resolution.Width / $BaseResolution.Width
        ScaleY       = [double]$resolution.Height / $BaseResolution.Height
        Density      = $density
        Manufacturer = $manufacturer
        Model        = $model
        Android      = $android
        Sdk          = $sdk
    }

    return $profile
}

function Test-DeviceProfileMatches {
    param(
        [pscustomobject]$Actual,
        [pscustomobject]$Expected
    )

    if (-not $Actual -or -not $Expected) {
        return $false
    }

    return $Actual.Model -eq $Expected.Model -and `
        $Actual.Width -eq $Expected.Width -and `
        $Actual.Height -eq $Expected.Height -and `
        ([string]::IsNullOrWhiteSpace($Expected.Density) -or $Actual.Density -eq $Expected.Density)
}

function Initialize-ExecutionModes {
    $script:FastModeActive = $false
    $script:FastDeviceProfile = $null
    $script:ExecutionMode = "AdaptiveMode"

    if (-not $FastMode) {
        return
    }

    if (-not $ManualStartMode) {
        return
    }

    foreach ($candidate in $KnownFastDeviceProfiles) {
        if (Test-DeviceProfileMatches -Actual $script:DeviceProfile -Expected $candidate) {
            $script:FastModeActive = $true
            $script:FastDeviceProfile = $candidate
            $script:ExecutionMode = "TurboMode"
            return
        }
    }
}

function Select-ExecutionMode {
    if (Test-RehearsalCacheValid -Cache $script:RehearsalCache) {
        $script:ExecutionMode = "Rehearsed TurboMode"
        return
    }

    if ($script:FastModeActive) {
        $script:ExecutionMode = "TurboMode"
        return
    }

    $script:ExecutionMode = "AdaptiveMode"
}

function Write-ExecutionModeLog {
    switch ($script:ExecutionMode) {
        "Rehearsed TurboMode" {
            Write-Log "Rehearsed TurboMode active."
            Write-Log "Using rehearsal-cached points."
        }
        "TurboMode" {
            Write-Log "TurboMode active."
            if ($script:CalibrationProfile) {
                Write-Log "Using calibration points."
            }
            elseif ($script:FastDeviceProfile) {
                Write-Log "Using profile points."
            }
        }
        default {
            Write-Log "AdaptiveMode active."
            Write-Log "Falling back to adaptive path."
        }
    }
}

function Get-ShellSecondsText {
    param([int]$Milliseconds)

    $seconds = [double]$Milliseconds / 1000
    return $seconds.ToString("0.###", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Ensure-StateDirectory {
    if (-not (Test-Path -LiteralPath $script:StateRoot)) {
        [void](New-Item -ItemType Directory -Path $script:StateRoot -Force)
    }

    return $script:StateRoot
}

function Get-SafeFileComponent {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return "unknown"
    }

    return ([regex]::Replace($Value, '[^A-Za-z0-9._-]+', '_')).Trim('_')
}

function Get-NamedValue {
    param(
        $Container,
        [string]$Name
    )

    if ($null -eq $Container) {
        return $null
    }

    if ($Container -is [System.Collections.IDictionary]) {
        if ($Container.Contains($Name)) {
            return $Container[$Name]
        }
        return $null
    }

    $property = $Container.PSObject.Properties[$Name]
    if ($property) {
        return $property.Value
    }

    return $null
}

function Get-DeviceSignature {
    if (-not $script:DeviceProfile) {
        return ""
    }

    return "{0}_{1}_{2}_{3}x{4}_{5}" -f `
        (Get-SafeFileComponent -Value $script:DeviceId), `
        (Get-SafeFileComponent -Value $script:DeviceProfile.Manufacturer), `
        (Get-SafeFileComponent -Value $script:DeviceProfile.Model), `
        $script:DeviceProfile.Width, `
        $script:DeviceProfile.Height, `
        (Get-SafeFileComponent -Value $script:DeviceProfile.Density)
}

function Get-RehearsalCachePath {
    $signature = Get-DeviceSignature
    if ([string]::IsNullOrWhiteSpace($signature)) {
        return ""
    }

    return Join-Path -Path (Ensure-StateDirectory) -ChildPath ("rehearsal_{0}.json" -f $signature)
}

function Get-CalibrationPath {
    $signature = Get-DeviceSignature
    if ([string]::IsNullOrWhiteSpace($signature)) {
        return ""
    }

    return Join-Path -Path (Ensure-StateDirectory) -ChildPath ("calibration_{0}.json" -f $signature)
}

function Convert-TapPointForStorage {
    param([pscustomobject]$Point)

    if (-not $Point) {
        return $null
    }

    return [ordered]@{
        X      = [int]$Point.X
        Y      = [int]$Point.Y
        Source = [string]$Point.Source
    }
}

function Convert-StoredPointToTapPoint {
    param(
        [string]$ActionName,
        $Point
    )

    if (-not $Point) {
        return $null
    }

    $x = Get-NamedValue -Container $Point -Name "X"
    $y = Get-NamedValue -Container $Point -Name "Y"
    if ($null -eq $x -or $null -eq $y) {
        return $null
    }

    return [pscustomobject]@{
        ActionName = $ActionName
        X          = [int]$x
        Y          = [int]$y
        Source     = [string](Get-NamedValue -Container $Point -Name "Source")
        MatchCount = 1
    }
}

function Load-JsonFile {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
    catch {
        Write-Log ("Ignoring invalid JSON state file {0}: {1}" -f $Path, $_.Exception.Message)
        return $null
    }
}

function Save-JsonFile {
    param(
        [string]$Path,
        $Value
    )

    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent)) {
        [void](New-Item -ItemType Directory -Path $parent -Force)
    }

    $Value | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Load-CalibrationProfile {
    $path = Get-CalibrationPath
    $profile = Load-JsonFile -Path $path
    if ($profile) {
        Write-Log ("Loaded calibration profile from {0}" -f $path)
    }

    return $profile
}

function Test-RehearsalCacheValid {
    param($Cache)

    if (-not $Cache) {
        return $false
    }

    $deviceSignature = [string](Get-NamedValue -Container $Cache -Name "DeviceSignature")
    if ([string]::IsNullOrWhiteSpace($deviceSignature) -or $deviceSignature -ne (Get-DeviceSignature)) {
        return $false
    }

    $points = Get-NamedValue -Container $Cache -Name "Points"
    foreach ($requiredPoint in @("RechargeCard", "ServiceExchange", "PopupPlus", "PopupExchange")) {
        if (-not (Get-NamedValue -Container $points -Name $requiredPoint)) {
            return $false
        }
    }

    $savedAtRaw = [string](Get-NamedValue -Container $Cache -Name "SavedAt")
    if ([string]::IsNullOrWhiteSpace($savedAtRaw)) {
        return $false
    }

    try {
        $savedAt = [datetimeoffset]::Parse($savedAtRaw, [System.Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        return $false
    }

    return ($savedAt.UtcDateTime -ge (Get-Date).ToUniversalTime().AddMinutes(-1 * $Rehearsal.CacheTtlMinutes))
}

function Load-RehearsalCache {
    $path = Get-RehearsalCachePath
    $cache = Load-JsonFile -Path $path
    if (-not (Test-RehearsalCacheValid -Cache $cache)) {
        return $null
    }

    Write-Log "Rehearsal cache is valid."
    return $cache
}

function Save-RehearsalCache {
    param(
        [pscustomobject]$RechargePoint,
        [pscustomobject]$ServicePoint,
        [pscustomobject]$PopupPlusPoint,
        [pscustomobject]$PopupExchangePoint,
        [pscustomobject]$PopupMinusPoint,
        [pscustomobject]$PopupCancelPoint,
        [int]$MeasuredRechargeDelayMs,
        [int]$MeasuredPopupDelayMs
    )

    $path = Get-RehearsalCachePath
    $payload = [ordered]@{
        DeviceId        = $script:DeviceId
        DeviceSignature = Get-DeviceSignature
        SavedAt         = (Get-Date).ToString("o", [System.Globalization.CultureInfo]::InvariantCulture)
        Valid           = $true
        Points          = [ordered]@{
            RechargeCard    = Convert-TapPointForStorage -Point $RechargePoint
            ServiceExchange = Convert-TapPointForStorage -Point $ServicePoint
            PopupPlus       = Convert-TapPointForStorage -Point $PopupPlusPoint
            PopupExchange   = Convert-TapPointForStorage -Point $PopupExchangePoint
            PopupMinus      = Convert-TapPointForStorage -Point $PopupMinusPoint
            PopupCancel     = Convert-TapPointForStorage -Point $PopupCancelPoint
        }
        Delays          = [ordered]@{
            RechargeToServiceTapDelayMs = [Math]::Max(0, $MeasuredRechargeDelayMs)
            ServiceToPopupTapDelayMs    = [Math]::Max(0, $MeasuredPopupDelayMs)
            PreBurstSettleMs            = [Math]::Max(0, $(if ($PreBurstSettleMs -gt 0) { $PreBurstSettleMs } else { $FastTiming.PreBurstSettleMs }))
            RehearsalSafetyMarginMs     = [Math]::Max(0, $RehearsalSafetyMarginMs)
        }
    }

    Save-JsonFile -Path $path -Value $payload
    $script:RehearsalCache = Load-RehearsalCache
    Write-Log "Rehearsal capture completed."
}

function Get-CurrentDeviceTime {
    $currentText = Get-DeviceTimeText
    $timeMatch = [regex]::Match($currentText, '\b\d{2}:\d{2}:\d{2}\b')
    if (-not $timeMatch.Success) {
        throw "Could not parse device time from adb output: '$currentText'"
    }

    return [TimeSpan]::Parse($timeMatch.Value)
}

function Get-MillisecondsUntilTargetTime {
    $currentTime = Get-CurrentDeviceTime
    return [int][Math]::Round(($TargetTime - $currentTime).TotalMilliseconds)
}

function Get-ConfiguredDelayMs {
    param(
        [string]$DelayName,
        [int]$DefaultValue,
        [int]$OverrideValue = 0,
        [switch]$IncludeRehearsalSafetyMargin
    )

    $value = $DefaultValue
    if ($OverrideValue -gt 0) {
        $value = $OverrideValue
    }
    elseif ($script:RehearsalCache) {
        $delays = Get-NamedValue -Container $script:RehearsalCache -Name "Delays"
        $cached = Get-NamedValue -Container $delays -Name $DelayName
        if ($cached -gt 0) {
            $value = [int]$cached
        }
    }
    elseif ($script:CalibrationProfile) {
        $delays = Get-NamedValue -Container $script:CalibrationProfile -Name "Delays"
        $calibrated = Get-NamedValue -Container $delays -Name $DelayName
        if ($calibrated -gt 0) {
            $value = [int]$calibrated
        }
    }
    elseif ($script:FastDeviceProfile) {
        $profileDelays = Get-NamedValue -Container $script:FastDeviceProfile -Name "Delays"
        $profileValue = Get-NamedValue -Container $profileDelays -Name $DelayName
        if ($profileValue -gt 0) {
            $value = [int]$profileValue
        }
    }

    if ($IncludeRehearsalSafetyMargin) {
        $margin = $RehearsalSafetyMarginMs
        if ($script:RehearsalCache) {
            $delays = Get-NamedValue -Container $script:RehearsalCache -Name "Delays"
            $cachedMargin = Get-NamedValue -Container $delays -Name "RehearsalSafetyMarginMs"
            if ($cachedMargin -ge 0) {
                $margin = [int]$cachedMargin
            }
        }
        $value += $margin
    }

    return [Math]::Max(0, [int]$value)
}

function Get-CurrentFocusText {
    $commands = @(
        @("shell", "dumpsys", "window", "windows"),
        @("shell", "dumpsys", "activity", "activities")
    )

    foreach ($command in $commands) {
        $raw = Invoke-Adb -Arguments ((Get-AdbDeviceArgs) + $command) -AllowFailure
        foreach ($pattern in @(
            "mCurrentFocus.+?\s(?<focus>\S+/\S+)",
            "mFocusedApp.+?\s(?<focus>\S+/\S+)",
            "topResumedActivity.*?(?<focus>\S+/\S+)",
            "mResumedActivity:.*?\s(?<focus>\S+/\S+)",
            "ResumedActivity:.*?\s(?<focus>\S+/\S+)"
        )) {
            $match = [regex]::Match($raw, $pattern)
            if ($match.Success) {
                return $match.Groups["focus"].Value
            }
        }
    }

    return ""
}

function Assert-QqMusicForeground {
    $focus = Get-CurrentFocusText
    if ([string]::IsNullOrWhiteSpace($focus)) {
        Write-Log "Could not determine the current foreground window from dumpsys."
        return $false
    }

    if ($focus -notmatch [regex]::Escape($QqMusicPackage)) {
        throw "QQ Music is not in the foreground. Current focus: $focus"
    }

    Write-Log "Foreground window check passed: $focus"
    return $true
}

function Assert-QqMusicInstalled {
    $args = (Get-AdbDeviceArgs) + @("shell", "pm", "path", $QqMusicPackage)
    $raw = Invoke-Adb -Arguments $args -AllowFailure
    if ($raw -notmatch "^package:") {
        throw "QQ Music ($QqMusicPackage) is not installed on the selected device."
    }
}

function Get-UiBounds {
    param([string]$BoundsText)

    $match = [regex]::Match($BoundsText, "^\[(\d+),(\d+)\]\[(\d+),(\d+)\]$")
    if (-not $match.Success) {
        return $null
    }

    $left = [int]$match.Groups[1].Value
    $top = [int]$match.Groups[2].Value
    $right = [int]$match.Groups[3].Value
    $bottom = [int]$match.Groups[4].Value

    return [pscustomobject]@{
        Left   = $left
        Top    = $top
        Right  = $right
        Bottom = $bottom
        Width  = [Math]::Max(0, $right - $left)
        Height = [Math]::Max(0, $bottom - $top)
        Center = [pscustomobject]@{
            X = [int][Math]::Round(($left + $right) / 2.0)
            Y = [int][Math]::Round(($top + $bottom) / 2.0)
        }
    }
}

function Get-UiDumpXmlText {
    $remotePath = "/sdcard/qq_music_exchange_window_dump.xml"
    $dumpCommand = @"
uiautomator dump --compressed $remotePath >/dev/null 2>&1 || uiautomator dump $remotePath >/dev/null 2>&1
cat $remotePath
"@

    $raw = Invoke-DeviceShell -Command $dumpCommand -AllowFailure
    $xmlStart = $raw.IndexOf("<?xml")
    if ($xmlStart -lt 0) {
        throw "Could not read UI hierarchy from the device."
    }

    $xmlText = $raw.Substring($xmlStart).Trim()
    $xmlEndMarker = "</hierarchy>"
    $xmlEnd = $xmlText.LastIndexOf($xmlEndMarker, [System.StringComparison]::OrdinalIgnoreCase)
    if ($xmlEnd -ge 0) {
        $xmlText = $xmlText.Substring(0, $xmlEnd + $xmlEndMarker.Length)
    }

    return $xmlText
}

function Get-UiNodesFromXmlDocument {
    param([xml]$XmlDocument)

    $nodes = New-Object System.Collections.Generic.List[object]

    foreach ($node in $XmlDocument.SelectNodes("//node")) {
        $bounds = Get-UiBounds -BoundsText $node.GetAttribute("bounds")
        if (-not $bounds) {
            continue
        }

        $nodes.Add([pscustomobject]@{
            Text       = $node.GetAttribute("text")
            Desc       = $node.GetAttribute("content-desc")
            ResourceId = $node.GetAttribute("resource-id")
            ClassName  = $node.GetAttribute("class")
            Package    = $node.GetAttribute("package")
            Clickable  = $node.GetAttribute("clickable")
            Enabled    = $node.GetAttribute("enabled")
            Bounds     = $bounds
        })
    }

    return $nodes
}

function Get-UiautomatorAttributeValue {
    param(
        [string]$AttributeBlob,
        [string]$Name
    )

    $pattern = '(?s)\b' + [regex]::Escape($Name) + '="(?<value>.*?)(?=\s+[A-Za-z_][A-Za-z0-9_.:-]*=|\s*/?$|$)'
    $match = [regex]::Match($AttributeBlob, $pattern)
    if (-not $match.Success) {
        return ""
    }

    return $match.Groups["value"].Value.TrimEnd('"')
}

function Get-UiNodesFromRawDump {
    param([string]$XmlText)

    $nodes = New-Object System.Collections.Generic.List[object]
    $tagMatches = [regex]::Matches($XmlText, '<node\b(?<attrs>[^>]*)/?>', [System.Text.RegularExpressions.RegexOptions]::Singleline)

    foreach ($tagMatch in $tagMatches) {
        $attrs = $tagMatch.Groups["attrs"].Value
        $boundsText = Get-UiautomatorAttributeValue -AttributeBlob $attrs -Name "bounds"
        $bounds = Get-UiBounds -BoundsText $boundsText
        if (-not $bounds) {
            continue
        }

        $nodes.Add([pscustomobject]@{
            Text       = Get-UiautomatorAttributeValue -AttributeBlob $attrs -Name "text"
            Desc       = Get-UiautomatorAttributeValue -AttributeBlob $attrs -Name "content-desc"
            ResourceId = Get-UiautomatorAttributeValue -AttributeBlob $attrs -Name "resource-id"
            ClassName  = Get-UiautomatorAttributeValue -AttributeBlob $attrs -Name "class"
            Package    = Get-UiautomatorAttributeValue -AttributeBlob $attrs -Name "package"
            Clickable  = Get-UiautomatorAttributeValue -AttributeBlob $attrs -Name "clickable"
            Enabled    = Get-UiautomatorAttributeValue -AttributeBlob $attrs -Name "enabled"
            Bounds     = $bounds
        })
    }

    return $nodes
}

function Get-UiSnapshot {
    param([switch]$ForceRefresh)

    if (-not $ForceRefresh -and $script:LastUiSnapshot) {
        return $script:LastUiSnapshot
    }

    $xmlText = Get-UiDumpXmlText
    $nodes = $null
    $xmlParseMode = "xml"

    try {
        [xml]$xml = $xmlText
        $nodes = Get-UiNodesFromXmlDocument -XmlDocument $xml
    }
    catch {
        $xmlParseMode = "tolerant"
        Write-Log ("UI dump is not valid XML on this device. Falling back to tolerant parsing. {0}" -f $_.Exception.Message)
        $nodes = Get-UiNodesFromRawDump -XmlText $xmlText
    }

    $nodeArray = @($nodes | Where-Object { $null -ne $_ })

    $script:LastUiSnapshot = [pscustomobject]@{
        DumpedAt  = Get-Date
        XmlText   = $xmlText
        ParseMode = $xmlParseMode
        NodeCount = $nodeArray.Count
        Nodes     = $nodeArray
    }

    return $script:LastUiSnapshot
}

function Invalidate-UiSnapshot {
    $script:LastUiSnapshot = $null
}

function Test-StringEquals {
    param(
        [string]$Actual,
        [string]$Expected
    )

    return $Actual -eq $Expected
}

function Test-StringContains {
    param(
        [string]$Actual,
        [string]$Expected
    )

    return (-not [string]::IsNullOrWhiteSpace($Actual)) -and ($Actual.IndexOf($Expected, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
}

function Test-UiNodeMatches {
    param(
        [pscustomobject]$Node,
        [hashtable]$Selector
    )

    if ($Selector.ContainsKey("TextEquals") -and -not (Test-StringEquals -Actual $Node.Text -Expected $Selector.TextEquals)) {
        return $false
    }
    if ($Selector.ContainsKey("TextContains") -and -not (Test-StringContains -Actual $Node.Text -Expected $Selector.TextContains)) {
        return $false
    }
    if ($Selector.ContainsKey("DescEquals") -and -not (Test-StringEquals -Actual $Node.Desc -Expected $Selector.DescEquals)) {
        return $false
    }
    if ($Selector.ContainsKey("DescContains") -and -not (Test-StringContains -Actual $Node.Desc -Expected $Selector.DescContains)) {
        return $false
    }
    if ($Selector.ContainsKey("ResourceIdContains") -and -not (Test-StringContains -Actual $Node.ResourceId -Expected $Selector.ResourceIdContains)) {
        return $false
    }
    if ($Selector.ContainsKey("ResourceIdEquals") -and -not (Test-StringEquals -Actual $Node.ResourceId -Expected $Selector.ResourceIdEquals)) {
        return $false
    }
    if ($Selector.ContainsKey("PackageEquals") -and -not (Test-StringEquals -Actual $Node.Package -Expected $Selector.PackageEquals)) {
        return $false
    }

    return $true
}

function Get-NodeSortTuple {
    param(
        [pscustomobject]$Node,
        [hashtable]$Selector
    )

    $clickableScore = if ($Selector.PreferClickable -and $Node.Clickable -eq "true") { 1 } else { 0 }
    $enabledScore = if ($Node.Enabled -eq "true") { 1 } else { 0 }
    $area = $Node.Bounds.Width * $Node.Bounds.Height

    switch ($Selector.Sort) {
        "RightMost" { return @(-$clickableScore, -$enabledScore, -$Node.Bounds.Center.X, $Node.Bounds.Center.Y, -$area) }
        "BottomMost" { return @(-$clickableScore, -$enabledScore, -$Node.Bounds.Center.Y, -$Node.Bounds.Center.X, -$area) }
        "Largest" { return @(-$clickableScore, -$enabledScore, -$area, $Node.Bounds.Center.Y, $Node.Bounds.Center.X) }
        default { return @(-$clickableScore, -$enabledScore, $Node.Bounds.Center.Y, $Node.Bounds.Center.X, -$area) }
    }
}

function Select-BestUiNode {
    param(
        [pscustomobject[]]$Nodes,
        [hashtable]$Selector
    )

    if (-not $Nodes -or $Nodes.Count -eq 0) {
        return $null
    }

    return $Nodes |
        Sort-Object `
            @{ Expression = { (Get-NodeSortTuple -Node $_ -Selector $Selector)[0] } }, `
            @{ Expression = { (Get-NodeSortTuple -Node $_ -Selector $Selector)[1] } }, `
            @{ Expression = { (Get-NodeSortTuple -Node $_ -Selector $Selector)[2] } }, `
            @{ Expression = { (Get-NodeSortTuple -Node $_ -Selector $Selector)[3] } }, `
            @{ Expression = { (Get-NodeSortTuple -Node $_ -Selector $Selector)[4] } } |
        Select-Object -First 1
}

function Test-BoundsContainsPoint {
    param(
        [pscustomobject]$Bounds,
        [int]$X,
        [int]$Y
    )

    return $Bounds -and `
        $X -ge $Bounds.Left -and `
        $X -le $Bounds.Right -and `
        $Y -ge $Bounds.Top -and `
        $Y -le $Bounds.Bottom
}

function Get-BoundsArea {
    param([pscustomobject]$Bounds)

    if (-not $Bounds) {
        return 0
    }

    return $Bounds.Width * $Bounds.Height
}

function Get-UnionBounds {
    param(
        [pscustomobject]$First,
        [pscustomobject]$Second
    )

    if (-not $First) {
        return $Second
    }
    if (-not $Second) {
        return $First
    }

    return Get-UiBounds -BoundsText ("[{0},{1}][{2},{3}]" -f `
        [Math]::Min($First.Left, $Second.Left), `
        [Math]::Min($First.Top, $Second.Top), `
        [Math]::Max($First.Right, $Second.Right), `
        [Math]::Max($First.Bottom, $Second.Bottom))
}

function Convert-BoundsToTapPoint {
    param(
        [string]$ActionName,
        [pscustomobject]$Bounds,
        [string]$Source
    )

    if (-not $Bounds) {
        return $null
    }

    return [pscustomobject]@{
        ActionName = $ActionName
        X          = $Bounds.Center.X
        Y          = $Bounds.Center.Y
        Source     = $Source
        MatchCount = 1
    }
}

function Select-ContainingNode {
    param(
        [pscustomobject[]]$Nodes,
        [pscustomobject]$FocusBounds,
        [switch]$PreferClickable
    )

    if (-not $FocusBounds) {
        return $null
    }

    $centerX = $FocusBounds.Center.X
    $centerY = $FocusBounds.Center.Y
    $candidates = @(
        $Nodes | Where-Object {
            $_.Enabled -eq "true" -and
            (-not $PreferClickable -or $_.Clickable -eq "true") -and
            (Test-BoundsContainsPoint -Bounds $_.Bounds -X $centerX -Y $centerY) -and
            (Get-BoundsArea -Bounds $_.Bounds) -ge (Get-BoundsArea -Bounds $FocusBounds)
        }
    )

    if ($candidates.Count -eq 0 -and $PreferClickable) {
        return Select-ContainingNode -Nodes $Nodes -FocusBounds $FocusBounds
    }

    return $candidates |
        Sort-Object `
            @{ Expression = { Get-BoundsArea -Bounds $_.Bounds } }, `
            @{ Expression = { $_.Bounds.Top } }, `
            @{ Expression = { $_.Bounds.Left } } |
        Select-Object -First 1
}

function Find-HeuristicTapPoint {
    param(
        [string]$ActionName,
        [pscustomobject]$Snapshot,
        [pscustomobject]$AnchorPoint
    )

    $nodes = @($Snapshot.Nodes)
    if ($nodes.Count -eq 0) {
        return $null
    }

    switch ($ActionName) {
        "RechargeCard" {
            $titleNode = @($nodes | Where-Object { Test-StringContains -Actual $_.Text -Expected $TextLeBiRecharge }) |
                Sort-Object @{ Expression = { $_.Bounds.Top } }, @{ Expression = { $_.Bounds.Left } } |
                Select-Object -First 1
            if (-not $titleNode) {
                return $null
            }

            $focusBounds = $titleNode.Bounds
            $subtitleNode = @(
                $nodes | Where-Object {
                    (Test-StringContains -Actual $_.Text -Expected $TextQuickRechargeEntry) -and
                    $_.Bounds.Top -ge $titleNode.Bounds.Top -and
                    $_.Bounds.Top -le ($titleNode.Bounds.Bottom + [Math]::Max(140, [int]($script:DeviceProfile.Height * 0.08)))
                }
            ) | Sort-Object @{ Expression = { $_.Bounds.Top } }, @{ Expression = { $_.Bounds.Left } } | Select-Object -First 1

            if ($subtitleNode) {
                $focusBounds = Get-UnionBounds -First $focusBounds -Second $subtitleNode.Bounds
            }

            $container = Select-ContainingNode -Nodes $nodes -FocusBounds $focusBounds -PreferClickable
            if ($container) {
                return Convert-BoundsToTapPoint -ActionName $ActionName -Bounds $container.Bounds -Source "heuristic:recharge-card-cluster"
            }

            return Convert-BoundsToTapPoint -ActionName $ActionName -Bounds $focusBounds -Source "heuristic:recharge-card-text"
        }
        "ServiceExchange" {
            $markerNode = @(
                $nodes | Where-Object {
                    (Test-StringContains -Actual $_.Text -Expected $TextCoinWelfare) -or
                    (Test-StringContains -Actual $_.Text -Expected $TextExchangeRateHint) -or
                    (Test-StringContains -Actual $_.Desc -Expected $TextCoinWelfare)
                }
            ) | Sort-Object @{ Expression = { $_.Bounds.Top } }, @{ Expression = { -$_.Bounds.Width } } | Select-Object -First 1

            if (-not $markerNode) {
                return $null
            }

            $candidate = @(
                $nodes | Where-Object {
                    $_.Enabled -eq "true" -and
                    $_.Bounds.Center.Y -ge ($markerNode.Bounds.Top - 30) -and
                    $_.Bounds.Center.Y -le ($markerNode.Bounds.Bottom + 30) -and
                    $_.Bounds.Left -ge ($markerNode.Bounds.Center.X - 10)
                }
            ) | Sort-Object `
                @{ Expression = { if ($_.Clickable -eq "true") { 0 } else { 1 } } }, `
                @{ Expression = { -$_.Bounds.Right } }, `
                @{ Expression = { Get-BoundsArea -Bounds $_.Bounds } } |
                Select-Object -First 1

            if ($candidate) {
                return Convert-BoundsToTapPoint -ActionName $ActionName -Bounds $candidate.Bounds -Source "heuristic:service-row-right"
            }

            return [pscustomobject]@{
                ActionName = $ActionName
                X          = [int][Math]::Round($markerNode.Bounds.Right - [Math]::Max(48 * $script:DeviceProfile.ScaleX, $script:DeviceProfile.Width * 0.08))
                Y          = $markerNode.Bounds.Center.Y
                Source     = "heuristic:service-row-edge"
                MatchCount = 1
            }
        }
        "PopupMinus" {
            $leBiUnitPattern = [regex]::Escape($TextExchangeLeBi.Substring(2))
            $strictAmountPattern = '^\d+\s*' + $leBiUnitPattern + '$'
            $looseAmountPattern = '^\d+.*' + $leBiUnitPattern + '$'
            $amountNode = @(
                $nodes | Where-Object {
                    ($_.Text -match $strictAmountPattern) -and
                    $_.Bounds.Top -ge ($script:DeviceProfile.Height * 0.35)
                }
            ) | Sort-Object @{ Expression = { $_.Bounds.Top } }, @{ Expression = { $_.Bounds.Left } } | Select-Object -First 1

            if (-not $amountNode) {
                $amountNode = @(
                    $nodes | Where-Object {
                        ($_.Text -match $looseAmountPattern) -and
                        $_.Bounds.Top -ge ($script:DeviceProfile.Height * 0.35)
                    }
                ) | Sort-Object @{ Expression = { $_.Bounds.Top } }, @{ Expression = { $_.Bounds.Left } } | Select-Object -First 1
            }

            if (-not $amountNode) {
                return $null
            }

            $candidate = @(
                $nodes | Where-Object {
                    $_.Enabled -eq "true" -and
                    $_.Bounds.Right -le ($amountNode.Bounds.Left + 10) -and
                    [Math]::Abs($_.Bounds.Center.Y - $amountNode.Bounds.Center.Y) -le [Math]::Max(40, $amountNode.Bounds.Height)
                }
            ) | Sort-Object `
                @{ Expression = { if ($_.Clickable -eq "true") { 0 } else { 1 } } }, `
                @{ Expression = { -$_.Bounds.Right } }, `
                @{ Expression = { Get-BoundsArea -Bounds $_.Bounds } } |
                Select-Object -First 1

            if ($candidate) {
                return Convert-BoundsToTapPoint -ActionName $ActionName -Bounds $candidate.Bounds -Source "heuristic:stepper-left"
            }

            return [pscustomobject]@{
                ActionName = $ActionName
                X          = [int][Math]::Round([Math]::Max(1, $amountNode.Bounds.Left - [Math]::Max(52 * $script:DeviceProfile.ScaleX, $amountNode.Bounds.Width * 0.18)))
                Y          = $amountNode.Bounds.Center.Y
                Source     = "heuristic:stepper-left-offset"
                MatchCount = 1
            }
        }
        "PopupPlus" {
            $leBiUnitPattern = [regex]::Escape($TextExchangeLeBi.Substring(2))
            $strictAmountPattern = '^\d+\s*' + $leBiUnitPattern + '$'
            $looseAmountPattern = '^\d+.*' + $leBiUnitPattern + '$'
            $amountNode = @(
                $nodes | Where-Object {
                    ($_.Text -match $strictAmountPattern) -and
                    $_.Bounds.Top -ge ($script:DeviceProfile.Height * 0.35)
                }
            ) | Sort-Object @{ Expression = { $_.Bounds.Top } }, @{ Expression = { $_.Bounds.Left } } | Select-Object -First 1

            if (-not $amountNode) {
                $amountNode = @(
                    $nodes | Where-Object {
                        ($_.Text -match $looseAmountPattern) -and
                        $_.Bounds.Top -ge ($script:DeviceProfile.Height * 0.35)
                    }
                ) | Sort-Object @{ Expression = { $_.Bounds.Top } }, @{ Expression = { $_.Bounds.Left } } | Select-Object -First 1
            }

            if (-not $amountNode) {
                return $null
            }

            $candidate = @(
                $nodes | Where-Object {
                    $_.Enabled -eq "true" -and
                    $_.Bounds.Left -ge ($amountNode.Bounds.Right - 10) -and
                    [Math]::Abs($_.Bounds.Center.Y - $amountNode.Bounds.Center.Y) -le [Math]::Max(40, $amountNode.Bounds.Height)
                }
            ) | Sort-Object `
                @{ Expression = { if ($_.Clickable -eq "true") { 0 } else { 1 } } }, `
                @{ Expression = { $_.Bounds.Left } }, `
                @{ Expression = { Get-BoundsArea -Bounds $_.Bounds } } |
                Select-Object -First 1

            if ($candidate) {
                return Convert-BoundsToTapPoint -ActionName $ActionName -Bounds $candidate.Bounds -Source "heuristic:stepper-right"
            }

            return [pscustomobject]@{
                ActionName = $ActionName
                X          = [int][Math]::Round([Math]::Min($script:DeviceProfile.Width - 1, $amountNode.Bounds.Right + [Math]::Max(52 * $script:DeviceProfile.ScaleX, $amountNode.Bounds.Width * 0.18)))
                Y          = $amountNode.Bounds.Center.Y
                Source     = "heuristic:stepper-offset"
                MatchCount = 1
            }
        }
        "PopupCancel" {
            $exchangePoint = Resolve-TapPointFromSnapshot -ActionName "PopupExchange" -Snapshot $Snapshot -AllowFallback
            if (-not $exchangePoint) {
                return $null
            }

            return [pscustomobject]@{
                ActionName = $ActionName
                X          = [int][Math]::Round([Math]::Max(1, $exchangePoint.X - ($script:DeviceProfile.Width * 0.36)))
                Y          = $exchangePoint.Y
                Source     = "heuristic:popup-cancel-offset"
                MatchCount = 1
            }
        }
    }

    return $null
}

function Find-UiTargetInSnapshot {
    param(
        [string]$ActionName,
        [pscustomobject]$Snapshot
    )

    $selectors = @($LocatorProfiles[$ActionName])
    if ($selectors.Count -eq 0) {
        return $null
    }

    foreach ($selector in $selectors) {
        $matches = @($snapshot.Nodes | Where-Object { Test-UiNodeMatches -Node $_ -Selector $selector })
        if ($matches.Count -gt 0) {
            $best = Select-BestUiNode -Nodes $matches -Selector $selector
            return [pscustomobject]@{
                ActionName = $ActionName
                X          = $best.Bounds.Center.X
                Y          = $best.Bounds.Center.Y
                Source     = "dynamic:$($selector.Label)"
                MatchCount = $matches.Count
            }
        }
    }

    return $null
}

function Get-ScaledTapPoint {
    param(
        [string]$ActionName,
        [pscustomobject]$AnchorPoint
    )

    if (-not $script:DeviceProfile) {
        throw "Device profile has not been initialized."
    }

    if ($RelativeFallbacks.ContainsKey($ActionName) -and $AnchorPoint) {
        $relative = $RelativeFallbacks[$ActionName]
        return [pscustomobject]@{
            ActionName = $ActionName
            X          = [int][Math]::Round($AnchorPoint.X + ($relative.DeltaX * $script:DeviceProfile.ScaleX))
            Y          = [int][Math]::Round($AnchorPoint.Y + ($relative.DeltaY * $script:DeviceProfile.ScaleY))
            Source     = "anchored-fallback:$($relative.Anchor)"
            MatchCount = 0
        }
    }

    $base = $BaseCoords[$ActionName]
    if (-not $base) {
        return $null
    }

    return [pscustomobject]@{
        ActionName = $ActionName
        X          = [int][Math]::Round($base.X * $script:DeviceProfile.ScaleX)
        Y          = [int][Math]::Round($base.Y * $script:DeviceProfile.ScaleY)
        Source     = "scaled-fallback:$($script:DeviceProfile.Resolution)"
        MatchCount = 0
    }
}

function Resolve-TapPoint {
    param(
        [string]$ActionName,
        [switch]$ForceRefresh,
        [switch]$AllowFallback,
        [pscustomobject]$AnchorPoint
    )

    $dynamicPoint = Find-UiTarget -ActionName $ActionName -ForceRefresh:$ForceRefresh
    if ($dynamicPoint) {
        $script:ResolvedTapPoints[$ActionName] = $dynamicPoint
        return $dynamicPoint
    }

    $snapshot = Get-UiSnapshot -ForceRefresh:$ForceRefresh
    $heuristicPoint = Find-HeuristicTapPoint -ActionName $ActionName -Snapshot $snapshot -AnchorPoint $AnchorPoint
    if ($heuristicPoint) {
        $script:ResolvedTapPoints[$ActionName] = $heuristicPoint
        return $heuristicPoint
    }

    if (-not $AllowFallback) {
        return $null
    }

    $fallbackPoint = Get-ScaledTapPoint -ActionName $ActionName -AnchorPoint $AnchorPoint
    if ($fallbackPoint) {
        return $fallbackPoint
    }

    return $null
}

function Write-StageDuration {
    param(
        [string]$StageName,
        [System.Diagnostics.Stopwatch]$Stopwatch
    )

    Write-Log ("{0} duration: {1} ms" -f $StageName, $Stopwatch.ElapsedMilliseconds)
}

function Get-StaticTapPoint {
    param(
        [string]$ActionName,
        [pscustomobject]$FallbackPoint
    )

    if ($script:CalibrationProfile) {
        $calibratedPoints = Get-NamedValue -Container $script:CalibrationProfile -Name "Points"
        $calibratedPoint = Convert-StoredPointToTapPoint -ActionName $ActionName -Point (Get-NamedValue -Container $calibratedPoints -Name $ActionName)
        if ($calibratedPoint) {
            $calibratedPoint.Source = "calibration:$ActionName"
            return $calibratedPoint
        }
    }

    if ($script:FastModeActive -and $script:FastDeviceProfile) {
        $coords = Get-NamedValue -Container (Get-NamedValue -Container $script:FastDeviceProfile -Name "Coords") -Name $ActionName
        if ($coords) {
            return [pscustomobject]@{
                ActionName = $ActionName
                X          = [int](Get-NamedValue -Container $coords -Name "X")
                Y          = [int](Get-NamedValue -Container $coords -Name "Y")
                Source     = "profile:$($script:DeviceProfile.Model)"
                MatchCount = 1
            }
        }
    }

    if ($FallbackPoint) {
        return $FallbackPoint
    }

    return Resolve-TapPoint -ActionName $ActionName -AllowFallback
}

function Get-RehearsalTapPoint {
    param([string]$ActionName)

    if (-not $script:RehearsalCache) {
        return $null
    }

    $point = Convert-StoredPointToTapPoint -ActionName $ActionName -Point (Get-NamedValue -Container (Get-NamedValue -Container $script:RehearsalCache -Name "Points") -Name $ActionName)
    if ($point) {
        $point.Source = "rehearsal-cache:$ActionName"
    }

    return $point
}

function Get-ProfileTapPoint {
    param([string]$ActionName)

    $point = Get-StaticTapPoint -ActionName $ActionName
    if ($point) {
        return $point
    }

    if ($script:FastModeActive -and $script:FastDeviceProfile -and (Get-NamedValue -Container (Get-NamedValue -Container $script:FastDeviceProfile -Name "Coords") -Name $ActionName)) {
        $coords = Get-NamedValue -Container (Get-NamedValue -Container $script:FastDeviceProfile -Name "Coords") -Name $ActionName
        return [pscustomobject]@{
            ActionName = $ActionName
            X          = [int](Get-NamedValue -Container $coords -Name "X")
            Y          = [int](Get-NamedValue -Container $coords -Name "Y")
            Source     = "profile:$($script:DeviceProfile.Model)"
            MatchCount = 1
        }
    }

    return $null
}

function Try-Resolve-QuickPoint {
    param(
        [string]$ActionName,
        [int]$TimeoutMs = $FastTiming.QuickCheckBudgetMs,
        [pscustomobject]$AnchorPoint,
        [switch]$AllowFallback
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    do {
        $point = Resolve-TapPoint -ActionName $ActionName -ForceRefresh -AllowFallback:$false -AnchorPoint $AnchorPoint
        if ($point) {
            return $point
        }

        if ($stopwatch.ElapsedMilliseconds -ge $TimeoutMs) {
            break
        }

        Start-Sleep -Milliseconds $FastTiming.QuickPollMs
    }
    while ($true)

    if ($AllowFallback) {
        return Resolve-TapPoint -ActionName $ActionName -ForceRefresh -AllowFallback -AnchorPoint $AnchorPoint
    }

    return $null
}

function Find-UiTarget {
    param(
        [string]$ActionName,
        [switch]$ForceRefresh
    )

    $snapshot = Get-UiSnapshot -ForceRefresh:$ForceRefresh
    return Find-UiTargetInSnapshot -ActionName $ActionName -Snapshot $snapshot
}

function Resolve-TapPointFromSnapshot {
    param(
        [string]$ActionName,
        [pscustomobject]$Snapshot,
        [switch]$AllowFallback,
        [pscustomobject]$AnchorPoint
    )

    $dynamicPoint = Find-UiTargetInSnapshot -ActionName $ActionName -Snapshot $Snapshot
    if ($dynamicPoint) {
        return $dynamicPoint
    }

    $heuristicPoint = Find-HeuristicTapPoint -ActionName $ActionName -Snapshot $Snapshot -AnchorPoint $AnchorPoint
    if ($heuristicPoint) {
        return $heuristicPoint
    }

    if ($AllowFallback) {
        return Get-ScaledTapPoint -ActionName $ActionName -AnchorPoint $AnchorPoint
    }

    return $null
}

function Build-TurboSequenceCommand {
    param(
        [pscustomobject]$RechargePoint,
        [pscustomobject]$ServicePoint,
        [int]$RechargeDelayMs,
        [int]$PopupDelayMs,
        [int]$PreBurstDelayMs,
        [pscustomobject]$PlusPoint,
        [pscustomobject]$ExchangePoint
    )

    $sleepRecharge = Get-ShellSecondsText -Milliseconds $RechargeDelayMs
    $sleepPopup    = Get-ShellSecondsText -Milliseconds $PopupDelayMs

    $exchangeCommand = if ($BurstOnly) { ':' } else {
        'cmd input tap {0} {1}' -f $ExchangePoint.X, $ExchangePoint.Y
    }

    $burstSegment = if ($PlusTapCount -gt 0) {
        'count=0; while [ $count -lt {0} ]; do cmd input tap {1} {2}; count=$((count+1)); done' -f `
            $PlusTapCount, $PlusPoint.X, $PlusPoint.Y
    } else { ':' }

    $segments = [System.Collections.Generic.List[string]]::new()
    $segments.Add('tapStart=$(date +%s%3N)')
    $segments.Add(('cmd input tap {0} {1}' -f $RechargePoint.X, $RechargePoint.Y))
    $segments.Add(('sleep {0}' -f $sleepRecharge))
    $segments.Add('serviceAt=$(date +%s%3N)')
    $segments.Add(('cmd input tap {0} {1}' -f $ServicePoint.X, $ServicePoint.Y))
    $segments.Add(('sleep {0}' -f $sleepPopup))
    if ($PreBurstDelayMs -gt 0) {
        $segments.Add(('sleep {0}' -f (Get-ShellSecondsText -Milliseconds $PreBurstDelayMs)))
    }
    $segments.Add('popupAt=$(date +%s%3N)')
    $segments.Add($burstSegment)
    $segments.Add('plusEnd=$(date +%s%3N)')
    $segments.Add($exchangeCommand)
    $segments.Add('end=$(date +%s%3N)')
    $segments.Add('echo "$tapStart $serviceAt $popupAt $plusEnd $end"')

    return ($segments -join '; ')
}

function Run-ExchangeFlowFast {
    param([pscustomobject]$RechargePoint)

    $overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $rechargeDelayMs = Get-ConfiguredDelayMs -DelayName "RechargeToServiceTapDelayMs" -DefaultValue $FastTiming.RechargeToServiceMs -OverrideValue $RechargeToServiceTapDelayMs
    $popupDelayMs    = Get-ConfiguredDelayMs -DelayName "ServiceToPopupTapDelayMs"    -DefaultValue $FastTiming.ServiceToPopupMs    -OverrideValue $ServiceToPopupTapDelayMs
    $preBurstDelayMs = Get-ConfiguredDelayMs -DelayName "PreBurstSettleMs"            -DefaultValue $FastTiming.PreBurstSettleMs    -OverrideValue $PreBurstSettleMs

    $rechargeExecutionPoint = Get-StaticTapPoint -ActionName "RechargeCard" -FallbackPoint $RechargePoint
    $servicePoint           = Get-StaticTapPoint -ActionName "ServiceExchange"
    $popupExchangePoint     = Get-StaticTapPoint -ActionName "PopupExchange"
    $popupPlusPoint         = Get-StaticTapPoint -ActionName "PopupPlus" -FallbackPoint (Get-ScaledTapPoint -ActionName "PopupPlus" -AnchorPoint $popupExchangePoint)

    Write-Log ("TurboMode tap points: RechargeCard ({0},{1}) via {2}; ServiceExchange ({3},{4}) via {5}." -f `
        $rechargeExecutionPoint.X, $rechargeExecutionPoint.Y, $rechargeExecutionPoint.Source, `
        $servicePoint.X, $servicePoint.Y, $servicePoint.Source)
    Write-Log ("TurboMode popup points: PopupExchange via {0}; PopupPlus via {1}." -f `
        $popupExchangePoint.Source, $popupPlusPoint.Source)

    if ($DryRun) {
        Write-Log "DryRun: full turbo sequence skipped."
        Write-StageDuration -StageName "Final total duration" -Stopwatch $overallStopwatch
        Write-Log "TurboMode path completed."
        return
    }

    $actionText = if ($BurstOnly) { "without the final exchange tap" } else { "and then tapping PopupExchange" }
    Write-Log ("Running full turbo sequence to {0} lebi with {1} plus tap(s) {2}." -f $TargetLeBi, $PlusTapCount, $actionText)

    $shellCmd = Build-TurboSequenceCommand `
        -RechargePoint   $rechargeExecutionPoint `
        -ServicePoint    $servicePoint `
        -RechargeDelayMs $rechargeDelayMs `
        -PopupDelayMs    $popupDelayMs `
        -PreBurstDelayMs $preBurstDelayMs `
        -PlusPoint       $popupPlusPoint `
        -ExchangePoint   $popupExchangePoint

    $raw = (Invoke-DeviceShell -Command $shellCmd).Trim()

    if ($raw -match '(\d{13})\s+(\d{13})\s+(\d{13})\s+(\d{13})\s+(\d{13})') {
        $tapStartMs  = [int64]$matches[1]
        $serviceAtMs = [int64]$matches[2]
        $popupAtMs   = [int64]$matches[3]
        $plusEndMs   = [int64]$matches[4]
        $endMs       = [int64]$matches[5]

        Write-Log ("Device-side RechargeCard -> ServiceExchange delay: {0} ms" -f ($serviceAtMs - $tapStartMs))
        Write-Log ("Device-side ServiceExchange -> Popup burst start: {0} ms"  -f ($popupAtMs   - $serviceAtMs))
        Write-Log ("Device-side burst +{0} taps: {1} ms"                       -f $PlusTapCount, ($plusEndMs - $popupAtMs))
        Write-Log ("Device-side full sequence: {0} ms"                         -f ($endMs - $tapStartMs))
    } elseif (-not [string]::IsNullOrWhiteSpace($raw)) {
        Write-Log "Turbo sequence raw output: $raw"
    }

    Write-StageDuration -StageName "Final total duration" -Stopwatch $overallStopwatch
    Write-Log "TurboMode path completed."
}

function Wait-ForTapPoint {
    param(
        [string]$ActionName,
        [int]$TimeoutMs,
        [int]$PollMs = $Timing.UiPollMs,
        [switch]$AllowFallback,
        [pscustomobject]$AnchorPoint
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    do {
        $point = Resolve-TapPoint -ActionName $ActionName -ForceRefresh -AllowFallback:$false
        if ($point) {
            return $point
        }

        Start-Sleep -Milliseconds $PollMs
    }
    while ($stopwatch.ElapsedMilliseconds -lt $TimeoutMs)

    if ($AllowFallback) {
        return Resolve-TapPoint -ActionName $ActionName -ForceRefresh -AllowFallback -AnchorPoint $AnchorPoint
    }

    return $null
}

function Wait-ForUiMarker {
    param(
        [string]$MarkerName,
        [int]$TimeoutMs,
        [int]$PollMs = $Timing.UiPollMs
    )

    return Wait-ForTapPoint -ActionName $MarkerName -TimeoutMs $TimeoutMs -PollMs $PollMs
}

function Write-TimeGateStatus {
    $currentText = Get-DeviceTimeText
    $timeMatch = [regex]::Match($currentText, '\b\d{2}:\d{2}:\d{2}\b')
    if (-not $timeMatch.Success) {
        throw "Could not parse device time from adb output: '$currentText'"
    }

    $currentText = $timeMatch.Value
    $currentTime = [TimeSpan]::Parse($currentText)

    if ($currentTime -lt $TargetTime) {
        Write-Log "Device time is $currentText. Waiting until $TargetTimeText before starting."
    }
    else {
        Write-Log "Device time is $currentText. Execution gate $TargetTimeText already passed."
    }
}

function Wait-UntilTargetTime {
    if ($SkipTimeGate) {
        Write-Log "SkipTimeGate is on. The noon wait is disabled."
        return
    }

    Write-TimeGateStatus

    $timeGateCommand = 'while true; do now=$(date +%H%M%S); if [ "$now" -ge {0} ]; then break; fi; sleep {1}; done' -f `
        $TargetTimeCompact, `
        (Get-ShellSecondsText -Milliseconds $Timing.TimeGatePollMs)

    if ($DryRun) {
        Write-Log "DryRun: device-side time gate is skipped."
        return
    }

    [void](Invoke-DeviceShell -Command $timeGateCommand)
}

function Invoke-DeviceTap {
    param(
        [int]$X,
        [int]$Y
    )

    if ($DryRun) {
        Write-Log "DryRun: skipping tap at ($X, $Y)."
        return
    }

    [void](Invoke-DeviceShell -Command ("cmd input tap {0} {1}" -f $X, $Y))
}

function Invoke-DeviceBack {
    if ($DryRun) {
        Write-Log "DryRun: skipping back keyevent."
        return
    }

    [void](Invoke-DeviceShell -Command "input keyevent 4")
    Invalidate-UiSnapshot
}

function Invoke-TapAction {
    param(
        [string]$ActionName,
        [pscustomobject]$Point
    )

    if (-not $Point) {
        throw "Tap point for '$ActionName' could not be resolved."
    }

    Write-Log ("Tapping {0} at ({1}, {2}) via {3}." -f $ActionName, $Point.X, $Point.Y, $Point.Source)
    Invoke-DeviceTap -X $Point.X -Y $Point.Y
    Invalidate-UiSnapshot
}

function Build-PopupBurstCommand {
    param(
        [pscustomobject]$PlusPoint,
        [pscustomobject]$ExchangePoint
    )

    $exchangeCommand = if ($BurstOnly) {
        ':'
    }
    else {
        'cmd input tap {0} {1}' -f $ExchangePoint.X, $ExchangePoint.Y
    }

    $segments = @(
        'start=$(date +%s%3N)',
        'plusStart=$start',
        'count=0',
        ('while [ $count -lt {0} ]; do cmd input tap {1} {2}; count=$((count+1)); done' -f $PlusTapCount, $PlusPoint.X, $PlusPoint.Y),
        'plusEnd=$(date +%s%3N)',
        $exchangeCommand,
        'end=$(date +%s%3N)',
        'echo "$start $plusStart $plusEnd $end"'
    )

    return ($segments -join '; ')
}

function Run-PopupBurst {
    param(
        [pscustomobject]$PlusPoint,
        [pscustomobject]$ExchangePoint
    )

    $actionText = if ($BurstOnly) { "without the final exchange tap" } else { "and then tapping PopupExchange" }
    Write-Log "Running popup burst to $TargetLeBi lebi with $PlusTapCount plus tap(s) $actionText."

    if ($DryRun) {
        Write-Log "DryRun: popup burst is skipped."
        return
    }

    $raw = (Invoke-DeviceShell -Command (Build-PopupBurstCommand -PlusPoint $PlusPoint -ExchangePoint $ExchangePoint)).Trim()
    if ($raw -match '(\d{13})\s+(\d{13})\s+(\d{13})\s+(\d{13})') {
        $startMs = [int64]$matches[1]
        $plusStartMs = [int64]$matches[2]
        $plusEndMs = [int64]$matches[3]
        $endMs = [int64]$matches[4]

        Write-Log "Popup burst start-to-plus duration on device: $($plusEndMs - $plusStartMs) ms"
        if (-not $BurstOnly) {
            Write-Log "Full popup burst duration on device: $($endMs - $startMs) ms"
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($raw)) {
        Write-Log "Popup burst raw output: $raw"
    }
}

function Resolve-RehearsalPopupPoints {
    $popupExchangePoint = Wait-ForTapPoint -ActionName "PopupExchange" -TimeoutMs $Timing.ServicePopupTimeoutMs -PollMs $FastTiming.QuickPollMs -AllowFallback
    if (-not $popupExchangePoint) {
        throw "PopupExchange could not be located during rehearsal."
    }

    $popupPlusPoint = if ($PlusTapCount -gt 0) {
        Wait-ForTapPoint -ActionName "PopupPlus" -TimeoutMs $Timing.PopupTitleTimeoutMs -PollMs $FastTiming.QuickPollMs -AllowFallback -AnchorPoint $popupExchangePoint
    }
    else {
        Get-ScaledTapPoint -ActionName "PopupPlus" -AnchorPoint $popupExchangePoint
    }
    if (-not $popupPlusPoint) {
        throw "PopupPlus could not be located during rehearsal."
    }

    $popupMinusPoint = Resolve-TapPoint -ActionName "PopupMinus" -ForceRefresh -AllowFallback
    $popupCancelPoint = Resolve-TapPoint -ActionName "PopupCancel" -ForceRefresh -AllowFallback -AnchorPoint $popupExchangePoint

    return [pscustomobject]@{
        PopupExchange = $popupExchangePoint
        PopupPlus     = $popupPlusPoint
        PopupMinus    = $popupMinusPoint
        PopupCancel   = $popupCancelPoint
    }
}

function Return-ToManualStartPage {
    param([int]$MaxBackCount = 4)

    for ($attempt = 1; $attempt -le $MaxBackCount; $attempt++) {
        $rechargePoint = Resolve-TapPoint -ActionName "RechargeCard" -ForceRefresh -AllowFallback:$false
        if ($rechargePoint) {
            return $rechargePoint
        }

        Invoke-DeviceBack
        Start-Sleep -Milliseconds $Rehearsal.ReturnSettleMs
    }

    $fallbackPoint = Resolve-TapPoint -ActionName "RechargeCard" -ForceRefresh -AllowFallback
    if ($fallbackPoint) {
        return $fallbackPoint
    }

    throw "Could not return to the RechargeCard search results page after rehearsal."
}

function Run-RehearsalFlow {
    param([pscustomobject]$RechargePoint)

    Write-Log "Rehearsal started."
    $rechargeExecutionPoint = Get-StaticTapPoint -ActionName "RechargeCard" -FallbackPoint $RechargePoint
    $rechargeDelaySw = [System.Diagnostics.Stopwatch]::StartNew()
    Invoke-TapAction -ActionName "RechargeCard" -Point $rechargeExecutionPoint
    Write-Log "Rehearsal step: RechargeCard tapped."

    $servicePoint = Wait-ForTapPoint -ActionName "ServiceExchange" -TimeoutMs $Timing.RechargeToServiceTimeoutMs -PollMs $FastTiming.QuickPollMs -AllowFallback
    if (-not $servicePoint) {
        throw "ServiceExchange was not ready during rehearsal."
    }
    $measuredRechargeDelayMs = [int]$rechargeDelaySw.ElapsedMilliseconds
    Write-Log ("Rehearsal measured Recharge -> Service delay: {0} ms" -f $measuredRechargeDelayMs)

    $popupDelaySw = [System.Diagnostics.Stopwatch]::StartNew()
    Invoke-TapAction -ActionName "ServiceExchange" -Point $servicePoint
    Write-Log "Rehearsal step: ServiceExchange tapped."

    $popupPoints = Resolve-RehearsalPopupPoints
    $measuredPopupDelayMs = [int]$popupDelaySw.ElapsedMilliseconds
    Write-Log ("Rehearsal measured Service -> Popup delay: {0} ms" -f $measuredPopupDelayMs)

    if ($popupPoints.PopupPlus) {
        Invoke-TapAction -ActionName "PopupPlus" -Point $popupPoints.PopupPlus
        Write-Log "Rehearsal step: PopupPlus tapped once."
        Start-Sleep -Milliseconds $Rehearsal.PopupRollbackGapMs
    }

    $rollbackCompleted = $false
    if ($popupPoints.PopupMinus) {
        Invoke-TapAction -ActionName "PopupMinus" -Point $popupPoints.PopupMinus
        $rollbackCompleted = $true
    }
    elseif ($popupPoints.PopupCancel) {
        Invoke-TapAction -ActionName "PopupCancel" -Point $popupPoints.PopupCancel
        $rollbackCompleted = $true
    }
    else {
        Invoke-DeviceBack
        $rollbackCompleted = $true
    }

    if ($rollbackCompleted) {
        Write-Log "Rehearsal rollback completed."
    }

    $returnRechargePoint = Return-ToManualStartPage
    Save-RehearsalCache `
        -RechargePoint $rechargeExecutionPoint `
        -ServicePoint $servicePoint `
        -PopupPlusPoint $popupPoints.PopupPlus `
        -PopupExchangePoint $popupPoints.PopupExchange `
        -PopupMinusPoint $popupPoints.PopupMinus `
        -PopupCancelPoint $popupPoints.PopupCancel `
        -MeasuredRechargeDelayMs $measuredRechargeDelayMs `
        -MeasuredPopupDelayMs $measuredPopupDelayMs

    return $returnRechargePoint
}

function Try-Prime-RehearsalCache {
    param([pscustomobject]$RechargePoint)

    if (-not $EnableRehearsal) {
        return $RechargePoint
    }

    if (-not $ManualStartMode) {
        return $RechargePoint
    }

    if (Test-RehearsalCacheValid -Cache $script:RehearsalCache) {
        return $RechargePoint
    }

    $remainingMs = Get-MillisecondsUntilTargetTime
    if ($remainingMs -lt $Rehearsal.LeadTimeMs) {
        Write-Log ("Skipping rehearsal because lead time is only {0} ms." -f $remainingMs)
        return $RechargePoint
    }

    $returnRechargePoint = Run-RehearsalFlow -RechargePoint $RechargePoint
    $script:RehearsalCache = Load-RehearsalCache
    Select-ExecutionMode
    Write-ExecutionModeLog
    return $returnRechargePoint
}

function Run-ExchangeFlowRehearsed {
    $overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $rechargeStage = [System.Diagnostics.Stopwatch]::StartNew()
    $rechargePoint = Get-RehearsalTapPoint -ActionName "RechargeCard"
    $servicePoint = Get-RehearsalTapPoint -ActionName "ServiceExchange"
    $popupPlusPoint = Get-RehearsalTapPoint -ActionName "PopupPlus"
    $popupExchangePoint = Get-RehearsalTapPoint -ActionName "PopupExchange"

    if (-not $rechargePoint -or -not $servicePoint -or -not $popupPlusPoint -or -not $popupExchangePoint) {
        throw "Rehearsal cache is missing one or more required tap points."
    }

    $rechargeDelayMs = Get-ConfiguredDelayMs -DelayName "RechargeToServiceTapDelayMs" -DefaultValue $FastTiming.RechargeToServiceMs -OverrideValue $RechargeToServiceTapDelayMs -IncludeRehearsalSafetyMargin
    $popupDelayMs = Get-ConfiguredDelayMs -DelayName "ServiceToPopupTapDelayMs" -DefaultValue $FastTiming.ServiceToPopupMs -OverrideValue $ServiceToPopupTapDelayMs -IncludeRehearsalSafetyMargin
    $preBurstDelayMs = Get-ConfiguredDelayMs -DelayName "PreBurstSettleMs" -DefaultValue $FastTiming.PreBurstSettleMs -OverrideValue $PreBurstSettleMs

    Invoke-TapAction -ActionName "RechargeCard" -Point $rechargePoint
    Start-Sleep -Milliseconds $rechargeDelayMs
    Write-StageDuration -StageName "Execution stage RechargeCard -> ServiceExchange ready" -Stopwatch $rechargeStage

    $serviceStage = [System.Diagnostics.Stopwatch]::StartNew()
    Invoke-TapAction -ActionName "ServiceExchange" -Point $servicePoint
    Start-Sleep -Milliseconds $popupDelayMs
    if ($preBurstDelayMs -gt 0) {
        Start-Sleep -Milliseconds $preBurstDelayMs
    }
    Write-StageDuration -StageName "Execution stage ServiceExchange -> Popup ready" -Stopwatch $serviceStage

    $burstStage = [System.Diagnostics.Stopwatch]::StartNew()
    Run-PopupBurst -PlusPoint $popupPlusPoint -ExchangePoint $popupExchangePoint
    Write-StageDuration -StageName "Execution stage Popup burst -> Exchange tap" -Stopwatch $burstStage
    Write-StageDuration -StageName "Final total duration" -Stopwatch $overallStopwatch
    Write-Log "Rehearsed TurboMode path completed."
}

function Save-RunDiagnostics {
    param(
        [string]$Reason,
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    try {
        $artifactDir = Ensure-ArtifactDirectory
        $deviceInfoPath = Join-Path $artifactDir "device_profile.json"
        $focusPath = Join-Path $artifactDir "current_focus.txt"
        $errorPath = Join-Path $artifactDir "error.txt"
        $xmlPath = Join-Path $artifactDir "window_dump.xml"
        $screenshotPath = Join-Path $artifactDir "screenshot.png"

        if ($script:DeviceProfile) {
            $script:DeviceProfile | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $deviceInfoPath -Encoding UTF8
        }

        $focus = Get-CurrentFocusText
        if (-not [string]::IsNullOrWhiteSpace($focus)) {
            Set-Content -LiteralPath $focusPath -Value $focus -Encoding UTF8
        }

        try {
            $uiSnapshot = Get-UiSnapshot -ForceRefresh
            Set-Content -LiteralPath $xmlPath -Value $uiSnapshot.XmlText -Encoding UTF8
        }
        catch {
            Set-Content -LiteralPath $xmlPath -Value "UI dump failed: $($_.Exception.Message)" -Encoding UTF8
        }

        try {
            Invoke-AdbToFile -Arguments ((Get-AdbDeviceArgs) + @("exec-out", "screencap", "-p")) -OutputPath $screenshotPath -AllowFailure
        }
        catch {
            Set-Content -LiteralPath $screenshotPath -Value "Screenshot failed: $($_.Exception.Message)" -Encoding UTF8
        }

        $errorLines = New-Object System.Collections.Generic.List[string]
        $errorLines.Add("Reason: $Reason")
        if ($ErrorRecord) {
            $errorLines.Add("Message: $($ErrorRecord.Exception.Message)")
            $errorLines.Add("ScriptStackTrace:")
            $errorLines.Add($ErrorRecord.ScriptStackTrace)
        }
        Set-Content -LiteralPath $errorPath -Value $errorLines -Encoding UTF8

        Write-Log "Diagnostics saved to $artifactDir"
    }
    catch {
        Write-Log "Failed to save diagnostics: $($_.Exception.Message)"
    }
}

function Resolve-RequiredPopupPoints {
    param([pscustomobject]$PopupExchangePoint)

    $popupTitle = Find-UiTarget -ActionName "PopupTitle"
    if (-not $popupTitle) {
        Write-Log "Popup title marker was not detected within the expected window. Continuing with button lookup."
    }

    if (-not $PopupExchangePoint) {
        $PopupExchangePoint = Resolve-TapPoint -ActionName "PopupExchange" -AllowFallback
    }

    if (-not $PopupExchangePoint) {
        throw "PopupExchange could not be located after tapping ServiceExchange."
    }

    $popupPlusPoint = if ($PlusTapCount -gt 0) {
        Resolve-TapPoint -ActionName "PopupPlus" -AllowFallback -AnchorPoint $PopupExchangePoint
    }
    else {
        Get-ScaledTapPoint -ActionName "PopupPlus" -AnchorPoint $PopupExchangePoint
    }

    if (-not $popupPlusPoint) {
        throw "PopupPlus could not be located after opening the popup."
    }

    Write-Log ("Resolved PopupExchange via {0}; PopupPlus via {1}." -f $PopupExchangePoint.Source, $popupPlusPoint.Source)
    return [pscustomobject]@{
        PopupExchange = $PopupExchangePoint
        PopupPlus     = $popupPlusPoint
    }
}

function Invoke-ServiceExchangeWithRetry {
    $servicePoint = Resolve-TapPoint -ActionName "ServiceExchange" -ForceRefresh -AllowFallback
    if (-not $servicePoint) {
        Start-Sleep -Milliseconds $Timing.ServiceRetryGapMs
        $servicePoint = Resolve-TapPoint -ActionName "ServiceExchange" -ForceRefresh -AllowFallback
    }
    if (-not $servicePoint) {
        throw "ServiceExchange did not appear after tapping RechargeCard."
    }

    for ($attempt = 1; $attempt -le 2; $attempt++) {
        Invoke-TapAction -ActionName "ServiceExchange" -Point $servicePoint
        Start-Sleep -Milliseconds $Timing.PopupOpenSettleMs

        $popupExchangePoint = Resolve-TapPoint -ActionName "PopupExchange" -ForceRefresh -AllowFallback
        if ($popupExchangePoint) {
            return Resolve-RequiredPopupPoints -PopupExchangePoint $popupExchangePoint
        }

        if ($attempt -lt 2) {
            Write-Log "PopupExchange was not detected after the first tap. Retrying ServiceExchange once."
            Start-Sleep -Milliseconds $Timing.ServiceRetryGapMs
        }
    }

    throw "PopupExchange did not appear after retrying ServiceExchange."
}

function Open-RechargeCardAndWaitForServicePage {
    param([pscustomobject]$RechargePoint)

    for ($attempt = 1; $attempt -le 2; $attempt++) {
        Invoke-TapAction -ActionName "RechargeCard" -Point $RechargePoint
        Start-Sleep -Milliseconds $Timing.RechargeOpenSettleMs
        $servicePageMarker = Resolve-TapPoint -ActionName "ServicePageMarker" -ForceRefresh
        if ($servicePageMarker) {
            Write-Log "Service page marker detected."
            return
        }

        if ($attempt -lt 2) {
            Write-Log "Service page was not detected after tapping RechargeCard. Retrying once."
            Start-Sleep -Milliseconds $Timing.PostTapSettleMs
        }
    }

    throw "The service page did not appear after tapping RechargeCard."
}

function Run-PreflightChecks {
    Resolve-Device
    $script:DeviceProfile = Get-DeviceProfile
    Initialize-ExecutionModes
    $script:CalibrationProfile = Load-CalibrationProfile
    $script:RehearsalCache = Load-RehearsalCache
    Select-ExecutionMode

    Write-Log ("Device profile: {0} {1}, Android {2}, resolution {3}, density {4}" -f `
        $script:DeviceProfile.Manufacturer, `
        $script:DeviceProfile.Model, `
        $script:DeviceProfile.Android, `
        $script:DeviceProfile.Resolution, `
        ($(if (-not [string]::IsNullOrWhiteSpace($script:DeviceProfile.Density)) { $script:DeviceProfile.Density } else { "unknown" })))
    Write-Log ("Scale factors relative to {0}x{1}: X={2:N3}, Y={3:N3}" -f `
        $BaseResolution.Width, `
        $BaseResolution.Height, `
        $script:DeviceProfile.ScaleX, `
        $script:DeviceProfile.ScaleY)
    Write-ExecutionModeLog

    if ($ManualStartMode) {
        Write-Log "ManualStartMode is on. Expecting the user to already be on QQ Music search results for '乐币' with the RechargeCard visible."
        Write-Log "Foreground assertion, app recovery, and post-exchange completion checks are skipped."
    }
    else {
        $foregroundConfirmed = Assert-QqMusicForeground
        if (-not $foregroundConfirmed) {
            Assert-QqMusicInstalled
        }
    }

    $rechargePoint = Resolve-TapPoint -ActionName "RechargeCard" -ForceRefresh -AllowFallback
    if (-not $rechargePoint) {
        throw "RechargeCard could not be located on the current screen."
    }

    Write-Log ("RechargeCard resolved via {0}. Manual start page assumption is satisfied." -f $rechargePoint.Source)
    return $rechargePoint
}

function Run-ExchangeFlow {
    $rechargePoint = Run-PreflightChecks

    if ($PreflightOnly) {
        Write-Log "PreflightOnly is on. No taps will be sent."
        Write-Log "Run check completed successfully."
        return
    }

    if ($RehearsalOnly) {
        [void](Run-RehearsalFlow -RechargePoint $rechargePoint)
        Write-Log "RehearsalOnly is on. Ending after rehearsal capture."
        return
    }

    $rechargePoint = Try-Prime-RehearsalCache -RechargePoint $rechargePoint

    Wait-UntilTargetTime
    if ($script:ExecutionMode -eq "Rehearsed TurboMode") {
        try {
            Run-ExchangeFlowRehearsed
            Write-Log "Exchange tap was sent. Ending flow immediately without post-exchange verification."
            return
        }
        catch {
            Write-Log ("Rehearsed TurboMode path failed: {0}. Falling back to the next mode." -f $_.Exception.Message)
            $script:RehearsalCache = $null
            Select-ExecutionMode
            Write-ExecutionModeLog
        }
    }

    if ($script:FastModeActive) {
        try {
            Run-ExchangeFlowFast -RechargePoint $rechargePoint
            Write-Log "Exchange tap was sent. Ending flow immediately without post-exchange verification."
            return
        }
        catch {
            Write-Log ("TurboMode path failed: {0}. Falling back to adaptive flow." -f $_.Exception.Message)
            Invalidate-UiSnapshot
            $script:ExecutionMode = "AdaptiveMode"
            Write-ExecutionModeLog
        }
    }

    $adaptiveOverall = [System.Diagnostics.Stopwatch]::StartNew()
    $adaptiveRechargeStage = [System.Diagnostics.Stopwatch]::StartNew()
    Open-RechargeCardAndWaitForServicePage -RechargePoint $rechargePoint
    Write-StageDuration -StageName "Execution stage RechargeCard -> ServiceExchange ready" -Stopwatch $adaptiveRechargeStage

    $adaptivePopupStage = [System.Diagnostics.Stopwatch]::StartNew()
    $popupPoints = Invoke-ServiceExchangeWithRetry
    Write-StageDuration -StageName "Execution stage ServiceExchange -> Popup ready" -Stopwatch $adaptivePopupStage

    $adaptiveBurstStage = [System.Diagnostics.Stopwatch]::StartNew()
    Run-PopupBurst -PlusPoint $popupPoints.PopupPlus -ExchangePoint $popupPoints.PopupExchange
    Write-StageDuration -StageName "Execution stage Popup burst -> Exchange tap" -Stopwatch $adaptiveBurstStage
    Write-StageDuration -StageName "Final total duration" -Stopwatch $adaptiveOverall
    Write-Log "Exchange tap was sent. Ending flow immediately without post-exchange verification."
}

if ($ListDevicesJson) {
    $deviceInfos = @(Get-ConnectedDeviceInfos)
    if ($deviceInfos.Count -eq 0) {
        Write-Output "[]"
    }
    else {
        $deviceInfos | ConvertTo-Json -Depth 4
    }
    exit 0
}

try {
    Run-ExchangeFlow
    exit 0
}
catch {
    Save-RunDiagnostics -Reason "Script failure" -ErrorRecord $_
    Write-Error $_
    exit 1
}

