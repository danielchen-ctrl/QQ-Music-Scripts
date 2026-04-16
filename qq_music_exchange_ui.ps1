param(
    [switch]$SmokeTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne [System.Threading.ApartmentState]::STA) {
    throw "Please run this UI script with -STA."
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:RootPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:EngineScriptPath = Join-Path $script:RootPath "qq_music_exchange.ps1"
$script:DefaultTargetLeBi = 330
$script:DefaultTargetTimeText = "12:00:00"
$script:CurrentProcess = $null
$script:CurrentRunIsPreflight = $false
$script:CurrentRunIsCalibrate = $false
$script:StdoutPath = $null
$script:StderrPath = $null
$script:StdoutOffset = 0L
$script:StderrOffset = 0L
$script:AvailableDevices = @()

if (-not (Test-Path -LiteralPath $script:EngineScriptPath)) {
    throw "Could not find qq_music_exchange.ps1 next to the UI script."
}

$script:Palette = @{
    WindowBack = [System.Drawing.ColorTranslator]::FromHtml("#F4EFE8")
    CardBack   = [System.Drawing.ColorTranslator]::FromHtml("#FFFDFC")
    CardBorder = [System.Drawing.ColorTranslator]::FromHtml("#D8C9B8")
    HeaderBack = [System.Drawing.ColorTranslator]::FromHtml("#173F43")
    HeaderInk  = [System.Drawing.ColorTranslator]::FromHtml("#F7F3ED")
    Accent     = [System.Drawing.ColorTranslator]::FromHtml("#1F8A70")
    AccentSoft = [System.Drawing.ColorTranslator]::FromHtml("#D9F0E8")
    Ink        = [System.Drawing.ColorTranslator]::FromHtml("#22333B")
    Muted      = [System.Drawing.ColorTranslator]::FromHtml("#6B7280")
    Warning    = [System.Drawing.ColorTranslator]::FromHtml("#C65D3A")
    Success    = [System.Drawing.ColorTranslator]::FromHtml("#1E7F5C")
    LogBack    = [System.Drawing.ColorTranslator]::FromHtml("#132028")
    LogInk     = [System.Drawing.ColorTranslator]::FromHtml("#EAF6F2")
}

$script:TextExecute = -join ([char[]](0x6267, 0x884C))
$script:TextClearLogs = -join ([char[]](0x6E05, 0x7A7A, 0x65E5, 0x5FD7))
$toolTip = New-Object System.Windows.Forms.ToolTip

function New-UiFont {
    param(
        [string]$Name,
        [float]$Size,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular
    )

    return New-Object System.Drawing.Font($Name, $Size, $Style)
}

function Read-NewText {
    param(
        [string]$Path,
        [ref]$Offset
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return ""
    }

    $stream = $null
    $reader = $null

    try {
        $stream = New-Object System.IO.FileStream(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )

        if ($stream.Length -lt $Offset.Value) {
            $Offset.Value = 0L
        }

        [void]$stream.Seek($Offset.Value, [System.IO.SeekOrigin]::Begin)
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true, 1024, $true)
        $text = $reader.ReadToEnd()
        $Offset.Value = $stream.Position
        return $text
    }
    finally {
        if ($reader) {
            $reader.Dispose()
        }
        if ($stream) {
            $stream.Dispose()
        }
    }
}

function Cleanup-RedirectFiles {
    foreach ($path in @($script:StdoutPath, $script:StderrPath)) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
            Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        }
    }

    $script:StdoutPath = $null
    $script:StderrPath = $null
    $script:StdoutOffset = 0L
    $script:StderrOffset = 0L
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "QQ Music Exchange Console"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(940, 720)
$form.MinimumSize = New-Object System.Drawing.Size(940, 720)
$form.BackColor = $script:Palette.WindowBack
$form.ForeColor = $script:Palette.Ink

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(940, 108)
$headerPanel.Anchor = "Top,Left,Right"
$headerPanel.BackColor = $script:Palette.HeaderBack
$form.Controls.Add($headerPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(28, 20)
$titleLabel.Size = New-Object System.Drawing.Size(500, 30)
$titleLabel.Text = "QQ Music Exchange Console"
$titleLabel.Font = New-UiFont -Name "Segoe UI Semibold" -Size 16
$titleLabel.ForeColor = $script:Palette.HeaderInk
$headerPanel.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Location = New-Object System.Drawing.Point(30, 58)
$subtitleLabel.Size = New-Object System.Drawing.Size(620, 24)
$subtitleLabel.Text = "Manually open QQ Music to the '乐币' search results page, then use the device-time trigger."
$subtitleLabel.Font = New-UiFont -Name "Segoe UI" -Size 9.5
$subtitleLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#DCEAE6")
$headerPanel.Controls.Add($subtitleLabel)

$chipLabel = New-Object System.Windows.Forms.Label
$chipLabel.Location = New-Object System.Drawing.Point(718, 32)
$chipLabel.Size = New-Object System.Drawing.Size(170, 36)
$chipLabel.Text = "Device-Time Trigger"
$chipLabel.TextAlign = "MiddleCenter"
$chipLabel.Font = New-UiFont -Name "Segoe UI Semibold" -Size 10
$chipLabel.BackColor = $script:Palette.AccentSoft
$chipLabel.ForeColor = $script:Palette.Accent
$headerPanel.Controls.Add($chipLabel)

$configCard = New-Object System.Windows.Forms.Panel
$configCard.Location = New-Object System.Drawing.Point(24, 126)
$configCard.Size = New-Object System.Drawing.Size(892, 264)
$configCard.Anchor = "Top,Left,Right"
$configCard.BackColor = $script:Palette.CardBack
$configCard.BorderStyle = "FixedSingle"
$form.Controls.Add($configCard)

$targetLabel = New-Object System.Windows.Forms.Label
$targetLabel.Location = New-Object System.Drawing.Point(26, 20)
$targetLabel.Size = New-Object System.Drawing.Size(220, 24)
$targetLabel.Text = "Target Lebi"
$targetLabel.Font = New-UiFont -Name "Segoe UI Semibold" -Size 11
$targetLabel.ForeColor = $script:Palette.Ink
$configCard.Controls.Add($targetLabel)

$targetHint = New-Object System.Windows.Forms.Label
$targetHint.Location = New-Object System.Drawing.Point(28, 116)
$targetHint.Size = New-Object System.Drawing.Size(240, 20)
$targetHint.Text = "Step is 10. Default is 330."
$targetHint.Font = New-UiFont -Name "Segoe UI" -Size 9
$targetHint.ForeColor = $script:Palette.Muted
$configCard.Controls.Add($targetHint)

$targetInput = New-Object System.Windows.Forms.NumericUpDown
$targetInput.Location = New-Object System.Drawing.Point(30, 54)
$targetInput.Size = New-Object System.Drawing.Size(220, 40)
$targetInput.Minimum = 10
$targetInput.Maximum = 100000
$targetInput.Increment = 10
$targetInput.Value = $script:DefaultTargetLeBi
$targetInput.Font = New-UiFont -Name "Segoe UI Semibold" -Size 18
$targetInput.BorderStyle = "FixedSingle"
$configCard.Controls.Add($targetInput)

$timeLabel = New-Object System.Windows.Forms.Label
$timeLabel.Location = New-Object System.Drawing.Point(292, 20)
$timeLabel.Size = New-Object System.Drawing.Size(220, 24)
$timeLabel.Text = "Execute At"
$timeLabel.Font = New-UiFont -Name "Segoe UI Semibold" -Size 11
$timeLabel.ForeColor = $script:Palette.Ink
$configCard.Controls.Add($timeLabel)

$timeHint = New-Object System.Windows.Forms.Label
$timeHint.Location = New-Object System.Drawing.Point(294, 116)
$timeHint.Size = New-Object System.Drawing.Size(250, 20)
$timeHint.Text = "Format HH:mm:ss. Default 12:00:00."
$timeHint.Font = New-UiFont -Name "Segoe UI" -Size 9
$timeHint.ForeColor = $script:Palette.Muted
$configCard.Controls.Add($timeHint)

$timeInput = New-Object System.Windows.Forms.TextBox
$timeInput.Location = New-Object System.Drawing.Point(296, 54)
$timeInput.Size = New-Object System.Drawing.Size(220, 40)
$timeInput.Text = $script:DefaultTargetTimeText
$timeInput.Font = New-UiFont -Name "Consolas" -Size 18
$timeInput.BorderStyle = "FixedSingle"
$configCard.Controls.Add($timeInput)

$deviceLabel = New-Object System.Windows.Forms.Label
$deviceLabel.Location = New-Object System.Drawing.Point(26, 146)
$deviceLabel.Size = New-Object System.Drawing.Size(220, 24)
$deviceLabel.Text = "Target Device"
$deviceLabel.Font = New-UiFont -Name "Segoe UI Semibold" -Size 11
$deviceLabel.ForeColor = $script:Palette.Ink
$configCard.Controls.Add($deviceLabel)

$deviceComboBox = New-Object System.Windows.Forms.ComboBox
$deviceComboBox.Location = New-Object System.Drawing.Point(30, 176)
$deviceComboBox.Size = New-Object System.Drawing.Size(488, 34)
$deviceComboBox.DropDownStyle = "DropDownList"
$deviceComboBox.Font = New-UiFont -Name "Segoe UI" -Size 10
$deviceComboBox.FlatStyle = "Flat"
$configCard.Controls.Add($deviceComboBox)

$refreshDevicesButton = New-Object System.Windows.Forms.Button
$refreshDevicesButton.Location = New-Object System.Drawing.Point(530, 174)
$refreshDevicesButton.Size = New-Object System.Drawing.Size(92, 34)
$refreshDevicesButton.Text = "Refresh"
$refreshDevicesButton.Font = New-UiFont -Name "Segoe UI" -Size 9.5
$refreshDevicesButton.BackColor = $script:Palette.CardBack
$refreshDevicesButton.ForeColor = $script:Palette.Ink
$refreshDevicesButton.FlatStyle = "Flat"
$refreshDevicesButton.FlatAppearance.BorderColor = $script:Palette.CardBorder
$refreshDevicesButton.FlatAppearance.BorderSize = 1
$configCard.Controls.Add($refreshDevicesButton)

$statusTitle = New-Object System.Windows.Forms.Label
$statusTitle.Location = New-Object System.Drawing.Point(556, 20)
$statusTitle.Size = New-Object System.Drawing.Size(140, 24)
$statusTitle.Text = "Status"
$statusTitle.Font = New-UiFont -Name "Segoe UI Semibold" -Size 11
$statusTitle.ForeColor = $script:Palette.Ink
$configCard.Controls.Add($statusTitle)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(560, 56)
$statusLabel.Size = New-Object System.Drawing.Size(190, 32)
$statusLabel.Text = "Idle"
$statusLabel.Font = New-UiFont -Name "Segoe UI Semibold" -Size 16
$statusLabel.ForeColor = $script:Palette.Accent
$configCard.Controls.Add($statusLabel)

$countdownTitle = New-Object System.Windows.Forms.Label
$countdownTitle.Location = New-Object System.Drawing.Point(556, 100)
$countdownTitle.Size = New-Object System.Drawing.Size(140, 24)
$countdownTitle.Text = "Countdown"
$countdownTitle.Font = New-UiFont -Name "Segoe UI Semibold" -Size 11
$countdownTitle.ForeColor = $script:Palette.Ink
$configCard.Controls.Add($countdownTitle)

$countdownLabel = New-Object System.Windows.Forms.Label
$countdownLabel.Location = New-Object System.Drawing.Point(560, 132)
$countdownLabel.Size = New-Object System.Drawing.Size(150, 32)
$countdownLabel.Text = "--:--:--"
$countdownLabel.Font = New-UiFont -Name "Consolas" -Size 16
$countdownLabel.ForeColor = $script:Palette.Muted
$configCard.Controls.Add($countdownLabel)

$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(734, 42)
$startButton.Size = New-Object System.Drawing.Size(126, 54)
$startButton.Text = $script:TextExecute
$startButton.Font = New-UiFont -Name "Segoe UI Semibold" -Size 12
$startButton.BackColor = $script:Palette.Accent
$startButton.ForeColor = [System.Drawing.Color]::White
$startButton.FlatStyle = "Flat"
$startButton.FlatAppearance.BorderSize = 0
$configCard.Controls.Add($startButton)

$clearLogButton = New-Object System.Windows.Forms.Button
$clearLogButton.Location = New-Object System.Drawing.Point(734, 100)
$clearLogButton.Size = New-Object System.Drawing.Size(126, 34)
$clearLogButton.Text = $script:TextClearLogs
$clearLogButton.Font = New-UiFont -Name "Segoe UI" -Size 9.5
$clearLogButton.BackColor = $script:Palette.CardBack
$clearLogButton.ForeColor = $script:Palette.Ink
$clearLogButton.FlatStyle = "Flat"
$clearLogButton.FlatAppearance.BorderColor = $script:Palette.CardBorder
$clearLogButton.FlatAppearance.BorderSize = 1
$configCard.Controls.Add($clearLogButton)

$preflightButton = New-Object System.Windows.Forms.Button
$preflightButton.Location = New-Object System.Drawing.Point(734, 148)
$preflightButton.Size = New-Object System.Drawing.Size(126, 34)
$preflightButton.Text = "Run Check"
$preflightButton.Font = New-UiFont -Name "Segoe UI" -Size 9.5
$preflightButton.BackColor = $script:Palette.CardBack
$preflightButton.ForeColor = $script:Palette.Ink
$preflightButton.FlatStyle = "Flat"
$preflightButton.FlatAppearance.BorderColor = $script:Palette.CardBorder
$preflightButton.FlatAppearance.BorderSize = 1
$configCard.Controls.Add($preflightButton)
[void]$toolTip.SetToolTip($preflightButton, "Verify the current page is already on QQ Music '乐币' search results and the RechargeCard can be resolved, without sending taps.")

$calibrateButton = New-Object System.Windows.Forms.Button
$calibrateButton.Location = New-Object System.Drawing.Point(734, 192)
$calibrateButton.Size = New-Object System.Drawing.Size(126, 48)
$calibrateButton.Text = "校准设备"
$calibrateButton.Font = New-UiFont -Name "Segoe UI Semibold" -Size 10
$calibrateButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#2E6DA4")
$calibrateButton.ForeColor = [System.Drawing.Color]::White
$calibrateButton.FlatStyle = "Flat"
$calibrateButton.FlatAppearance.BorderSize = 0
$configCard.Controls.Add($calibrateButton)
[void]$toolTip.SetToolTip($calibrateButton, "新设备首次使用时点击：自动导航各页面，捕获当前手机的精确坐标，无需时间限制，完成后数据永久保存供后续执行使用。")

$logCard = New-Object System.Windows.Forms.Panel
$logCard.Location = New-Object System.Drawing.Point(24, 412)
$logCard.Size = New-Object System.Drawing.Size(892, 250)
$logCard.Anchor = "Top,Bottom,Left,Right"
$logCard.BackColor = $script:Palette.CardBack
$logCard.BorderStyle = "FixedSingle"
$form.Controls.Add($logCard)

$logTitle = New-Object System.Windows.Forms.Label
$logTitle.Location = New-Object System.Drawing.Point(22, 16)
$logTitle.Size = New-Object System.Drawing.Size(220, 26)
$logTitle.Text = "Run Log"
$logTitle.Font = New-UiFont -Name "Segoe UI Semibold" -Size 11
$logTitle.ForeColor = $script:Palette.Ink
$logCard.Controls.Add($logTitle)

$logSubtitle = New-Object System.Windows.Forms.Label
$logSubtitle.Location = New-Object System.Drawing.Point(24, 44)
$logSubtitle.Size = New-Object System.Drawing.Size(360, 22)
$logSubtitle.Text = "Standard output and errors both appear below."
$logSubtitle.Font = New-UiFont -Name "Segoe UI" -Size 9
$logSubtitle.ForeColor = $script:Palette.Muted
$logCard.Controls.Add($logSubtitle)

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(22, 76)
$logBox.Size = New-Object System.Drawing.Size(846, 202)
$logBox.Anchor = "Top,Bottom,Left,Right"
$logBox.ReadOnly = $true
$logBox.BackColor = $script:Palette.LogBack
$logBox.ForeColor = $script:Palette.LogInk
$logBox.BorderStyle = "FixedSingle"
$logBox.Font = New-UiFont -Name "Consolas" -Size 10
$logCard.Controls.Add($logBox)

$pollTimer = New-Object System.Windows.Forms.Timer
$pollTimer.Interval = 250

function Add-LogText {
    param(
        [string]$Text,
        [System.Drawing.Color]$Color
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return
    }

    $sanitized = $Text -replace "`0", ""
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.SelectionLength = 0
    $logBox.SelectionColor = $Color
    $logBox.AppendText($sanitized)
    $logBox.SelectionColor = $logBox.ForeColor
    $logBox.ScrollToCaret()
}

function Add-LogLine {
    param(
        [string]$Text,
        [System.Drawing.Color]$Color
    )

    Add-LogText -Text ($Text + [Environment]::NewLine) -Color $Color
}

function Invoke-EngineCapture {
    param([string[]]$ExtraArgs)

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()

    try {
        $argumentList = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $script:EngineScriptPath
        ) + $ExtraArgs

        $process = Start-Process `
            -FilePath "powershell.exe" `
            -ArgumentList $argumentList `
            -WorkingDirectory $script:RootPath `
            -WindowStyle Hidden `
            -Wait `
            -PassThru `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath

        $stdout = if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath -Raw -Encoding UTF8 -ErrorAction SilentlyContinue } else { "" }
        $stderr = if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw -Encoding UTF8 -ErrorAction SilentlyContinue } else { "" }

        if ($process.ExitCode -ne 0) {
            $message = if (-not [string]::IsNullOrWhiteSpace($stderr)) { $stderr.Trim() } elseif (-not [string]::IsNullOrWhiteSpace($stdout)) { $stdout.Trim() } else { "Unknown error." }
            throw $message
        }

        return $stdout
    }
    finally {
        Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function Get-SelectedDeviceId {
    if ($deviceComboBox.SelectedIndex -lt 0) {
        return $null
    }
    if ($deviceComboBox.SelectedIndex -ge $script:AvailableDevices.Count) {
        return $null
    }

    return $script:AvailableDevices[$deviceComboBox.SelectedIndex].Id
}

function Set-UiRunningState {
    param([bool]$IsRunning)

    $targetInput.Enabled = -not $IsRunning
    $timeInput.Enabled = -not $IsRunning
    $deviceComboBox.Enabled = (-not $IsRunning) -and ($script:AvailableDevices.Count -gt 0)
    $refreshDevicesButton.Enabled = -not $IsRunning
    $startButton.Enabled = (-not $IsRunning) -and ($script:AvailableDevices.Count -gt 0)
    $preflightButton.Enabled = (-not $IsRunning) -and ($script:AvailableDevices.Count -gt 0)
    $calibrateButton.Enabled = (-not $IsRunning) -and ($script:AvailableDevices.Count -gt 0)
}

function Update-CountdownDisplay {
    $targetTimeText = $timeInput.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($targetTimeText)) {
        $countdownLabel.Text = "--:--:--"
        $countdownLabel.ForeColor = $script:Palette.Muted
        return
    }

    try {
        $targetTime = [TimeSpan]::ParseExact(
            $targetTimeText,
            "hh\:mm\:ss",
            [System.Globalization.CultureInfo]::InvariantCulture
        )
    }
    catch {
        $countdownLabel.Text = "Invalid"
        $countdownLabel.ForeColor = $script:Palette.Warning
        return
    }

    $now = Get-Date
    $targetDateTime = $now.Date.Add($targetTime)
    $remaining = $targetDateTime - $now

    if ($remaining.TotalSeconds -le 0) {
        if ($script:CurrentProcess -and -not $script:CurrentProcess.HasExited) {
            $countdownLabel.Text = "Executing"
            $countdownLabel.ForeColor = $script:Palette.Accent
        }
        else {
            $countdownLabel.Text = "Passed"
            $countdownLabel.ForeColor = $script:Palette.Warning
        }
        return
    }

    $remainingWhole = [TimeSpan]::FromSeconds([Math]::Ceiling($remaining.TotalSeconds))
    $countdownLabel.Text = $remainingWhole.ToString("hh\:mm\:ss")

    if ($remainingWhole.TotalSeconds -le 60) {
        $countdownLabel.ForeColor = $script:Palette.Warning
    }
    else {
        $countdownLabel.ForeColor = $script:Palette.Accent
    }
}

function Refresh-DeviceList {
    $previousId = Get-SelectedDeviceId

    try {
        $raw = Invoke-EngineCapture -ExtraArgs @("-ListDevicesJson")
        if ([string]::IsNullOrWhiteSpace($raw)) {
            $devices = @()
        }
        else {
            $parsed = ConvertFrom-Json -InputObject $raw
            $devices = @($parsed | Where-Object { $_.State -eq "device" })
        }
    }
    catch {
        $script:AvailableDevices = @()
        $deviceComboBox.Items.Clear()
        Set-UiRunningState -IsRunning $false
        Add-LogLine -Text ("[UI] Device refresh failed: {0}" -f $_.Exception.Message) -Color $script:Palette.Warning
        return
    }

    $script:AvailableDevices = $devices
    $deviceComboBox.Items.Clear()

    foreach ($device in $script:AvailableDevices) {
        $labelParts = New-Object System.Collections.Generic.List[string]
        $labelParts.Add($device.Id)

        if (-not [string]::IsNullOrWhiteSpace($device.Model)) {
            $labelParts.Add(($device.Model -replace "_", " "))
        }
        elseif (-not [string]::IsNullOrWhiteSpace($device.Device)) {
            $labelParts.Add($device.Device)
        }

        if (-not [string]::IsNullOrWhiteSpace($device.Product)) {
            $labelParts.Add($device.Product)
        }

        [void]$deviceComboBox.Items.Add(($labelParts -join " | "))
    }

    if ($script:AvailableDevices.Count -eq 0) {
        Add-LogLine -Text "[UI] No online Android devices detected." -Color $script:Palette.Warning
    }
    else {
        $selectedIndex = 0
        if (-not [string]::IsNullOrWhiteSpace($previousId)) {
            for ($index = 0; $index -lt $script:AvailableDevices.Count; $index++) {
                if ($script:AvailableDevices[$index].Id -eq $previousId) {
                    $selectedIndex = $index
                    break
                }
            }
        }

        $deviceComboBox.SelectedIndex = $selectedIndex
        Add-LogLine -Text ("[UI] Device list refreshed. {0} device(s) available." -f $script:AvailableDevices.Count) -Color $script:Palette.Muted
    }

    Set-UiRunningState -IsRunning $false
}

function Flush-ProcessLogs {
    $stdoutText = Read-NewText -Path $script:StdoutPath -Offset ([ref]$script:StdoutOffset)
    if (-not [string]::IsNullOrEmpty($stdoutText)) {
        Add-LogText -Text $stdoutText -Color $script:Palette.LogInk
    }

    $stderrText = Read-NewText -Path $script:StderrPath -Offset ([ref]$script:StderrOffset)
    if (-not [string]::IsNullOrEmpty($stderrText)) {
        Add-LogText -Text $stderrText -Color $script:Palette.Warning
    }
}

function Finish-Run {
    if (-not $script:CurrentProcess) {
        return
    }

    Flush-ProcessLogs
    try {
        $script:CurrentProcess.WaitForExit()
    }
    catch {
    }
    $exitCode = $script:CurrentProcess.ExitCode
    if ($null -eq $exitCode) {
        $exitCode = -1
    }
    $script:CurrentProcess.Dispose()
    $script:CurrentProcess = $null
    Set-UiRunningState -IsRunning $false
    Update-CountdownDisplay

    if ($exitCode -eq 0) {
        if ($script:CurrentRunIsCalibrate) {
            $statusLabel.Text = "Calibrated"
            $statusLabel.ForeColor = $script:Palette.Success
            Add-LogLine -Text ("[UI] 设备校准完成 at {0}。后续执行将自动使用此设备的坐标。" -f (Get-Date -Format "HH:mm:ss")) -Color $script:Palette.Success
        } elseif ($script:CurrentRunIsPreflight) {
            $statusLabel.Text = "Checked"
            $statusLabel.ForeColor = $script:Palette.Success
            Add-LogLine -Text ("[UI] Run check completed successfully at {0}." -f (Get-Date -Format "HH:mm:ss")) -Color $script:Palette.Success
        } else {
            $statusLabel.Text = "Done"
            $statusLabel.ForeColor = $script:Palette.Success
            Add-LogLine -Text ("[UI] Script finished successfully at {0}." -f (Get-Date -Format "HH:mm:ss")) -Color $script:Palette.Success
        }
    }
    else {
        $statusLabel.Text = "Failed"
        $statusLabel.ForeColor = $script:Palette.Warning
        Add-LogLine -Text ("[UI] Script failed with exit code {0}." -f $exitCode) -Color $script:Palette.Warning
    }

    Cleanup-RedirectFiles
    $script:CurrentRunIsPreflight = $false
    $script:CurrentRunIsCalibrate = $false
}

function Start-Run {
    param(
        [switch]$PreflightOnly,
        [switch]$Calibrate
    )

    if ($script:CurrentProcess -and -not $script:CurrentProcess.HasExited) {
        return
    }

    $deviceId = Get-SelectedDeviceId
    if ([string]::IsNullOrWhiteSpace($deviceId)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Please refresh and choose one connected Android device first.",
            "No Device Selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $targetLeBi    = [int]$targetInput.Value
    $targetTimeText = if ($Calibrate) { "00:00:00" } else { $timeInput.Text.Trim() }

    if (-not $Calibrate) {
        try {
            [void][TimeSpan]::ParseExact(
                $targetTimeText,
                "hh\:mm\:ss",
                [System.Globalization.CultureInfo]::InvariantCulture
            )
        }
        catch {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Time must use HH:mm:ss, for example 12:00:00.",
                "Invalid Time",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            $timeInput.Focus()
            return
        }
    }

    Cleanup-RedirectFiles
    $script:StdoutPath = Join-Path $env:TEMP ("qq_music_exchange_stdout_{0}.log" -f ([guid]::NewGuid().ToString("N")))
    $script:StderrPath = Join-Path $env:TEMP ("qq_music_exchange_stderr_{0}.log" -f ([guid]::NewGuid().ToString("N")))

    if ($Calibrate) {
        Add-LogLine -Text ("[UI] 校准请求: device={0}" -f $deviceId) -Color $script:Palette.Accent
    } else {
        Add-LogLine -Text ("[UI] Start request: device={0}, target={1}, executeAt={2}, preflightOnly={3}" -f $deviceId, $targetLeBi, $targetTimeText, $PreflightOnly.IsPresent) -Color $script:Palette.Accent
    }
    $script:CurrentRunIsPreflight = $PreflightOnly.IsPresent
    $script:CurrentRunIsCalibrate = $Calibrate.IsPresent

    $argumentList = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $script:EngineScriptPath,
        "-DeviceId", $deviceId,
        "-TargetLeBi", [string]$targetLeBi,
        "-TargetTimeText", $targetTimeText
    )

    if ($PreflightOnly) { $argumentList += "-PreflightOnly" }
    if ($Calibrate)     { $argumentList += "-Calibrate" }

    $script:CurrentProcess = Start-Process `
        -FilePath "powershell.exe" `
        -ArgumentList $argumentList `
        -WorkingDirectory $script:RootPath `
        -WindowStyle Hidden `
        -PassThru `
        -RedirectStandardOutput $script:StdoutPath `
        -RedirectStandardError $script:StderrPath

    $script:StdoutOffset = 0L
    $script:StderrOffset = 0L
    if ($Calibrate) {
        $statusLabel.Text = "Calibrating"
    } elseif ($PreflightOnly) {
        $statusLabel.Text = "Checking"
    } else {
        $statusLabel.Text = "Running"
    }
    $statusLabel.ForeColor = $script:Palette.Accent
    Set-UiRunningState -IsRunning $true
    $pollTimer.Start()
}

$pollTimer.Add_Tick({
    Update-CountdownDisplay
    Flush-ProcessLogs

    if ($script:CurrentProcess -and $script:CurrentProcess.HasExited) {
        Finish-Run
    }
})

$clearLogButton.Add_Click({
    $logBox.Clear()
    Add-LogLine -Text ("[UI] Log cleared at {0}." -f (Get-Date -Format "HH:mm:ss")) -Color $script:Palette.Muted
})

$startButton.Add_Click({
    Start-Run
})

$preflightButton.Add_Click({
    Start-Run -PreflightOnly
})

$calibrateButton.Add_Click({
    Start-Run -Calibrate
})

$refreshDevicesButton.Add_Click({
    Refresh-DeviceList
})

$timeInput.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $_.SuppressKeyPress = $true
        Start-Run
    }
})

$timeInput.Add_TextChanged({
    Update-CountdownDisplay
})

$form.Add_FormClosing({
    if ($script:CurrentProcess -and -not $script:CurrentProcess.HasExited) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "The script is still running. Close the panel and stop this run?",
            "Confirm Close",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
            $_.Cancel = $true
            return
        }

        try {
            $script:CurrentProcess.Kill()
        }
        catch {
        }
    }

    $pollTimer.Stop()
    Cleanup-RedirectFiles
})

Add-LogLine -Text "[UI] Ready. First stop on QQ Music '乐币' search results with the RechargeCard visible, then click Execute." -Color $script:Palette.Muted

if ($SmokeTest) {
    Write-Output ("UI ready: target={0}, time={1}" -f $targetInput.Value, $timeInput.Text)
    Cleanup-RedirectFiles
    return
}

Refresh-DeviceList
Update-CountdownDisplay
$pollTimer.Start()
[void]$form.ShowDialog()
