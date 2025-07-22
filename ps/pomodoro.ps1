Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Windows PowerShell version check
if ($PSVersionTable.PSVersion.Major -lt 3) {
    Write-Host "PowerShell 3.0 or higher is required" -ForegroundColor Red
    return
}

# XAML definition
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Pomodoro Timer" Height="500" Width="500"
        ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen">
    <Grid Background="#f0f0f0">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Title -->
        <TextBlock Grid.Row="0" Text="Pomodoro Timer" 
                   FontSize="24" FontWeight="Bold" HorizontalAlignment="Center" 
                   Margin="0,10,0,20" Foreground="#d32f2f"/>
        
        <!-- Main timer display -->
        <Border Grid.Row="1" Background="White" CornerRadius="10" 
                Margin="20" BorderBrush="#ddd" BorderThickness="2">
            <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                <TextBlock x:Name="StatusText" Text="Ready" 
                           FontSize="18" HorizontalAlignment="Center" 
                           Margin="0,0,0,10" Foreground="#666"/>
                <TextBlock x:Name="TimerDisplay" Text="25:00" 
                           FontSize="48" FontWeight="Bold" 
                           HorizontalAlignment="Center" 
                           Foreground="#d32f2f"/>
                <ProgressBar x:Name="ProgressBar" Height="8" Width="200" 
                             Margin="0,20,0,0" Background="#f0f0f0" 
                             Foreground="#d32f2f"/>
            </StackPanel>
        </Border>
        
        <!-- Time setting panel -->
        <Border Grid.Row="2" Background="White" CornerRadius="5" 
                Margin="20,0,20,10" Padding="15" BorderBrush="#ddd" BorderThickness="1">
            <StackPanel>
                <TextBlock Text="Time Settings" FontSize="14" FontWeight="Bold" 
                           Margin="0,0,0,10" Foreground="#333"/>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <!-- Work time -->
                    <StackPanel Grid.Column="0" Margin="0,0,10,0">
                        <TextBlock Text="Work (min)" FontSize="12" Margin="0,0,0,5"/>
                        <TextBox x:Name="WorkMinutes" Text="25" 
                                 HorizontalAlignment="Stretch" Height="25" 
                                 TextAlignment="Center"/>
                    </StackPanel>
                    
                    <!-- Short break -->
                    <StackPanel Grid.Column="1" Margin="5,0,5,0">
                        <TextBlock Text="Short Break (min)" FontSize="12" Margin="0,0,0,5"/>
                        <TextBox x:Name="ShortBreakMinutes" Text="5" 
                                 HorizontalAlignment="Stretch" Height="25" 
                                 TextAlignment="Center"/>
                    </StackPanel>
                    
                    <!-- Long break -->
                    <StackPanel Grid.Column="2" Margin="10,0,0,0">
                        <TextBlock Text="Long Break (min)" FontSize="12" Margin="0,0,0,5"/>
                        <TextBox x:Name="LongBreakMinutes" Text="15" 
                                 HorizontalAlignment="Stretch" Height="25" 
                                 TextAlignment="Center"/>
                    </StackPanel>
                </Grid>
            </StackPanel>
        </Border>
        
        <!-- Control buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" 
                    HorizontalAlignment="Center" Margin="20,0,20,20">
            <Button x:Name="StartButton" Content="Start" 
                    Width="80" Height="35" Margin="0,0,10,0" 
                    Background="#4caf50" Foreground="White" 
                    FontWeight="Bold" BorderThickness="0" 
                    Cursor="Hand"/>
            <Button x:Name="StopButton" Content="Stop" 
                    Width="80" Height="35" Margin="0,0,10,0" 
                    Background="#f44336" Foreground="White" 
                    FontWeight="Bold" BorderThickness="0" 
                    Cursor="Hand" IsEnabled="False"/>
            <Button x:Name="ResetButton" Content="Reset" 
                    Width="80" Height="35" Margin="0,0,10,0" 
                    Background="#ff9800" Foreground="White" 
                    FontWeight="Bold" BorderThickness="0" 
                    Cursor="Hand"/>
            <Button x:Name="SkipButton" Content="Skip" 
                    Width="80" Height="35" 
                    Background="#9c27b0" Foreground="White" 
                    FontWeight="Bold" BorderThickness="0" 
                    Cursor="Hand" IsEnabled="False"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load XAML (Windows PowerShell)
try {
    $reader = New-Object System.Xml.XmlNodeReader([xml]$xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Host "XAML loading error: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Get controls
try {
    $TimerDisplay = $window.FindName("TimerDisplay")
    $StatusText = $window.FindName("StatusText")
    $ProgressBar = $window.FindName("ProgressBar")
    $WorkMinutes = $window.FindName("WorkMinutes")
    $ShortBreakMinutes = $window.FindName("ShortBreakMinutes")
    $LongBreakMinutes = $window.FindName("LongBreakMinutes")
    $StartButton = $window.FindName("StartButton")
    $StopButton = $window.FindName("StopButton")
    $ResetButton = $window.FindName("ResetButton")
    $SkipButton = $window.FindName("SkipButton")
} catch {
    Write-Host "Control acquisition error: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Timer variables
$script:Timer = $null
$script:RemainingSeconds = 0
$script:TotalSeconds = 0
$script:IsRunning = $false
$script:CurrentSession = "work"  # work, shortbreak, longbreak
$script:SessionCount = 0

# Format time to min:sec
function Format-Time {
    param($seconds)
    $minutes = [Math]::Floor($seconds / 60)
    $remainingSeconds = $seconds % 60
    $minutesStr = $minutes.ToString("00")
    $secondsStr = $remainingSeconds.ToString("00")
    return "$minutesStr`:$secondsStr"
}

# Show notification (修正版)
function Show-Notification {
    param($title, $message)
    
    try {
        # Show dialog first
        [System.Windows.MessageBox]::Show($message, $title, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        
        # Show balloon notification (エラー処理を強化)
        $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
        if ($notifyIcon) {
            $notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
            $notifyIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
            $notifyIcon.BalloonTipTitle = $title
            $notifyIcon.BalloonTipText = $message
            $notifyIcon.Visible = $true
            $notifyIcon.ShowBalloonTip(5000)
            
            # タイマーを使用してアイコンを削除（修正版）
            $cleanupTimer = New-Object System.Windows.Forms.Timer
            $cleanupTimer.Interval = 6000
            
            # イベントハンドラーをより安全に実装
            $cleanupTimer.Add_Tick({
                try {
                    if ($notifyIcon) {
                        $notifyIcon.Visible = $false
                        $notifyIcon.Dispose()
                        $notifyIcon = $null
                    }
                    if ($cleanupTimer) {
                        $cleanupTimer.Stop()
                        $cleanupTimer.Dispose()
                        $cleanupTimer = $null
                    }
                } catch {
                    # エラーを無視（既に削除されている可能性があるため）
                }
            })
            
            $cleanupTimer.Start()
        }
    } catch {
        # 通知が失敗した場合はダイアログのみ表示
        try {
            [System.Windows.MessageBox]::Show($message, $title, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        } catch {
            Write-Host "$title`: $message" -ForegroundColor Green
        }
    }
}

# Start timer
function Start-Timer {
    param($minutes)
    
    $script:TotalSeconds = $minutes * 60
    $script:RemainingSeconds = $script:TotalSeconds
    $script:IsRunning = $true
    
    # Initialize progress bar
    $ProgressBar.Maximum = $script:TotalSeconds
    $ProgressBar.Value = 0
    
    # Create timer
    $script:Timer = New-Object System.Windows.Threading.DispatcherTimer
    $script:Timer.Interval = [TimeSpan]::FromSeconds(1)
    
    # Timer event
    $script:Timer.Add_Tick({
        try {
            $script:RemainingSeconds--
            
            # Update display
            $TimerDisplay.Text = Format-Time -seconds $script:RemainingSeconds
            $ProgressBar.Value = $script:TotalSeconds - $script:RemainingSeconds
            
            # Timer completion
            if ($script:RemainingSeconds -le 0) {
                $script:Timer.Stop()
                $script:IsRunning = $false
                
                # Update button states
                $StartButton.IsEnabled = $true
                $StopButton.IsEnabled = $false
                $SkipButton.IsEnabled = $false
                
                # Complete session
                Complete-Session
            }
        } catch {
            # タイマーエラー時の処理
            Write-Host "Timer error: $($_.Exception.Message)" -ForegroundColor Red
            if ($script:Timer) {
                $script:Timer.Stop()
            }
            $script:IsRunning = $false
            $StartButton.IsEnabled = $true
            $StopButton.IsEnabled = $false
            $SkipButton.IsEnabled = $false
        }
    })
    
    # Start timer
    $script:Timer.Start()
    
    # Update button states
    $StartButton.IsEnabled = $false
    $StopButton.IsEnabled = $true
    $SkipButton.IsEnabled = $true
}

# Complete session
function Complete-Session {
    switch ($script:CurrentSession) {
        "work" {
            $script:SessionCount++
            Show-Notification -title "Work Complete!" -message "Good job! Time for a break."
            
            # Long break after 4 work sessions
            if ($script:SessionCount % 4 -eq 0) {
                $script:CurrentSession = "longbreak"
                $StatusText.Text = "Long Break Time"
                $StatusText.Foreground = "#4caf50"
                $TimerDisplay.Foreground = "#4caf50"
                $ProgressBar.Foreground = "#4caf50"
            } else {
                $script:CurrentSession = "shortbreak"
                $StatusText.Text = "Short Break Time"
                $StatusText.Foreground = "#2196f3"
                $TimerDisplay.Foreground = "#2196f3"
                $ProgressBar.Foreground = "#2196f3"
            }
        }
        "shortbreak" {
            Show-Notification -title "Break Complete!" -message "Time to get back to work!"
            $script:CurrentSession = "work"
            $StatusText.Text = "Work Time"
            $StatusText.Foreground = "#d32f2f"
            $TimerDisplay.Foreground = "#d32f2f"
            $ProgressBar.Foreground = "#d32f2f"
        }
        "longbreak" {
            Show-Notification -title "Long Break Complete!" -message "Refreshed? Time for the next work session!"
            $script:CurrentSession = "work"
            $StatusText.Text = "Work Time"
            $StatusText.Foreground = "#d32f2f"
            $TimerDisplay.Foreground = "#d32f2f"
            $ProgressBar.Foreground = "#d32f2f"
        }
    }
    
    # Show next session time
    $nextMinutes = Get-NextSessionMinutes
    $nextSeconds = $nextMinutes * 60
    $TimerDisplay.Text = Format-Time -seconds $nextSeconds
    $ProgressBar.Value = 0
}

# Get next session minutes
function Get-NextSessionMinutes {
    switch ($script:CurrentSession) {
        "work" { return [int]$WorkMinutes.Text }
        "shortbreak" { return [int]$ShortBreakMinutes.Text }
        "longbreak" { return [int]$LongBreakMinutes.Text }
    }
}

# Reset timer
function Reset-Timer {
    if ($script:Timer) {
        $script:Timer.Stop()
        $script:Timer = $null
    }
    
    $script:IsRunning = $false
    $script:CurrentSession = "work"
    $script:SessionCount = 0
    
    # Reset display
    $StatusText.Text = "Ready"
    $StatusText.Foreground = "#666"
    $workSeconds = [int]$WorkMinutes.Text * 60
    $TimerDisplay.Text = Format-Time -seconds $workSeconds
    $TimerDisplay.Foreground = "#d32f2f"
    $ProgressBar.Foreground = "#d32f2f"
    $ProgressBar.Value = 0
    
    # Reset button states
    $StartButton.IsEnabled = $true
    $StopButton.IsEnabled = $false
    $SkipButton.IsEnabled = $false
}

# Event handlers
$StartButton.Add_Click({
    try {
        if (-not $script:IsRunning) {
            $minutes = Get-NextSessionMinutes
            Start-Timer -minutes $minutes
            
            # Update status
            switch ($script:CurrentSession) {
                "work" { $StatusText.Text = "Working..." }
                "shortbreak" { $StatusText.Text = "Short Break..." }
                "longbreak" { $StatusText.Text = "Long Break..." }
            }
        }
    } catch {
        Write-Host "Start button error: $($_.Exception.Message)" -ForegroundColor Red
    }
})

$StopButton.Add_Click({
    try {
        if ($script:IsRunning -and $script:Timer) {
            $script:Timer.Stop()
            $script:IsRunning = $false
            
            # Update button states
            $StartButton.IsEnabled = $true
            $StopButton.IsEnabled = $false
            $SkipButton.IsEnabled = $false
            
            # Update status
            $StatusText.Text = "Paused"
            $StatusText.Foreground = "#ff9800"
        }
    } catch {
        Write-Host "Stop button error: $($_.Exception.Message)" -ForegroundColor Red
    }
})

$ResetButton.Add_Click({
    try {
        Reset-Timer
    } catch {
        Write-Host "Reset button error: $($_.Exception.Message)" -ForegroundColor Red
    }
})

$SkipButton.Add_Click({
    try {
        if ($script:IsRunning) {
            $script:Timer.Stop()
            $script:IsRunning = $false
            
            # Update button states
            $StartButton.IsEnabled = $true
            $StopButton.IsEnabled = $false
            $SkipButton.IsEnabled = $false
            
            # Complete session
            Complete-Session
        }
    } catch {
        Write-Host "Skip button error: $($_.Exception.Message)" -ForegroundColor Red
    }
})

# TextBox input validation (エラー処理を追加)
$WorkMinutes.Add_TextChanged({
    try {
        $pattern = "^[0-9]+$"
        if ($WorkMinutes.Text -match $pattern -and [int]$WorkMinutes.Text -gt 0) {
            if (-not $script:IsRunning -and $script:CurrentSession -eq "work") {
                $seconds = [int]$WorkMinutes.Text * 60
                $TimerDisplay.Text = Format-Time -seconds $seconds
            }
        }
    } catch {
        # エラーを無視
    }
})

$ShortBreakMinutes.Add_TextChanged({
    try {
        $pattern = "^[0-9]+$"
        if ($ShortBreakMinutes.Text -match $pattern -and [int]$ShortBreakMinutes.Text -gt 0) {
            if (-not $script:IsRunning -and $script:CurrentSession -eq "shortbreak") {
                $seconds = [int]$ShortBreakMinutes.Text * 60
                $TimerDisplay.Text = Format-Time -seconds $seconds
            }
        }
    } catch {
        # エラーを無視
    }
})

$LongBreakMinutes.Add_TextChanged({
    try {
        $pattern = "^[0-9]+$"
        if ($LongBreakMinutes.Text -match $pattern -and [int]$LongBreakMinutes.Text -gt 0) {
            if (-not $script:IsRunning -and $script:CurrentSession -eq "longbreak") {
                $seconds = [int]$LongBreakMinutes.Text * 60
                $TimerDisplay.Text = Format-Time -seconds $seconds
            }
        }
    } catch {
        # エラーを無視
    }
})

# Window closing event
$window.Add_Closing({
    try {
        if ($script:Timer) {
            $script:Timer.Stop()
            $script:Timer = $null
        }
    } catch {
        # エラーを無視
    }
})

# Initial display
Reset-Timer

# Show window
try {
    $window.ShowDialog() | Out-Null
} catch {
    Write-Host "Window display error: $($_.Exception.Message)" -ForegroundColor Red
}
