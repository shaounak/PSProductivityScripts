# . $PSScriptRoot\..\Invoke-Profile.ps1
Add-Type -AssemblyName PresentationFramework
start spotify

Function Notify-Host {
    [CmdletBinding()]
    Param (
        #Duration of your Pomodoro Session
        [string]$NotificationSound = "C:\Windows\Media\Windows Proximity Connection.wav",
        [Int]$times = 2
    )
    #Playing end of focus session notification
    $player = New-Object System.Media.SoundPlayer $NotificationSound -ErrorAction SilentlyContinue
        1..$times | ForEach-Object {
            $player.Play()
        Start-Sleep -m 1400 
    }
}

function Start-Pomo {
    [CmdletBinding()]
    param (
        [Int]$Time = $PomodoroSessionDuration
    )
    Notify-Host -times 1
    #Minimize-Window -Process $Process
    Write-Host "Pomodoro Session # $SessionCounter : Started"
    Start-CountDownTimer -Minutes $Time
    Notify-Host -NotificationSound "C:\Windows\Media\Windows Proximity Notification.wav"
    Write-Host "Pomodoro Session # $SessionCounter : Ended"
	Minimize-Window -Process $Process
    Restore-Window -Process $Process
    "=" * 70
}

$SessionCounter = 1 # Session Counter
$ContinuePomodoroIteration = $true # Continue Pomodoro Iteration
$PomodoroSessionDuration = 25 # Minutes
$ShortBreakDuration = 5 # Minutes
$LongBreakDuration = 15 # Minutes
$Process =  $(Get-Process | Where-Object -Property ID -eq $PID)

while ($ContinuePomodoroIteration -eq $true) {
    $option = [System.Windows.MessageBox]::Show('Start the Pomodoro Focus Session#: ' + $SessionCounter + '?', 'Start Session', 'YesNoCancel', 'Information')
    if ( $option -eq "Yes") {
        Start-Pomo -Time $PomodoroSessionDuration
    }
    
    "=" * 70

    if ( $SessionCounter -gt 3 ) {    
        $SessionCounter = 1
        $option = [System.Windows.MessageBox]::Show('Pomodoro Session Finished, Time for a long break?', 'Time for a long break', 'YesNoCancel', 'Information')
        if ( $option -eq "Yes") {
            Write-Host 'Pomodoro Session Finished, Time for a long break'
            "-" * 70
            
            Start-CountDownTimer -Minutes $LongBreakDuration
            $SessionCounter = 1
            Write-Host 'Break finished, Time for a next pomodoro'
            "=" * 70
        }
        else {
            Write-Host 'Skipping Long Break'
        }
    }
    else {
        $option = [System.Windows.MessageBox]::Show('Pomodoro Session Finished, Time for a Short break?', 'Time for a short break', 'YesNoCancel', 'Information')
        if ( $option -eq "Yes") {
            Write-Host 'Pomodoro Session Finished, Time for a Short break'
            "-" * 70
            Start-CountDownTimer -Minutes $ShortBreakDuration
            Write-Host 'Break finished, Time for a next pomodoro'
            "=" * 70
        }
        else {
            Write-Host 'Skipping Short Break'
        }
        
    }
    
    if ( $option -eq "Cancel") {
        $ContinuePomodoroIteration = $false
    }
    $SessionCounter += 1
}
