<#      Script allows to continuously monitor disks free space.
#       In input parameters you can specify a time difference for checking the free space on disks.
#       It is mandatory to provide a GMail email and password for the account to send email notifications.
#       Also, you need to specify the receiver email address for notifications
#
#
#       Known bugs:
#           - Sometimeas BaloonTip won't appear but I have no idea why...
#           - Another problem with BaloonTip - BaloonTipClicked sometimes won't take action. Also, no idea why
#           - maybe it is a problem with my notification settings
#       Fortunately, GMail sending works well
#       PowerShell version: 7.1
#       Created by: Maciej Bąk, 400666 geoinf - 17.04.2021
#>


param (
    [int]$Threshold = 5
)
function Show-Notification {
    # shows BalloonTip and reacts for click event
    param(
        [string[]]$title,
        [string[]]$text,
        [int32]$duration,
        [scriptblock]$Action
    )
     # create .NET Forms NotifyIcon instance
    $notification = New-Object System.Windows.Forms.NotifyIcon
    $notification.Icon = [System.Drawing.SystemIcons]::Warning
    # BallonTip main settings
    $notification.BalloonTipTitle = $title
    $notification.BalloonTipText = $text
    # make it visible and show it
    $notification.Visible = $true
    $notification.ShowBalloonTip($duration)
    # unregister previous event form BalloonTip named "balloonClicked"
    Unregister-Event -SourceIdentifier balloonClicked -ErrorAction SilentlyContinue
    # register new event - reaction to click BalloonTip.
    # action passed to function will be executed
    Register-ObjectEvent -InputObject $notification -EventName "BalloonTipClicked" -SourceIdentifier balloonClicked -Action $Action | Out-Null
    # wait some time for Evebt
    Wait-Event -timeout 15 -sourceIdentifier balloonClicked > $null
}

function Send-GMailMessage {
    # sends mail from gmial
    param (
        [string]$mailFrom,
        [SecureString]$password,
        [string]$mailTo,
        [string]$topic,
        [string]$content
    )

    # set authentication data
    $authentication = New-Object System.Net.NetworkCredential($mailFrom, $password)

    # set SMTP
    $smtpClient = New-Object Net.Mail.SmtpClient("smtp.gmail.com", 587)
    $smtpClient.Credentials = $authentication
    $smtpClient.EnableSsl = $true
    $smtpClient.Timeout = 10000 # time for timeout exception
    $smtpClient.UseDefaultCredentials = $false

    # set MailMessage
    $mailMessage = New-Object Net.Mail.MailMessage
    $mailMessage.From = $mailFrom
    $mailMessage.To.Add($mailTo)
    $mailMessage.Subject = $topic
    $mailMessage.Body = $content
    $mailMessage.IsBodyHtml = $true
    # send it!
    try {
        $smtpClient.Send($mailMessage)
    }
    catch [System.Management.Automation.MethodInvocationException] {
        $mess = "An error occured - the operation has timed out. Check your internet connection and try again."
        Write-Error $mess
        Show-Notification -title "Message from Disks Monitor." -text $mess -duration 5000 -Action $Action
        Start-Sleep -s 30
        exit
    }

    Write-Host "Mail Message was sent."
}

function Get-MailInfo {
    # get login and pswd for GMail
    [hashtable]$mailInfo = @{}

    Write-Host "You need to provide some parameters to run script correctly.`r`n" -
    Write-Host "Email address to recieve notifications: "
    [string]$reciever = Read-Host
    $mailInfo.reciever = $reciever
    Write-Host "Gmail email address to send notifications: "
    [string]$sender = Read-Host
    $mailInfo.sender = $sender
    Write-Host "GMail password: "
    [SecureString]$senderPsw = Read-Host -AsSecureString
    $mailInfo.senderPsw = $senderPsw

    return $mailInfo
}

function Convert-ToMB {
    # convert value to MB
    param (
        [System.Array]$inputArray
    )
    $inputArray = $inputArray | ForEach-Object { $_ / 1MB }
    $inputArray = $inputArray | ForEach-Object {[math]::Round($_,2)}
    return $inputArray
}

function Build-Message {
    # build message to email
    param (
        [System.Array]$Id,
        [System.Array]$Free,
        [System.Array]$Used,
        [System.Array]$Perc
    )

    $head = "
        <head><style>
        .mess {
        border: 5px outset #538d22;
        background-color: #aad576;
        text-align: left;
        color: #143601;
        padding: 12px;
        font-family: 'Roboto', sans-serif;
        font-size: medium;
        }
        ol,li {text-align: left;}</style></head>
    "
    $div = '
        <div class = "mess">
        <h1 style="text-align: center;">Report from MB computer</h1>
        <p>Your disks are running out of space: <ul>
    ' -f  $env:computername

    if ($Id.Count -eq 1) {
        $UL = '
        <li> On <b>{1}</b> drive there is <b>{2} MB</b> of free space which is <b>{3}%</b> of drive capacity. </li> </ul></p>
        ' -f $env:computername, $Id[0], $Free[0], $Perc[0]
        $endUL = '</ul></p></div>'
        # concat strings
        $message = $head + $div + $UL + $endUL
    }

    if ($Id.Count -gt 1) {
        $OL = ' '
        for ($i = 0; $i -lt $Id.Count; $i++) {
            $OL_new= '
                <li> On <b>{0}</b> drive there is <b>{1} MB</b> of free space which is <b>{2}%</b> of drive capacity. </li>
            ' -f $Id[$i], $Free[$i], $Perc[$i]
            $OL +=$OL_new
        }
        $endOL = '</ol></p></div>'
        # concat strings
        $message = $head + $div + $OL + $endOL
    }

    return $message
}

function Build-Title {
    # buiild title for email
    param (
        [System.Array]$Id
    )
    if ($Id.Count -eq 1) {
        $title = "Drive $($Id) is running out of space!"
    }

    if ($Id.Count -gt 1) {
        [string]$string = $Id
        $string = $string -replace (" ", ", ")
        $title = "Drives: $($string) are running out of space!"
    }
    return $title
}

# import .NET Forms class to create NotifyIcon instance
try {
    New-Object System.Windows.Forms.NotifyIcon
}
catch [System.Management.Automation.PSArgumentException]{
    Write-Host "Assembly will be added."
    Add-Type -AssemblyName System.Windows.Forms -PassThru
}

# get emails and password
$mailInfo = Get-MailInfo

# script main loop
while ($true) {
    Write-Host "Disk free space checking procedure was started..." -ForegroundColor Green
    # get Disks info
    $disks = Get-CimInstance -Class Win32_LogicalDisk

    # arrays for disks info
    $disksFree = @()
    $disksUsed = @()
    $disksPerc = @()
    # first - choose which disks are running of space and save names in an array
    $disksId = @()
    foreach ($disk in $disks) {
        if ($disk.FreeSpace -lt ($disk.Size) * 0.1) { # było 0.1
            $disksId += $disk.deviceId
            $disksFree += $disk.FreeSpace
            $disksUsed += ($disk.Size - $disk.FreeSpace)
            $disksPerc += [math]::Round((($disk.FreeSpace / ($disk.Size - $disk.FreeSpace)) * 100),2) #free space in %
        }
    }

    # only if there is a problem with one (or more) of drives
    if ($disksId.Count -ne 0) {
        Write-Host "Something happened!" -ForegroundColor Red
        # units conversion
        $disksFree = Convert-ToMB $disksFree
        $disksUsed = Convert-ToMB $disksUsed

        # build mail message
        $message = Build-Message -Id $disksId -Free $disksFree -Used $disksUsed -Perc $disksPerc
        #send mail message
        Send-GMailMessage -mailFrom $mailInfo.sender -password $mailInfo.senderPsw -mailTo $mailInfo.reciever -topic "Sometnig happened on $($env:computername) computer!" -content $message
        # parameters for BaloonTip
        $text = "The report was sent to Your email. Click to open This PC."
        # action for BaloonTipClicked event
        $Action = {
            Write-Host "BalloonTipClicked event occured" -ForegroundColor Grey
            Start-Process explorer.exe file:
        }
        $title = Build-Title -Id $disksId
        Write-Host "$($title)" -ForegroundColor Red
        Show-Notification -title $title -text $text -duration 5000 -Action $Action
    }

    # clear disk alerts array
    $disksId, $disksFree, $disksUsed, $disksPerc = $null

    Start-Sleep -s $Threshold
}
