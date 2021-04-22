<#      Script allows to continuously monitor Kindle price on amazon.pl.
#       In input parameters you can specify a time difference for checking the price.
#       It is mandatory to provide GMail email and password for the account to send email notifications.
#       Also, you need to specify the receiver email address for notifications
#
#       PowerShell version: 7.1
#       Created by: Maciej Bąk, 400666 geoinf - 17.04.2021
#>

param (
    [int]$Threshold = 5
)

function Send-GMailMessage {
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
    $smtpClient.Timeout = 8000 # time for timeout exception
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
        $mess = "An error occured - the operation has timed out. Check your internet connection and make sure that you provide correct password. Please, re-run the script."
        Write-Host "$($mess)" -ForegroundColor Red
        exit
    }
    Write-Host "Mail Message was sent." -ForegroundColor Blue
}
function Get-KindlePrice {
    param (
        [string]$url
    )
    # create WebRequest
    try{
         $website = Invoke-WebRequest -Uri $url -UseBasicParsing
    } catch { # catch Invike-WebRequest exceptions.

        $Status = $_.Exception.Message
        Write-Host $Status -ForegroundColor Red
        exit
    }
    # create HTML object to save HTML from WebRequest
    $htmlObject = New-Object -ComObject "HTMLFile"
    # get HTML content od website and cast to string
    [string]$body = $website.Content
    # pass content to previously created HTML file
    $htmlObject.write([ref]$body)
    # get a class with a price
    $filter = $htmlObject.getElementsByClassName("a-size-medium a-color-price priceBlockBuyingPriceString")
    # get the price - need to use Subexpression operator $(). $filter.textContent returns $null
    $rawprice = $($filter).textContent
    # get only numbers
    $priceString = $rawprice -replace "[^0-9]", ''
    $priceDouble = [int]$priceString / 100
    # throw if the price value is strange
    if (($null -eq $priceDouble) -or ($priceDouble -eq 0)) { # if the website won't have 'a-size-medium...' class $priceDouble will be null
        throw [System.Management.Automation.GetValueException]::new("Can't find the price. Kindle may be unavailable, or the amazon website could be different now.")
    }

    return $priceDouble
}

function Get-MailInfo {
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

# get emails and password
$mailInfo = Get-MailInfo


# script main loop
while ($true) {

    $url = "https://www.amazon.pl/dp/B07741S7Y8/ref=gw_pl_desk_mso_sc_eink_moo_launch?pf_rd_r=VXVWM4ZT3QZBN2XJGGPG&pf_rd_p=47b0de49-e429-4d75-852c-841aa314dd19"

    try {
        $price1 = Get-KindlePrice -url $url
    }
    catch {
        Write-Host "Error - $($_.Exception.Message)" -ForegroundColor Red
        exit
    }

    Write-Host "The Kindle price is $($price1) zł."
    Start-Sleep -s $Threshold

    # for test - comment try-catch block and uncomment $price2
    try {
        $price2 = Get-KindlePrice -url $url
    }
    catch {
        Write-Host "Error - $($_.Exception.Message)" -ForegroundColor Red
        exit
    }

    # $price2 = 100
    Write-Host "The Kindle price is $($price2) zł."

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
        <div class="mess">
        <h1>Amazon Kindle Price is now lower!</h1>
        <p>Previous price was {0} zł. Now, it is {1} zł.</p>
        <p> Click to go to <a href="{2}">amazon.pl</a></p></div>
    ' -f $price1, $price2, $url

    # concat message
    $message = $head + $div
    # 2 lower than 1? send email!
    if ($price2 -lt $price1) {
        Write-Host "The price is lower now! Previus: $($price1) zł Now: $($price2) zł!" -ForegroundColor Green
        Send-GMailMessage -mailFrom $mailInfo.sender -password $mailInfo.senderPsw -mailTo $mailInfo.reciever -topic "Kindle Price Update!" -content $message
    }
}
