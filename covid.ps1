<#      Script allows to get data from covid19api.com
# 
#       Created by: Maciej Bąk, 400666 geoinf - 17.04.2021
#       PowerShell version: 7.1
#>

$url = 'https://api.covid19api.com/summary'

$headers = @{}
$response = Invoke-RestMethod $url -Method 'GET' -Headers $headers

$currentDate = Get-Date

$result = $response | Select-Object -expand Countries | Sort-Object -Property TotalConfirmed -Descending | Select-Object Country, TotalConfirmed, TotalRecovered, TotalDeaths, @{n='LastUpdate'; e={
    $dateDiff = New-TimeSpan -Start $response.Date -End $currentDate

    if ($dateDiff.Days -lt 1) {
        if ($dateDiff.Hours -lt 1) {
            return $dateDiff.Minutes
        } else {
            return "$($dateDiff.Hours) h $($dateDiff.Minutes) min "
        } 
    } else {
        if ($dateDiff.Days -eq 1) {
            return  "$($dateDiff.Days) day $($dateDiff.Hours) h $($dateDiff.Minutes) min"
        } else {
            return  "$($dateDiff.Days) days $($dateDiff.Hours) h $($dateDiff.Minutes) min"
        }
    }

}} | Format-Table

$result


