<#
  .SYNOPSIS
  This script is used to garner license information from a single tenant and output that information to a CSV file
#>

    param (
        [Parameter(Mandatory=$true)]
        [string]$LicenseFilePath, #Defines file path to place license mapping 

        [Parameter(Mandatory=$true)]
        [string]$OutputPath, #Defines output file path for Licenses csv

        [Parameter(Mandatory=$true)]
        [string]$AccessToken # Access Token to call Microsoft Graph API
    )

    #Define Licnese CDN 
    $url = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'

    #Export the CSV to the local path
    Invoke-WebRequest -Uri $url -OutFile $LicenseFilePath

    function Get-Licenses {
    param (
        [Parameter(Mandatory=$true)]
        [string]$AccessToken,

        [Parameter(Mandatory=$true)]
        [string]$LicenseFilePath
    )

    # Graph API endpoint
    $LicenseURL = "https://graph.microsoft.com/beta/directory/subscriptions"
    $subscribedSkusURL = "https://graph.microsoft.com/beta/subscribedSkus"

    # Prepare Headers with AccessToken
    $headers = @{
    Authorization = "Bearer $($AccessToken)"
    'Accept'      = 'application/json'
    }

    # Request Licenses
    $LicenseCounts = (Invoke-RestMethod -Method Get -Uri $subscribedSkusURL -Headers $headers).value
    $MainLicenseArray = (Invoke-RestMethod -Method Get -Uri $LicenseURL -Headers $headers).value | Where-Object -Property nextLifecycleDateTime -GT (Get-Date) | Select-Object *,
    @{Name = 'consumedUnits'; Expression = { ($LicenseCounts | Where-Object -Property skuid -EQ $_.skuId).consumedUnits } },
    @{Name = 'prepaidUnits'; Expression = { ($LicenseCounts | Where-Object -Property skuid -EQ $_.skuId).prepaidUnits } }


    $ProductMappingTable = Import-CSV $LicenseFilePath

    $LicenseTable = foreach ($license in $MainLicenseArray) {
        $ProductName = ($ProductMappingTable | Where-Object { $_.GUID -eq $license.skuId}).'Product_Display_Name' | Select-Object -Last 1 
        if(!$ProductName) { $ProductName = $license.skuPartNumber }

        #Get the Time until renewal
        # Convert the string to a datetime object
        $licenseExpiryDate = [datetime]::ParseExact($license.nextLifecycleDateTime, "yyyy-MM-ddTHH:mm:ssZ", $null)

        # Calculate the difference in days
        $DaysTillRenewal = ($licenseExpiryDate - (Get-Date)).Days

        [pscustomobject]@{
        ProductName      = [string]$ProductName
        UnusedLicenses   = [string]$license.prepaidUnits.enabled - $license.consumedUnits
        TotalLicenses    = [string]"$($license.TotalLicenses)"
        RenewalDate      = [string]$license.nextLifecycleDateTime
        DaysUntilRenewal = [string]$DaysTillRenewal
        isTrial          = [bool]$license.isTrial

        }
        }

# Define a CSS style for the table
$style = @"
<style>
    body { font-family: Arial, sans-serif; }
    table { border-collapse: collapse; width: 80%; margin: 50px auto; }
    th, td { border: 1px solid #cccccc; padding: 10px; }
    th { background-color: #e6f2ff; font-weight: bold; }
    tr:nth-child(even) { background-color: #f2f2f2; }
    td span.highlight { color: red; }
</style>
"@

# Convert LicenseTable to HTML
$htmlString = $LicenseTable | ConvertTo-Html -Property ProductName, UnusedLicenses, TotalLicenses, RenewalDate, DaysUntilRenewal, isTrial -PreContent "<h2>Licenses</h2>" -Head $style | 
Out-String

# Treat the HTML as XML
[xml]$htmlXml = $htmlString

# Iterate through each row of the table
foreach ($row in $htmlXml.html.body.table.tr) {
    # Ensure it's not the header row
    if ($row.td -and $row.td[1]) {
        $unusedLicensesValue = $row.td[1].'#text'
        if ($unusedLicensesValue -gt 1) {
            $row.td[1].InnerXml = "<span class='highlight'>$unusedLicensesValue</span>"
        }
    }
}

# Convert XML back to string
$html = $htmlXml.OuterXml

# Save the HTML content to a file
$htmlFilePath = "C:/TEMP/LicenseReport.html"
$html | Out-File $htmlFilePath -Encoding utf8

# Open the created HTML file in the default browser
Start-Process $htmlFilePath

}
