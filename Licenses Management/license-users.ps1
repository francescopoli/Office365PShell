<#
    ###############Disclaimer#####################################################
    The sample scripts are not supported under any Microsoft standard support 
    program or service. The sample scripts are provided AS IS without warranty  
    of any kind. Microsoft further disclaims all implied warranties including,  
    without limitation, any implied warranties of merchantability or of fitness for 
    a particular purpose. The entire risk arising out of the use or performance of  
    the sample scripts and documentation remains with you. In no event shall 
    Microsoft, its authors, or anyone else involved in the creation, production, or 
    delivery of the scripts be liable for any damages whatsoever (including, 
    without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use 
    of or inability to use the sample scripts or documentation, even if Microsoft 
    has been advised of the possibility of such damages.
    ###############Disclaimer#####################################################
#>

Function License-Users {
[CmdletBinding()]

Param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$false)] $CSVFile,
    [Parameter(Mandatory=$false,ValueFromPipeline=$false)] $LogDirectory,

    [parameter(Mandatory=$false)] # not available in CU,IR,KP,SD,SY
    [ValidateSet("AF","AX","AL","DZ","AS","AD","AO","AI","AQ","AG","AR","AM","AW","AU","AT","AZ","BS","BH","BD",
                 "BB","BY","BE","BZ","BJ","BM","BT","BO","BQ","BA","BW","BV","BR","IO","BN","BG","BF","BI","CV",
                 "KH","CM","CA","KY","CF","TD","CL","CN","CX","CC","CO","KM","CD","CG","CK","CR","CI","HR",
                 "CW","CY","CZ","DK","DJ","DM","DO","EC","EG","SV","GQ","ER","EE","ET","FK","FO","FJ","FI","FR",
                 "GF","PF","TF","GA","GM","GE","DE","GH","GI","GR","GL","GD","GP","GU","GT","GG","GN","GW","GY",
                 "HT","HM","VA","HN","HK","HU","IS","IN","ID","IQ","IE","IM","IL","IT","JM","JP","JE","JO",
                 "KZ","KE","KI","KR","KW","KG","LA","LV","LB","LS","LR","LY","LI","LT","LU","MO","MK","MG",
                 "MW","MY","MV","ML","MT","MH","MQ","MR","MU","YT","MX","FM","MD","MC","MN","ME","MS","MA","MZ",
                 "MM","NA","NR","NP","NL","NC","NZ","NI","NE","NG","NU","NF","MP","NO","OM","PK","PW","PS","PA",
                 "PG","PY","PE","PH","PN","PL","PT","PR","QA","RE","RO","RU","RW","BL","SH","KN","LC","MF","PM",
                 "VC","WS","SM","ST","SA","SN","RS","SC","SL","SG","SX","SK","SI","SB","SO","ZA","GS","SS","ES",
                 "LK","SR","SJ","SZ","SE","CH","TW","TJ","TZ","TH","TL","TG","TK","TO","TT","TN","TR",
                 "TM","TC","TV","UG","UA","AE","GB","UM","US","UY","UZ","VU","VE","VN","VG","VI","WF","EH","YE","ZM","ZW")] [string] $UsageLocation
    )


    Begin {

        If (!(Get-Command Get-AzureADUser) ) {
            If ( !(Get-Module -ListAvailable AzureAD) )
            { 
                Write-Host "[WARN] Azure AD Module missing, trying to install..."
                Write-Host "in case of failure, open a new PowerShell in"
                Write-Host " Administrator mode, and execute: "
                Write-Host "Install-Module -Name AzureAD -Confirm:$false -Force -AllowClobber"

                Install-Module -Name AzureAD -Confirm:$false -Force -AllowClobber
            }
        }
        $Error.Clear()
        
        try {
            $testSKU = (Get-AzureADSubscribedSku -ErrorAction SilentlyContinue)
            }
        catch{
            # not needed, will check just after the block, 
            # this only to prevent error to surface in the console
        }
    
        If ($Error.Count -ge 1) {

            $Error.Clear()
            Write-Host "Connessione ad AzureAD Powershell..."
        
            Import-Module AzureAD -Verbose:$false
            Connect-AzureAD -LogLevel None | Out-Null

            $testSKU = (Get-AzureADSubscribedSku -ErrorAction SilentlyContinue)   

            If ($Error.Count -ge 1) {
                Write-Verbose "[ERR] Unable to connect to AzureAD PowerShell. Exiting..."
                Throw ("[ERR] Unable to connect to AzureAD PowerShell. Exiting...") 
            }
        }
        Write-Verbose "AzureAD connection Executed"
        Write-Verbose "                                 "
  
    }

    Process{
        $validCountryCodes = @("AF","AX","AL","DZ","AS","AD","AO","AI","AQ","AG","AR","AM","AW","AU","AT","AZ","BS","BH","BD", `
                 "BB","BY","BE","BZ","BJ","BM","BT","BO","BQ","BA","BW","BV","BR","IO","BN","BG","BF","BI","CV",`
                 "KH","CM","CA","KY","CF","TD","CL","CN","CX","CC","CO","KM","CD","CG","CK","CR","CI","HR",`
                 "CW","CY","CZ","DK","DJ","DM","DO","EC","EG","SV","GQ","ER","EE","ET","FK","FO","FJ","FI","FR",`
                 "GF","PF","TF","GA","GM","GE","DE","GH","GI","GR","GL","GD","GP","GU","GT","GG","GN","GW","GY",`
                 "HT","HM","VA","HN","HK","HU","IS","IN","ID","IQ","IE","IM","IL","IT","JM","JP","JE","JO",`
                 "KZ","KE","KI","KR","KW","KG","LA","LV","LB","LS","LR","LY","LI","LT","LU","MO","MK","MG",`
                 "MW","MY","MV","ML","MT","MH","MQ","MR","MU","YT","MX","FM","MD","MC","MN","ME","MS","MA","MZ",`
                 "MM","NA","NR","NP","NL","NC","NZ","NI","NE","NG","NU","NF","MP","NO","OM","PK","PW","PS","PA",`
                 "PG","PY","PE","PH","PN","PL","PT","PR","QA","RE","RO","RU","RW","BL","SH","KN","LC","MF","PM",`
                 "VC","WS","SM","ST","SA","SN","RS","SC","SL","SG","SX","SK","SI","SB","SO","ZA","GS","SS","ES",`
                 "LK","SR","SJ","SZ","SE","CH","TW","TJ","TZ","TH","TL","TG","TK","TO","TT","TN","TR",`
                 "TM","TC","TV","UG","UA","AE","GB","UM","US","UY","UZ","VU","VE","VN","VG","VI","WF","EH","YE","ZM","ZW")

        Write-Host " "
        Write-Host "------ START LICENSES ASSIGNMENT ------"
        Write-Host " "
        $CountryCode = $null #Country code for the UsageLocation (mandatory)
        if ($UsageLocation) {$DefaultCountryCode = $UsageLocation}
        else {$DefaultCountryCode = $null}
        
        $PreviousTemplate = "" # keep track of previous template user
        $logs = ""

        # checking if logging has been required, if true, check the folder
        # and generate the file
        if ($LogDirectory){
            if (!(Test-Path $LogDirectory) ){
                New-Item -ItemType Directory -Path $LogDirectory | Out-Null
            }

            if (!(Test-Path $LogDirectory) ){
                Write-Host "Unable to access logs folder, Using only video logging"
            }
            else{
                $LogDirectory = Get-Item $LogDirectory
                $date = Get-Date
                $LogName = "$($date.Year)-$($date.Month)-$($date.Day)_$($date.Hour)-$($date.Minute)_Licenses.Log"
                if (!(Test-Path -Path ($LogDirectory.FullName + "\" + $LogName) )){
                    New-Item -ItemType File -Path $LogDirectory -Name $LogName | Out-Null
                }
                $logs = ($LogDirectory.FullName + "\" + $LogName)
            }
        }



        if ( !(Test-Path $CSVFile) ){
            Write-Host "Unable to open CSV file $($CSVFile) " -ForegroundColor Red
            if ($logs) {
                "Unable to open CSV file $($CSVFile)" | 
                    out-file -Append -FilePath $logs 
            }
        }

        $UsersList = Import-Csv -Path $CSVFile

        foreach ($user in $UsersList) {
            $Template = $user.TemplateEmail
            $TargetUser = $user.EmailAddress
            $CountryCode = $user.UsageLocation

            $License =  New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
            $Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses

            # Check if template user is the same, else check new current licenses
            if ( $PreviousTemplate -ne $Template ){ 
                $TemplateADUser = Get-AzureADUser -ObjectId $Template
                $PreviousTemplate = $Template
            }

            Write-Verbose "Processing User $($TargetUser) assigning licenses as per template $($Template)"
            
            #Check user location
            $utenteAD = Get-AzureADUser -ObjectId $TargetUser
            if ( $utenteAD.UsageLocation -eq $null ) {
                if ($CountryCode -eq $null) {
                    if ($DefaultCountryCode -eq $null){
                        Write-Host "UsageLocation must be set on the user to be able to assing a license"
                        $reAsk = $true
                        while ($reAsk) {
                            $CountryCode = Read-Host "Please enter the 2 letters country code for the user"
                            if ($validCountryCodes -notcontains $countryCode){
                                Write-Host "This is not a valid country code, please re-enter"
                            }
                            else{
                                $reAsk = $false
                                Write-Host  "Use this as default for all users without code in the csv? [" -NoNewline 
                                Write-Host "Y" -NoNewline -ForegroundColor Yellow
                                Write-Host "]\[n]" -NoNewline
                                $def = Read-Host " "
                                if ( ($def -eq "") -or ($def -eq "Y") ) {
                                    $DefaultCountryCode = $countryCode
                                }
                            }
                        }
                    } 
                    else {
                        $CountryCode = $DefaultCountryCode
                    }                 
                }
                Set-AzureADUser -ObjectId $TargetUser -UsageLocation $countryCode
                Write-Verbose "====> Assigned UsageLocation: $($countryCode) to user" 
                if ($logs) {
                    "====> Assigned UsageLocation: $($countryCode) to user" | 
                    out-file -Append -FilePath $logs 
                }

            }
            
            $count = 0
            $Error.Clear()
            # assign all the template user SKUS to the user
            if ( $TemplateADUser.AssignedLicenses.SkuId -ne $null){
                foreach ($sku in $TemplateADUser.AssignedLicenses.SkuId){
                    $License.SkuId = $sku
                    $Licenses.AddLicenses = $License
                    Set-AzureADUserLicense -ObjectId $TargetUser -AssignedLicenses $Licenses -ErrorAction SilentlyContinue
                    $count = $count + 1
                }
            }

            if ($Error.Count -ge 1) {
                Write-Host "There were some errors in assignign licenses to user " -ForegroundColor Red -NoNewline
                Write-Host "$($TargetUser) " -ForegroundColor Yellow -NoNewline
                Write-Host "please check licenses from the Office365 portal" -ForegroundColor Red
                if ($logs) {
                   "There were some errors in assignign licenses to user  $($TargetUser) please check licenses from the Office365 portal" | 
                    out-file -Append -FilePath $logs 
                }
            }
            else{
                if ($count -eq 0){
                    Write-Verbose "$($TargetUser) !!!!!!!!  $($count) Licenses assigned !!!!!!!"
                    Write-Host "Error:: "-ForegroundColor Red -NoNewline
                    Write-Host "$($TargetUser)" -ForegroundColor Yellow -NoNewline
                    Write-Host " not processed -> check template user licenses: " -NoNewline
                    Write-Host "$($Template)" -ForegroundColor Yellow
                    if ($logs) {
                       "Error :: $($TargetUser) not processed -> check template user licenses: $($Template)" | 
                        out-file -Append -FilePath $logs 
                     }
                }
                else {
                    if ($count -gt 1) {
                        Write-Host "$($TargetUser)" -ForegroundColor Cyan -NoNewline
                        write-host " ===> OK " -NoNewline
                        Write-Host "$($count) Licenses" -ForegroundColor Yellow -NoNewline
                        Write-Host " assigned"
                        if ($logs) {
                           "$($TargetUser) ===> OK $($count) Licenses assigned as per template $($Template)" | 
                            out-file -Append -FilePath $logs 
                        }
                    }
                    else{
                        Write-Host "$($TargetUser)" -ForegroundColor Cyan -NoNewline
                        write-host " ===> OK " -NoNewline
                        Write-Host "$($count) License" -ForegroundColor Yellow -NoNewline
                        Write-Host " assigned" 
                        if ($logs) {
                           "$($TargetUser) ===> OK $($count) License assigned as per template $($Template)" | 
                            out-file -Append -FilePath $logs 
                        }   
                    }
                }

            }

        }
        Write-Host " "
        Write-Host "------ END LICENSES ASSIGNMENT ------"
        if ($logs) {
            Write-Host " "
            Write-Host "Log file:"
            Write-Host "    $($logs) "
            Write-Host "available for review"
            Write-Host " "
        }
    }
}