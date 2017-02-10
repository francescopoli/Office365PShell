<#       
    .SYNOPSIS
    License-EMS is to facilitate licenses assingment process for Office 365 EMS SKU.

    .DESCRIPTION
    Provide instrumentation to:
        Add full EMS licenses to users.
        Add full EMS licenses to users with only some workloads enabled.
        Disable specific workloads
        Enable specific workloads
        Remove EMS license from users

    All that is not explicitly disabled, will be enabled. Always pass the full list of plans to be disabled.

    Accept a collection of users from pipeline or by a user entry.
    It support CSV file input with 2 fields, named: 
        1st: UserPrincipalName -OR- EmailAddress -OR- WindowsEmailAddress -OR- Users -OR- User 
        2nd: usageLocation [must respect iso codes https://www.iso.org/obp/ui/#search/code/]
            Sample CSV headers: 
                UserPrincipalName,usageLocation
                EmailAddress,usageLocation
                WindowsEmailAddress,usageLocation
                Users,usageLocation
                User,usageLocation
        Altough usageLocation is not mandatory and can be omitted in csv. UsageLocation is used(and needed) ONLY in case user is not
        having any prior license assigned from, also, any other Office 365 workloads. 
        If no usageLocation is specified and user doesn`t have one, then a prompt is returned.

    Input from pipeline can also come from Get-MsolUser,Get-MsolGroupMember and Get-Mailbox
	
	Author: Francesco Poli fpoli@microsoft.com
	
    .PARAMETER Users
     Single user object in valid smtp address format user@contoso.com or collection of objects from CSV file or pipeline input.
     Pipelined input must present at least one attribute named UserPrincipalName -OR- EmailAddress -OR- WindowsEmailAddress -OR- Users -OR- User 

	.PARAMETER usageLocation
    Valid 2 letter standard ISO code [https://www.iso.org/obp/ui/#search/code/] representing markets where Office365 is currently available
    Parameter is used(applied) when user has no prior UsageLocation already assinged. If UsageLocation is already assigned to user, parameter
    will not be used or applied to user, even so it will trigger an error if an invalid code is used.

    .PARAMETER RemoveEMSLicense
    Completely remove the EMS license from the user, if this parameter is used, is the same as passing all the -Disabled* ones

    .PARAMETER DisableRMS
    Disable plan RMS_S_PREMIUM -> Azure Rights Management
    If this parameter is used, it will trigger the DisableAzureIRM too, reason is that Azure Information Protection depends on AT LEAST 
    an assigned Azure Rights Management plan, even if in any other SKU (like the Enterprise). Here assuming that if RMS is to be disabled, there
    is not wish to use it from an another plan, so will be unlikely the Azure Infomrmation Protection will be used too.
    Please note that the RMSBASIC (-> Right Management Adhoc) in Office365 portal is not enough for the RMS_S_ENTERPRISE plan to be enabled.

    .PARAMETER DisableAzureIRM
    Disable plan RMS_S_ENTERPRISE - Azure Information Protection Plan 1

    .PARAMETER DisableIntune
    Disable plan INTUNE_A -> Intune A Direct 

    .PARAMETER DisableAADPremium
    Disable plan AAD_PREMIUM -> Azure Active Directory Premium Plan 1 

    .PARAMETER DisableMultifactor
    Disable plan MFA_PREMIUM -> Azure Multi-Factor Authentication 

    .PARAMETER Verbose
    Verbose output for console
        
    .EXAMPLE
    Example -1- License a single and assing Italy as user location
    License-EMSUser user@contoso.com -DisableAzureIRM -DisableRMS -DisableMultiFactor -usageLocation IT   
    
    .EXAMPLE
    Example -2- License a user with all available plans
    License-EMSUser user@contoso.com

    .EXAMPLE
    Example -3- License all users in a groups from Azure AD and keep the Multifactor Authentication disabled
    Get-MsolGroupMember -GroupObjectId 614162e2-67dd-4b33-875d-c486892a0ada -MaxResults unlimited | License-EMSUser -DisableMultiFactor

    .EXAMPLE
    Example -4- License all users in a groups from Exchange Online and keep the Multifactor Authentication disabled
    Get-DistributionGroupMember -Identity intunegroup -ResultSize unlimited | License-EMSUser -DisableMultiFactor
    
    .EXAMPLE
    Example -5- License all users in an Office365 Group for all plans
    Get-MsolGroupMember -GroupObjectId (Get-UnifiedGroup GroupName).ExternalDirectoryObjectId | License-EMSUser
   
    .EXAMPLE
    Example -6- License all users from a CSV
    CSV File:|  Users,usageLocation  |
    name    :|  user1@contoso.com,it |         
    file.csv:|  user2@contoso.com,us | 
    $csv = import-csv -path .\file.csv
    $csv | License-EMSUser
    
    .EXAMPLE
    Example -7- Remove all licenses from a user
    License-EMSUser user@contoso.com -disableAzureIRM -disableMultiFactor -DisableRMS -DisableIntune -DisableAADPremium
    or
    License-EMSUser user@contoso.com -RemoveEMSLicense

    
#>
Function License-EMSUser{
    [CmdletBinding()]
    Param (

    [Parameter(Mandatory=$true,ValueFromPipeline=$true,
    HelpMessage="Users parameter need a collection of users in valid SMTP address format. Input can come from explicit declaration or pipelined input. `
For pipelined input it support: CSV file with field named UserPrincipalName, EmailAddress,WindowsEmailAddress,Users,User. Get-MsolUser, `
Get-MsolGroupMember and Get-Mailbox are supported.")] $Users,

    [parameter(Mandatory=$false)]
    [switch] $RemoveEMSLicense, #remove all EMS SKU
    [parameter(Mandatory=$false)]
    [switch] $DisableRMS,#RMS_S_PREMIUM - Azure Rights Management 
    [parameter(Mandatory=$false)]
    [switch] $DisableAzureIRM, #RMS_S_ENTERPRISE - Azure Information Protection Plan 1 
    [parameter(Mandatory=$false)]
    [switch] $DisableIntune, #INTUNE_A - Intune A Direct 
    [parameter(Mandatory=$false)]
    [switch] $DisableAADPremium, #AAD_PREMIUM - Azure Active Directory Premium Plan 1 
    [parameter(Mandatory=$false)]
    [switch] $DisableMultiFactor, #MFA_PREMIUM - Azure Multi-Factor Authentication 
    
    [parameter(Mandatory=$false)]
    [ValidateSet("AF","AX","AL","DZ","AS","AD","AO","AI","AQ","AG","AR","AM","AW","AU","AT","AZ","BS","BH","BD",
                 "BB","BY","BE","BZ","BJ","BM","BT","BO","BQ","BA","BW","BV","BR","IO","BN","BG","BF","BI","CV",
                 "KH","CM","CA","KY","CF","TD","CL","CN","CX","CC","CO","KM","CD","CG","CK","CR","CI","HR","CU",
                 "CW","CY","CZ","DK","DJ","DM","DO","EC","EG","SV","GQ","ER","EE","ET","FK","FO","FJ","FI","FR",
                 "GF","PF","TF","GA","GM","GE","DE","GH","GI","GR","GL","GD","GP","GU","GT","GG","GN","GW","GY",
                 "HT","HM","VA","HN","HK","HU","IS","IN","ID","IR","IQ","IE","IM","IL","IT","JM","JP","JE","JO",
                 "KZ","KE","KI","KP","KR","KW","KG","LA","LV","LB","LS","LR","LY","LI","LT","LU","MO","MK","MG",
                 "MW","MY","MV","ML","MT","MH","MQ","MR","MU","YT","MX","FM","MD","MC","MN","ME","MS","MA","MZ",
                 "MM","NA","NR","NP","NL","NC","NZ","NI","NE","NG","NU","NF","MP","NO","OM","PK","PW","PS","PA",
                 "PG","PY","PE","PH","PN","PL","PT","PR","QA","RE","RO","RU","RW","BL","SH","KN","LC","MF","PM",
                 "VC","WS","SM","ST","SA","SN","RS","SC","SL","SG","SX","SK","SI","SB","SO","ZA","GS","SS","ES",
                 "LK","SD","SR","SJ","SZ","SE","CH","SY","TW","TJ","TZ","TH","TL","TG","TK","TO","TT","TN","TR",
                 "TM","TC","TV","UG","UA","AE","GB","UM","US","UY","UZ","VU","VE","VN","VG","VI","WF","EH","YE","ZM","ZW")] [string] $usageLocation = ""  
)

Begin{
#region MSOnline Connection Check
    If (!(Get-Command Get-MsolUser) )
    {
        If ( !(Get-Module -ListAvailable MSOnline) ){ Write-Host "Azure AD Module required, please install from: https://technet.microsoft.com/en-us/library/dn975125.aspx "}
        Else 
        {
            Write-Host "Connecting to MsOnline Powershell..."
            Import-Module MSOnline
            Connect-MsolService
        }
    }
    else
    {
        #EMS SKU from tenant
        $Error.Clear()
        $skuEMS = (Get-MsolAccountSku -ErrorAction SilentlyContinue).where({$_.accountSkuId -like "*:EMS"})
        $skuIdEMS = $skuEMS.AccountSkuId#(Get-MsolAccountSku -ErrorAction SilentlyContinue).where({$_.accountSkuId -like "*:EMS"}).AccountSkuId
        If ($Error.Count -ge 1) 
        {
            $Error.Clear()
            Write-Host "Connecting to MsOnline Powershell..."
            Import-Module MSOnline
            Connect-MsolService
            $skuEMS = (Get-MsolAccountSku -ErrorAction SilentlyContinue).where({$_.accountSkuId -like "*:EMS"})
            $skuIdEMS = $skuEMS.AccountSkuId#(Get-MsolAccountSku -ErrorAction SilentlyContinue).where({$_.accountSkuId -like "*:EMS"}).AccountSkuId
            If ($Error.Count -ge 1){exit}
        }
    }
#endregion
    
    $validCountryCodes = @("AF","AX","AL","DZ","AS","AD","AO","AI","AQ","AG","AR","AM","AW","AU","AT","AZ","BS","BH","BD", `
                 "BB","BY","BE","BZ","BJ","BM","BT","BO","BQ","BA","BW","BV","BR","IO","BN","BG","BF","BI","CV",`
                 "KH","CM","CA","KY","CF","TD","CL","CN","CX","CC","CO","KM","CD","CG","CK","CR","CI","HR","CU",`
                 "CW","CY","CZ","DK","DJ","DM","DO","EC","EG","SV","GQ","ER","EE","ET","FK","FO","FJ","FI","FR",`
                 "GF","PF","TF","GA","GM","GE","DE","GH","GI","GR","GL","GD","GP","GU","GT","GG","GN","GW","GY",`
                 "HT","HM","VA","HN","HK","HU","IS","IN","ID","IR","IQ","IE","IM","IL","IT","JM","JP","JE","JO",`
                 "KZ","KE","KI","KP","KR","KW","KG","LA","LV","LB","LS","LR","LY","LI","LT","LU","MO","MK","MG",`
                 "MW","MY","MV","ML","MT","MH","MQ","MR","MU","YT","MX","FM","MD","MC","MN","ME","MS","MA","MZ",`
                 "MM","NA","NR","NP","NL","NC","NZ","NI","NE","NG","NU","NF","MP","NO","OM","PK","PW","PS","PA",`
                 "PG","PY","PE","PH","PN","PL","PT","PR","QA","RE","RO","RU","RW","BL","SH","KN","LC","MF","PM",`
                 "VC","WS","SM","ST","SA","SN","RS","SC","SL","SG","SX","SK","SI","SB","SO","ZA","GS","SS","ES",`
                 "LK","SD","SR","SJ","SZ","SE","CH","SY","TW","TJ","TZ","TH","TL","TG","TK","TO","TT","TN","TR",`
                 "TM","TC","TV","UG","UA","AE","GB","UM","US","UY","UZ","VU","VE","VN","VG","VI","WF","EH","YE","ZM","ZW")
#region Disabled Plans options
        $disabledPlans = @()
        #available Plans in EMS sku: RMS_S_PREMIUM,INTUNE_A,RMS_S_ENTERPRISE,AAD_PREMIUM,MFA_PREMIUM
        If ($disableAzureIRM) {$disabledPlans+="RMS_S_ENTERPRISE"} 
        If ($disableRMS) {
            $disabledPlans+="RMS_S_PREMIUM"
            If( !($disabledPlans -contains "RMS_S_ENTERPRISE") ) 
            { 
                $disabledPlans+="RMS_S_ENTERPRISE" 
                # Azure Information Protection depends on AT LEAST an assigned Azure Rights Management plans
                # in any other sku (like enterprise) but not the Right Management Adhoc
            }
        }
        If ($disableIntune) {$disabledPlans+="INTUNE_A"}
        If ($disableAADPremium) {$disabledPlans+="AAD_PREMIUM"}
        If ($disableMultiFactor) {$disabledPlans+="MFA_PREMIUM"}
    
        If($disabledPlans.count -eq $skuEMS.servicestatus.Count)
        {
            Write-Host "You are disabling all the $($disabledPlans.count) plans available with EMS. I will enforce the -RemoveEMSLicense to let you save the license for another user assignment."
            $RemoveEMSLicense = $true
        }
        Else
        {
            #if nothing to be disabled, then consider all to be enabled, so passing an empty $disabledPlan is legit
            $ExcludedLicenses = New-MsolLicenseOptions -AccountSkuId $skuIdEMS -DisabledPlans $disabledPlans
            Write-Verbose "Following plans will be disabled: $($disabledPlans)"
        }
#endregion
}

Process{
     
#region Entry\User validation 
    # Given the entry from parameter or pipeline, try to match it coming from various sources
    If ( $Users.GetType().Name -eq "String")
    {
        $user = $Users
        Write-Verbose "Processing entry: `"$user`" as command line parameter"
    }
    Else
    {
        If ($Users.UserPrincipalName) {$user = $Users.UserPrincipalName; Write-Verbose "Processing entry: `"$user`" as UserPrincipalName"}
        ElseIf ($Users.EmailAddress) {$user = $Users.EmailAddress; Write-Verbose "Processing entry: `"$user`" as EmailAddress"}
        ElseIf ( ($Users.EmailAddresses -cmatch "^(SMTP)")) 
        {
            # assuming entry is coming from pipelining Exo Online Get-Mailbox, pulling out the primary SMTP address from EmailAddresses attribute
            # note that this take precendece over the next ELSEIF where address is also likely to come from same pipelined command
            # as side note, if you pull in the pipeline from Get-Mailbox, it is unlikely you will reach this stage because the UserPrincipalName is also part
            # of the cmdLet output and is the first If in the chain
            $user = ($Users.EmailAddresses -cmatch "^(SMTP)").Split(":")[1]; 
            Write-host "Processing entry: `"$user`" as Primary SMTP address from EmailAddresses"
        }
        ElseIf ($Users.WindowsEmailAddress) {$user = $Users.WindowsEmailAddress; Write-Verbose "Processing entry: `"$user`" as WindowsEmailAddress"}
        ElseIf ($Users.Users) {$user = $Users.Users; Write-Verbose "Processing entry: `"$user`" as Users CSV field"}
        ElseIf ($Users.User) {$user = $Users.User; Write-Verbose "Processing entry: `"$user`" as User CSV field"}  
    }
        
    [Bool]$shallStop = $false
    If (!($user -match "^[A-Z0-9._%+-]+@(?:[A-Z0-9-]+\.)+[A-Z]{2,}$"))
    {
        Write-host -ForegroundColor red "`"$user`" is not a valid email address"
        $shallStop = $true
    }
    Else{
        # Validate if provided value match to an email address. Technically is not needed, because the Get-MSOLUser -searchstring witll
        # try to match the attribute with partial term, but i keep it for the sake of avoid multiple results upon search.
        $userToLicense = Get-MsolUser -SearchString $user
        If (!$userToLicense)
        {
            Write-host -ForegroundColor red "User: `"$user`" Not Found using Get-MsolUser."
            $shallStop = $true
        }
    }
#endregion 



    If( !($shallStop) )    {

        If ($RemoveEMSLicense)
        {
            #remove the whole EMS Package
            If ( $userToLicense.Licenses.Where({$_.accountskuid -like "*:ems"}) )
            {
                Write-Verbose "Removing EMS license from: $($Users.UserPrincipalName)"
                Set-MsolUserLicense -UserPrincipalName $userToLicense.UserPrincipalName -RemoveLicenses $skuIdEMS
            }
            Else{;}
        }
        Else
        {
#region usageLocation validation
            If (!($userToLicense.usageLocation))
            {
                If ($Users.usageLocation)
                {
                    #coming from pipeline
                    $location = $Users.usageLocation
                }
                Else
                {
                    #coming from parameter or not coming at all
                    $location = $usageLocation
                }
     

                If (!($validCountryCodes -contains $location))
                { 
                    Write-Verbose "Invalid usageLocation: $($location) from input."
                    $location = ""
                } 
                     

                If (($location -eq "") )
                {
                    $companyLocation= ((get-MsolCompanyInformation).CountryLetterCode)
                    Write-host -ForegroundColor Cyan "User Location missing" 
                        Get-Help Set-MsolUser -Parameter UsageLocation
                    Write-host -ForegroundColor Cyan "Type here the Country location code" -NoNewline
                    Write-host -ForegroundColor White "[Default:" -NoNewline
                    Write-host -ForegroundColor Yellow "$($companyLocation)" -NoNewline
                    Write-host -ForegroundColor White "]" -NoNewline
                    Write-host -ForegroundColor Cyan ": " -NoNewline
                    $location  = read-host
                    $location = $location.Trim()

                    If (!($location.Length -lt 2))
                    {
                        # Location entered is longer than 2 chars, not using companyLocation and now validating the country code
                        If ( !($validCountryCodes -contains $location) ) 
                        {
                            # bad country code, aborting or using companyLocation
                            If (!(Read-Host "Provided location `"$($location)`" is not valid. Hit enter to use `" $($companyLocation)`" or type something else to exit:"))
                            {
                                # use companyLocation as usageLocation
                                $location = $companyLocation
                            }
                            Else
                            {
                                exit; #aborting Execution
                            }
                        }
                        Else
                        {
                            # good country code, nothing to do here as $location seems valid
                        }
                    }
                    Else
                    {
                        # location less than 2 chars, assuming user hit enter and decided to use companyLocation
                            $location = $companyLocation
                    }
                }
    
                Try{
                    Write-Verbose "Assigning user location: `"$($location)`" to `"$($userToLicense.UserPrincipalName)`""
                    Set-MsolUser -UserPrincipalName $userToLicense.UserPrincipalName -UsageLocation $location
                }
                Catch{
                    Write-Host -ForegroundColor red "Something bad happened while assigninig user location, aborting"
                    Exit
                } 
            }     
            Else
            {
                # for sake of honesty, this line is useless as the location will not be used anylonger aferward in the code
                #$location = $userToLicense.usageLocation
            } 
    
#endregion

#region Plans management 

            If (!($userToLicense.Licenses))
            {
                # assigning license, location assingment should be succeded at this stage
                Write-verbose "Assigning license: `"$($skuIdEMS)`" to `"$($userToLicense.UserPrincipalName)`""
                Set-MsolUserLicense -UserPrincipalName $userToLicense.UserPrincipalName -AddLicenses $skuIdEMS -LicenseOptions $ExcludedLicenses
            } 
            Else 
            {
            # if user has any licenses he may not have the EMS one already, or he may have but without assigned plans
                If ( $userToLicense.Licenses.Where({$_.accountskuid -like "*:ems"}) )
                {
                    # EMS license present, no need to use the -addLicense parameter
                    Write-Verbose "Assigning license: `"$($skuIdEMS)`" to `"$($userToLicense.UserPrincipalName)`""
                    Set-MsolUserLicense -UserPrincipalName $userToLicense.UserPrincipalName -LicenseOptions $ExcludedLicenses
                }
                Else
                {
                    # if not EMS license, then pass the -addLicenses paramter too
                    Write-Verbose "Assigning license: `"$($skuIdEMS)`" to `"$($userToLicense.UserPrincipalName)`""
                    Set-MsolUserLicense -UserPrincipalName $userToLicense.UserPrincipalName -AddLicenses $skuIdEMS -LicenseOptions $ExcludedLicenses
                }
            }               
#endregion
        }
    } #end if $shallStop    
    Else
    { 
        #Write-host -ForegroundColor red "$user is not a valid email address"
    }
}

End{ Write-Verbose "Completed"}

}