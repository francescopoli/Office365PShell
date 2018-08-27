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

<#
    Extract Office 365 endpoints IP Addresses and generate csv file from it
#>

<#       
    .DESCRIPTION
        Get the Office 365 Endpoints using the webservice endpoint and generate per 
        service csv file (and per service\tcp port txt files).

        Once the CSV files are generated, open them with Excel and convert the content
        to a table + use the Top Align feature after selecting all the content.
        To convert only the csv content, and not all the csv to a table in Excel, try:
        - Open the file with Excel
        - Select cell A:1
        --- Keep ctrl + Shift pressed
        --- press Right Arrow 
        --- press Down Arrow
        Use the convert to table from the menu` and then select the whole table and 
        use the Align Top format tool

    .PARAMETER $Path
       "path\" where to save the csv files. If omitted ".\" will be used.
	.PARAMETER $GenerateTXT
        Switch parameter, if passed with -GenerateTXT it will cause the scrip to generate
        a per Service folder, containing a list of txt files named as the TCP port used
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$false,ValueFromPipeline=$false)] [string] $Path = ".\",
    [Parameter(Mandatory=$false,ValueFromPipeline=$false)] [switch] $GenerateTXT = $false

)
$json = Invoke-RestMethod -Method GET `
        -Uri "https://endpoints.office.com/endpoints/worldwide?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7"

foreach ($entry in $json | Sort-Object serviceArea){

    $path = $path.Trim()
    if ($path[-1] -ne "\") {$path = $path + "\"}

    if (!(Test-Path "$($path)$($entry.serviceArea)")){
        New-Item -ItemType Directory -Path "$($path)" -Name $entry.serviceArea
    }

    if ($generateTXT){
        foreach($url in $entry.urls){
            $url | Out-File -FilePath "$($path)$($entry.serviceArea)\$($entry.tcpports).txt" -Append
        }

        foreach($ip in $entry.ips){
            $ip | Out-File -FilePath "$($path)$($entry.serviceArea)\$($entry.tcpports).txt" -Append
        }
    }

    $u = "" # urls
    foreach ($url in $entry.urls){ $u = $u + $url.tostring() + "`n" }
    if ($u.Length -gt 0){ $u = $u.Substring(0,$u.Length-1) } 
    
    $i = "" # ip
    foreach ($ip in $entry.ips){ $i = $i + $ip.ToString() + "`n" }
    if ($i.Length -gt 0){$i = $i.Substring(0,$i.Length-1)}

    $p = "" # ports
    foreach ($port in $entry.tcpports){ $p = $p + $port.tostring() + ";`n" }
    if ($p.Length -gt 0){ $p = $p.Substring(0,$p.Length-1) } 

    $props = [ordered]@{}
    $props.'ServiceArea' = $entry.serviceArea
	$props.'ServiceAreaDisplayName' = $entry.ServiceAreaDisplayName
	$props.'Urls' = $u
	$props.'Ips' = $i
	$props.'TcpPorts' = $p
	$props.'expressRoute' = $entry.expressRoute
	$props.'category' = $entry.category;
    $props.'Required' = $entry.Required;

    # Out a PsObject with all the properties defined
	$csv = New-Object -TypeName PSObject -Property $props
    $csv | Export-Csv -NoTypeInformation -Path ".\$($entry.serviceArea).csv" -Append -Delimiter "," -NoClobber

}