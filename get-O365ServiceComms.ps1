#region Header
<#
get-O365ServiceComms.ps1 in Office365_Service_Comms_2SCP Repo - DL Cooper - University of Pennsylvania 2021

    Use Office 365 Service Communications API to retrieve all message types, temp store in ASCII, send via SCP to Linux Host for parsing/web display
        Simplified, but activaly working to produce results

#DLC 01/03/2020 - MS removed some attributes, added others
#DLC 01/03/2020 - Microsoft Changed Output - Updating connection and information pull
#DLC 01/07/2020 - Added Deliminator (|||||) to end of each record retrieved with each API pull
#DLC 2021MAY21 - Convert for Public Repo (remove any specific Penn Data/Info)

#Reference: https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference


    Instructions:
        Create AzureAD App with 'Office 365 Management APis' Permissions:
                        ServiceHealth.Read / Type: Application / Admin Consent Required: Yes

            Call with splats or 
        Edit $APIauthSettings hashtable with AzureAD App information
        Edit $SCPauthSettings hashtable with target SCP host information


#>
#endregion
param(
    # API Auth Settings Splat - not mandatory - edit below if not providing via script call
    [Parameter(Mandatory = $false)]
    [hashtable]
    $APIAuthSplat,
    # SCP Auth Settings splat - not mandatory - edit directly in script below if not providing via script call
    [Parameter(Mandatory = $false)]
    [hashtable]
    $SCPauthSplat,
    # Start Transcript
    [Parameter(Mandatory = $false)]
    [bool]
    $StartTranscript = $false,
    # UseSCP - True if using SCP to send to target system (as configured)
    [Parameter(Mandatory = $false)]
    [bool]
    $UseSCPtoSend = $false 
)

#region Globals

## for export to SCP target
if (!(import-module WinSCP)) { install-module WinSCP; import-module WinSCP }

if ($APIAuthSplat) { 
    $global:APIauthSettings = $APIAuthSplat 
}
if (!($APIAuthSplat)) {
    write-host -f green "No GraphAPI Auth Settings provided, using settings in script"

    # Settings used for URL builds and Auth Token retrieval
    $global:APIauthSettings = @{
        "TenantDomain" = "contoso.com"
        "Tenant"       = "12345678-9abc-defg-hijk-lmnopqrstuvw"
        "ClientID"     = "wvutsrqp-onml-kjih-gfed-cba987654321"
        "redirectUri"  = "urn:ietf:wg:oauth:2.0:oob"
        "clientSecret" = "ObtainSecurely"
    }
}

if ($SCPauthSplat) { 
    
    $global:SCPauthSettings = $SCPauthSplat 
}
if (!($SCPauthSplat)) {
    write-host -f green "No SCP Auth Settings provided, using settings in script"
    
    $global:SCPauthSettings = @{
        #if SFTP needed instead of SCP, see UseSCP function options
        "hostName"              = "SCPHostName.Some.Host"
        "portNumber"            = "22"
        "userName"              = "someGreatUser"
        "password"              = "ObtainSecurly"
        "SshHostKeyFingerprint" = "ecdsa-sha2-nistp256 256 aa:bb:cc:dd:ee:ff:gg:hh:ii:jj:00:11:22:33:44:55"
        "remotePutDirectory"    = "/home/someGreatUser/"
    }
}

#URLs to Run
$messageCenterURLs = @{

    "CurrentStatus"    = "https://manage.office.com/api/v1.0/$($APIAuthSettings.TenantDomain)/ServiceComms/CurrentStatus"
    "HistoricalStatus" = "https://manage.office.com/api/v1.0/$($APIAuthSettings.TenantDomain)/ServiceComms/HistoricalStatus"
    "Messages"         = "https://manage.office.com/api/v1.0/$($APIAuthSettings.TenantDomain)/ServiceComms/Messages"

}

# Deliminator to add after each record (for easier sorting/seperation after export)
$global:Delim = "|||||"

#logs Path, if not exist create
$thisPath = (Get-Item -Path ".\" -Verbose).FullName
if (!(Test-Path ($thisPath + "\logs\"))) {
    New-Item -ItemType Directory -Force -Path ($thisPath + "\logs\")
}
#Exports Path, if not exist create
if (!(Test-Path ($thisPath + "\exports\"))) {
    New-Item -ItemType Directory -Force -Path ($thisPath + "\exports\")
}

[string]$fileDateTime = (get-date -Format yyyyMMMdd_hhmmss_)

#Log path for start-trascript recording, if needed
if ($StartTranscript) {
    $thisLogPath = $thisPath + "\logs\"
    $thisTranscriptFile = $thisLogPath + $fileDateTime + "$($tenantDomainFileName)_Transcript.txt"
    Start-Transcript -LiteralPath $thisTranscriptFile -IncludeInvocationHeader
}

#Export path for local File storage
$thisExportsPath = $thisPath + "\exports\"

$tenantDomainFileName = ($APIauthSettings.TenantDomain).Replace(".", "_")
$exportTXT = $thisExportsPath + $fileDateTime + "$($tenantDomainFileName)_ServiceComms.txt"
# $exportCSV = $thisExportsPath + $fileDateTime + "$($tenant)_ServiceComms.csv"

$formatEnumPre = $FormatEnumerationLimit
$FormatEnumerationLimit = -1

#endregion

#region Functions
function get-accesstoken {
    [CmdletBinding()]
    param($tenant, $clientID, $redirectURL, $clientSecret)
    try {
        [string]$randomState = Get-Random -Minimum 1 -Maximum 64
        [uri]$auth2URL = "https://login.microsoftonline.com/$($tenant)/oauth2/token"
        $result = Invoke-RestMethod $auth2URl.AbsoluteUri  `
            -Method Post -ContentType "application/x-www-form-urlencoded" `
            -Body @{client_id = $clientId; 
            client_secret     = $clientSecret; 
            redirect_uri      = $redirectURL; 
            grant_type        = "client_credentials";
            resource          = "https://manage.office.com";
            state             = $randomState
        } -ErrorVariable InvokeError
  
        if ($null -ne $result) { return $result }
    }
    catch {
        write-output "Could not retrieve Auth Token"
        # Exception is stored in the automatic variable _
        write-output $InvokeError
        BREAK
    }
  
}
function get-authheader {

    $accesstoken = Get-AccessToken -tenant $APIauthSettings.Tenant -ClientID $APIauthSettings.ClientID -redirectURL $APIauthSettings.redirectUri -clientSecret $APIauthSettings.clientSecret

    $token = $accesstoken.Access_Token
    $tokenexp = $accesstoken.expires_on

    ## Debug
    # write-output ""
    # write-output "AuthToken Retrieved"
    # write-output ""
    # write-output "Token Expiration Date:"
    # write-output "$tokenexp"   
    
    $global:authHeader = @{
        'Content-Type'  = 'application/json'
        'Authorization' = "Bearer " + $token
        'ExpiresOn'     = $tokenexp
    }


    ## Debug
    # $global:AHexpLocal = (ConvertFromCtime -ctime $authheader.ExpiresOn).tolocaltime()

    # write-output "Auth Header Token Expires:"
    # write-output "$AHexpLocal"
    # write-output "Token DateTime stored in Global `$AHexpLocal variable."
}
function ConvertFromCtime ([Int]$ctime) {
    [datetime]$epoch = '1970-01-01 00:00:00'    
    [datetime]$result = $epoch.AddSeconds($Ctime)
    return $result
}
function get-serviceCommsInfo {
    param(
        # URL
        [Parameter(Mandatory = $true)]
        [string]
        $URL,
        # filter, if needed
        [Parameter(Mandatory = $false)]
        [string]
        $urlFilter,
        # top, if needed
        [Parameter(Mandatory = $false)]
        [string]
        $urlTop
    )

    $thisURL = new-object System.UriBuilder -ArgumentList $URL

    if ($urlFilter) { $thisURL.Query = "filter=" + $urlFilter }
    if ($urlTop) { $thisURL.Query = "top" + $urlTop }

    $ServCommsInfo = Invoke-RestMethod -Method Get -Uri $thisURL.Uri -Headers $authHeader -ErrorVariable GetServiceCommsErr -ContentType "application/json"

    if ($GetServiceCommsErr) { return $GetGroupInfoViaGUIDError }
    else { return $ServCommsInfo.value }

}
function UseSCP {
    param
    (
        [parameter(Mandatory = $true)]
        [string]
        $eventsfile
    )
	
    try {

        # Setup session options
        $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
            #winscp.exe path if needed (not in local PATH, etc)
            #ExecutablePath = "C:\Program Files (x86)\WinSCP"
            #Protocol = [WinSCP.Protocol]::Sftp
            Protocol              = [WinSCP.Protocol]::Scp
            HostName              = $SCPauthSettings.hostName
            PortNumber            = $SCPauthSettings.portNumber
            UserName              = $SCPauthSettings.userName		
            Password              = $SCPauthSettings.password
            SshHostKeyFingerprint = $SCPauthSettings.SshHostKeyFingerprint			
            #	PrivateKeyPassphrase = $SSHRSAPubKey
        }
        
        $session = New-Object WinSCP.Session
		
        try {
            # Connect
            $session.Open($sessionOptions)
		
            # Upload files
            $transferOptions = New-Object WinSCP.TransferOptions
            $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
            #$transferOptions.TransferMode = [WinSCP.TransferMode]::Automatic
			
            #$session.PutFiles('localPath', 'remotePath', $remove, $options)
            $transferResult = $session.PutFiles($eventsfile, $($SCPauthSettings.remotePutDirectory), $false, $transferOptions)
            
			
            # Throw on any error
            $transferResult.Check()
			
            # Print results
            foreach ($transfer in $transferResult.Transfers) {
                write-output ("Upload of {0} succeeded" -f $transfer.FileName)
            }
        }
        finally {
            # Disconnect, clean up
            $session.Dispose()
        }
		
        #exit 0
    }
    catch [Exception] {
        write-output ("Error: {0}" -f $_.Exception.Message)
        #exit 1
    }
	
}
#endregion

#region Process
get-authheader

## uses URLs in $messageCenterURLs Hash to retrieve all available messages
### if filter or top needed - see get-serviceCommsInfo function to run seperately
$reportCollection = @()
$messageCenterURLProcessCount = 0
foreach ($messageCenterURL in $messageCenterURLs.Keys) {
    $messageCenterURLProcessCount++
    $messageCenterURLsKeysCount = ($messageCenterURLs.Keys | measure).Count

    Write-Progress -Activity "Processing URLs - $messageCenterURL" -CurrentOperation "$messageCenterURLProcessCount of $messageCenterURLsKeysCount" -Status 'Progress->' -Id 0 -PercentComplete (($messageCenterURLProcessCount / $messageCenterURLsKeysCount) * 100 )

    $thisPull = @(get-serviceCommsInfo -URL $messageCenterURLs[$messageCenterURL])

    $thisPullProcessCount = 0
    foreach ($thisPullEntry in $thisPull) {
        $thisPullProcessCount++
        $thisPullCount = ($thisPull | measure).Count

        Write-Progress -Activity "Processing Records" -CurrentOperation "$messageCenterURL - $thisPullProcessCount of $thisPullCount" -Status 'Progress' -id 1 -ParentId 0 -PercentComplete (($thisPullProcessCount / $thisPullCount) * 100)

        #Adds Delim after each record
        Add-Member -InputObject $thisPullEntry -MemberType NoteProperty -Name Delim -Value $global:Delim

        $reportCollection += $thisPullEntry

    }

    Write-Progress -Activity "Processing Records" -CurrentOperation "$messageCenterURL - $thisPullProcessCount of $thisPullCount" -Status "Complete" -id 1 -ParentId 0 -Completed


}

Write-Progress -Activity "Processing URLs" -CurrentOperation "$messageCenterURLProcessCount of $messageCenterURLsKeysCount" -Status "Complete" -Id 0 -Completed

#endregion

#region Export

#export to local TXT file as needed for Target (SCP) processing
$reportCollection | out-file $exportTXT -Encoding ASCII -ErrorAction Inquire
write-host -f Magenta "Report Collection exported to:"
write-host -f white -b DarkMagenta $($exportTXT)

#Use SCP to transer to target
if ($UseSCPtoSend) {
    UseSCP -eventsfile $exportTXT -ErrorAction Inquire
}
#endregion

#region Cleanup

# remove local file
#remove-item $exportTXT

#cleanup globals
remove-variable APIauthSettings
remove-variable SCPauthSettings

#set $formatEnumerationLimit back to original value
$FormatEnumerationLimit = $formatEnumPre

if ($StartTranscript) { Stop-Transcript }
#endregion