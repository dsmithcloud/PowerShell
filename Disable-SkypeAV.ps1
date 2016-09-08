<#
.SYNOPSIS
    Script that disables Audio/Video functionality for all Skype Users
    except those that are explicitly allowed by virtue of membership in a specified Group

.DESCRIPTION
    Script that disables Audio/Video functionality for all Skype Users
    except those that are explicitly allowed by virtue of membership in a specified Group

    By default, that group name is Office365-AllowSkypeAV but any group name can be provided
    as long as the group is available in Azure AD in the tenant.

    You can provide your own group names at the command line with the parameters:

.PARAMETER UserID
    Specifies the userid to be used to authenticate to the Office 365 tenant.
   
.PARAMETER Password
    Specifies the password to be used to authenticate to the Office 365 tenant.

.PARAMETER AVAllowedGroup
    Specifies a custom name for the security group for users allowed to use Audio/Video.  The default value is Office365-AllowSkypeAV
    
.PARAMETER From
    The SMTP address that the email should be sent from.

.PARAMETER To
    The SMTP address that the email should be sent to.

.PARAMETER SMTPServer
    The SMTP server FQDN or IP address that should be used to relay the email.

.EXAMPLE 
    PS C:\> .\Disable-SkypeAV.ps1 -UserId admin@contoso.onmicrosoft.com -password pass@word1 -SIPDomain cumulusnerd.com

.EXAMPLE 
    PS C:\> .\Disable-SkypeAV.ps1 -UserId admin@contoso.onmicrosoft.com -password pass@word1 -From: david@cumulusnerd.com -To david@texmx.net -SMTPServer smtp.contoso.com

.EXAMPLE
    PS C:\> .\Disable-SkypeAV.ps1 -UserId admin@contoso.onmicrosoft.com -password pass@word1 -AVAllowedGroup "All Employees"
            
.NOTES 
    AUTHOR: David Smith
    LASTEDIT: 03/09/2016
    KEYWORDS: Office365, Skype
    Blog: texmx.net
    Email: david.smith@quisitive.com
    The script is provided “AS IS” with no guarantees, no warranties, and confers no rights
        
.LINK 
    Blog Link http://www.texmx.net/
    Quisitive http://www.Quisitive.com/
    
#>

# Establish script parameters
[CmdletBinding()]
param (
        [parameter(Position=0,Mandatory=$true)] [string]$UserId,
        [parameter(Position=1,Mandatory=$true)] [string]$Password,
        
        [parameter(Mandatory=$false)] [string]$AVAllowedGroup =  "Office365-AllowSkypeAV",

        [parameter(Mandatory=$false,ParameterSetName='SendLogs')]  [string]$From,
        [parameter(Mandatory=$false,ParameterSetName='SendLogs')]  [string]$To,
        [parameter(Mandatory=$false,ParameterSetName='SendLogs')]  [string]$SMTPServer
)


# Establish WriteTo-Log function
function WriteTo-Log
 {
    param (
        [string]$String="*",
        [string]$Logfile = $Logfile,
        [Switch]$OutputToScreen,
        [ValidateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
        [String]$ForegroundColor=(Get-Host).ui.RawUI.ForegroundColor,
        [ValidateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
        [String]$BackgroundColor=(Get-Host).ui.RawUI.BackgroundColor
        )
    
    if ($LogFile -eq "") {
        $LogFile = ('.\'+(Get-History -Id ($MyInvocation.HistoryId -1) | select StartExecutionTime).startexecutiontime.tostring('yyyyMMdd-HHmm')+'-'+[io.path]::GetFileNameWithoutExtension($MyInvocation.ScriptName)+'.log')
    }

    if (!(Test-Path $LogFile)) {
        Write-Output "Creating log file $LogFile"
        $LogFile = New-Item $LogFile -Type file
    }

	$datetime = (Get-Date).ToString('yyyyMMdd HH:mm:ss')
    $StringToWrite = "$datetime | $String"
	if ($OutputToScreen) {Write-Host $StringToWrite -ForegroundColor $ForegroundColor -BackgroundColor $BackgroundColor}
    Add-Content -Path $LogFile -Value $StringToWrite
 }


# Establish Connect-Skype4BOnline function
function Connect-Skype4BOnline
{
    $cred=New-Object System.Management.Automation.PSCredential $userid,(ConvertTo-SecureString -String $password -AsPlainText -Force)
    Import-Module LyncOnlineConnector
    $Script:cssession = New-CsOnlineSession -Credential $cred
    Import-PSSession $cssession -AllowClobber
}



# Establish Connect-AzureActiveDirectory function
function Connect-AzureActiveDirectory
{
    $cred=New-Object System.Management.Automation.PSCredential $userid,(ConvertTo-SecureString -String $password -AsPlainText -Force)
    Import-Module MSOnline
    Connect-MsolService -Credential $cred
}


# Establish value for $Logfile name to be used throughout the script when WriteTo-Log is called
[string]$Script:Logfile = '.\'+(Get-Date).tostring('yyyyMMdd-HHmm')+'-Disable-SkypeAV.log'



# Connect to Skype for Business Online
WriteTo-Log "Logging on to Azure Active Directory with user id $userid" -OutputToScreen -ForegroundColor White
Connect-AzureActiveDirectory
WriteTo-Log "Logging on to Skype For Business with user id $userid" -OutputToScreen -ForegroundColor White
Connect-Skype4BOnline



# Retrieve list of In-Scope Skype Users
WriteTo-Log "Retrieving list of Skype Users" -OutputToScreen -ForegroundColor White
# Search for all valid Skype users.  Users without a Skype license will have an empty SipAddress, so exclude them from the returned results.
$InScopeUsers = get-csonlineuser -ResultSize Unlimited | ?{$_.SipAddress -ne ""}  
WriteTo-Log ("Found "+$InScopeUsers.count+" Skype users") -OutputToScreen -ForegroundColor Green



# Retrieve list of users with Protocol Allowance Exception
[System.Collections.ArrayList]$SkypeAVAllowed = Get-MsolGroupMember -GroupObjectId (Get-MsolGroup -SearchString $AVAllowedGroup).objectid
WriteTo-Log ("Count of users with Audio/Video exception "+$SkypeAVAllowed.count) -OutputToScreen



# Begin processing of in-scope users
foreach ($user in $InScopeUsers) {

    # Set the ProtocolEnabled Variables
    if ($user.userprincipalname -in $SkypeAVAllowed.EmailAddress) {$SkypeAVDisabled = $false} else {$SkypeAVDisabled = $true}


    # Attempt to set the AudioVideoDisabled values on the Skype User if not already set as required
    if ($user.AudioVideoDisabled -ne $SkypeAVDisabled) {
        try {
            Set-CsUser -Identity $user.SipAddress -AudioVideoDisabled $SkypeAVDisabled -ErrorAction stop
            WriteTo-Log ("Updating user "+$user.UserPrincipalName) -OutputToScreen -ForegroundColor Green 
            WriteTo-Log ("OLD Value: AudioVideoDisabled: "+$user.AudioVideoDisabled) -OutputToScreen -ForegroundColor Green
            WriteTo-Log ("NEW Value: AudioVideoDisabled: "+$SkypeAVDisabled) -OutputToScreen -ForegroundColor Green
        }
        catch {
            WriteTo-Log ("ERROR: Could not update "+$user.UserPrincipalName) -OutputToScreen -ForegroundColor Red 
            WriteTo-Log $Error[0].Exception -ForegroundColor Red -OutputToScreen
        }
    } Else {
        WriteTo-Log ("Skipping user "+$user.UserPrincipalName+".  A/V settings already correct.")
    }
}



# Disconnect from Skype for Business Online
WriteTo-Log "Disconnecting from Skype for Business Online" -OutputToScreen -ForegroundColor White 
Remove-PSSession $cssession


# Send the log via Email if From, To and SMTPServer parameters are used
If ($PSBoundParameters.ContainsKey('From') -and $PSBoundParameters.ContainsKey('To') -and $PSBoundParameters.ContainsKey('SMTPServer')) {
    WriteTo-Log "Emailing Log file" -OutputToScreen
    WriteTo-Log "From: $From" -OutputToScreen
    WriteTo-Log "To: $To" -OutputToScreen
    WriteTo-Log "Subject: $Logfile" -OutputToScreen
    WriteTo-Log "Attachment Name: $Logfile" -OutputToScreen
    Send-MailMessage -From $From -To $To -Subject $Logfile -Attachments $Logfile -SmtpServer $SMTPServer
}