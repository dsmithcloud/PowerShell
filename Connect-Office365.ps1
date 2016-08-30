function Connect-Office365{
<#
   .SYNOPSIS
        Connects remote PowerShell sessions for Office 365

   .DESCRIPTION
        This script automates connectivity to remote PowerShell sessions in Office 365 for:
            Azure Active Directory
            Exchange Online
            Lync Online
            SharePoint Online
   
   .PARAMETER ProxyEnabled
        Specifies that IE Proxy settings should be utilized.

   .PARAMETER AzureAD
        Specifies that a connection to Azure Active Directory should be made.
   
   .PARAMETER Exchange
        Specifies that a connection to Exchange Online should be made.
   
   .PARAMETER Skype
        Specifies that a connection to Skype for Buisness Online should be made.
   
   .PARAMETER SharePoint
        Specifies that a connection to SharePoint Online should be made.
   
   .PARAMETER TenantName
        Specifies the tenant name for connecting to SharePoint Online
   
   .PARAMETER Credential
        Specifies the credentials used to connect to Office 365.
   
   .EXAMPLE 
        PS C:\> Connect-Office365

        cmdlet Connect-Office365 at command pipeline position 1
        Supply values for the following parameters:
        Credential
        Connecting to Azure Active Directory

   .EXAMPLE 
        PS C:\> Connect-Office365 -Exchange

        cmdlet Connect-Office365 at command pipeline position 1
        Supply values for the following parameters:
        Credential
        Connecting to Azure Active Directory
        Connecting to Exchange Online
        WARNING: Your connection has been redirected to the following URI:
        "https://pod12345psh.outlook.com/powershell-liveid?PSVersion=4.0 "
        WARNING: The names of some imported commands from the module 'tmp_ljq4d0bh.w4o' include unapproved verbs that might
        make them less discoverable. To find the commands with unapproved verbs, run the Import-Module command again with the
        Verbose parameter. For a list of approved verbs, type Get-Verb.
    
        ModuleType Version    Name                                ExportedCommands
        ---------- -------    ----                                ----------------
        Script     1.0        tmp_ljq4d0bh.w4o                    {Add-AvailabilityAddressSpace, Add-DistributionGroupMember...

   .EXAMPLE 
        PS C:\> Connect-Office365 -SharePoint -TenantName contoso

        cmdlet Connect-Office365 at command pipeline position 1
        Supply values for the following parameters:
        Credential
        Connecting to Azure Active Directory
        Connecting to SharePoint Online for contoso

   .EXAMPLE 
        PS C:\> Connect-Office365 -Credential admin@contoso.com
        Connecting to Azure Active Directory

   .EXAMPLE 
        PS C:\> $credentials = Get-Credential

        cmdlet Get-Credential at command pipeline position 1
        Supply values for the following parameters:
        Credential
    
        PS C:\> Connect-Office365 -Credential $credentials
        Connecting to Azure Active Directory

   .NOTES 
        NAME: Microsoft.PowerShell_profile.ps1 
        AUTHOR: David Smith
        LASTEDIT: 01/13/2015
        KEYWORDS: Office365, Exchange Online, Skype for Business Online, SharePoint Online, Azure AD
        VERSION: 2.1 (added #Requires features for all required modules)
        The script are provided “AS IS” with no guarantees, no warranties, and they confer no rights.
   
   .LINK 
        Blog Link
            Http://www.texmx.net/

   .LINK
        Microsoft Online Services Sign-In Assistant for IT Professionals RTW
            http://www.microsoft.com/en-us/download/details.aspx?id=41950

   .LINK
        Azure Active Directory Module for Windows PowerShell (64-bit version)
            http://go.microsoft.com/fwlink/p/?linkid=236297

   .LINK
        Windows PowerShell Module for Skype for Business Online
            http://www.microsoft.com/en-us/download/details.aspx?id=39366

   .LINK
        SharePoint Online Management Shell
            http://www.microsoft.com/en-us/download/details.aspx?id=35588
   
#>

    [CmdLetBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [Switch]$ProxyEnabled,

        [Parameter(Mandatory=$false)]
        [Switch]$AzureAD,
        
        [Parameter(Mandatory=$false)]
        [Switch]$Exchange,

        [Parameter(Mandatory=$false)]
        [Switch]$Skype,

        [Parameter(Mandatory=$false)]
        [Switch]$SharePoint,

        [Parameter(Mandatory=$false)]
        [String]$TenantName,

        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.CredentialAttribute()]
        $Global:Credential
    )

    
    if($ProxyEnabled) {
        $MySessionOption = New-PsSessionOption -ProxyAccessType IEConfig -ProxyAuthentication basic
    } Else {
        $MySessionOption = New-PSSessionOption -ProxyAccessType None -ProxyAuthentication Negotiate
    }

    if($AzureAD) {
        Write-Host "Connecting to Azure Active Directory" -ForegroundColor Yellow
        Import-Module MSOnline
        Connect-MsolService -Credential $Credential 
    }

    if($Exchange) {
        if ((Get-PSSession | ?{$_.ConfigurationName -eq "Microsoft.Exchange"}).State -ne "Opened") {
            Write-Host "Connecting to Exchange Online" -ForegroundColor Yellow
            $ExchangeSession = New-Pssession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/PowerShell-liveId -Credential $Credential -Authentication Basic -AllowRedirection -sessionOption $MySessionOption
            Import-PSSession $ExchangeSession
        } Else {
            Write-Host "Connection to Exchange Online exists for " -NoNewline -ForegroundColor Yellow
            Write-Host (Get-OrganizationConfig).Name -ForegroundColor White
        }
    }
    
    if($Skype) {
        if ((Get-PSSession | ?{$_.ComputerName -like "*.online.lync.com"}).State -ne "Opened") {
            Write-Host "Connecting to Skype for Business Online" -ForegroundColor Yellow
            $SkypeSession = New-CsOnlineSession -Credential $Credential     -SessionOption $MySessionOption
            Import-PSSession $SkypeSession
        } Else {
            Write-Host "Connection to Skype for Business Online exists for " -NoNewline -ForegroundColor Yellow
            Write-Host (Get-CSTenant).DisplayName -ForegroundColor White
        }
    }

    if($SharePoint) {
        if ($TenantName -eq $null) {
            Write-Host "Please enter Tenant Name" -ForegroundColor Yellow
            $TenantName = Read-Host
        }
        Get-SPOSite -Identity ("https://"+$TenantName+".sharepoint.com/") -ErrorAction SilentlyContinue
        if ($?) {
            Write-Host "Connection to SharePoint Online exists for " -NoNewline -ForegroundColor Yellow
            Write-Host $TenantName -ForegroundColor White
        } Else {
            $SPOURL = ("https://"+$TenantName+"-admin.sharepoint.com")
            Write-Host ("Connecting to SharePoint Online for ") -ForegroundColor Yellow -NoNewline
            Write-Host $TenantName -ForegroundColor White
            Connect-SPOService -Url $SPOURL -Credential $Credential    
        }
    }
}
 
function Disconnect-Office365 {
    [CmdLetBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [Switch]$All=$True,

        [Parameter(Mandatory=$false)]
        [Switch]$Exchange,

        [Parameter(Mandatory=$false)]
        [Switch]$Skype,

        [Parameter(Mandatory=$false)]
        [Switch]$SharePoint
    )
    
    if ($Exchange -or $Skype -or $SharePoint) {$All = $False}

    if ($All -or $Exchange)   {Write-Host "Disconnecting Exchange Online" -ForegroundColor Yellow; Get-PSSession | ?{$_.ComputerName -like "*outlook.com"} | Remove-PSSession}
    if ($All -or $Skype)      {Write-Host "Disconnecting Skype for Business Online" -ForegroundColor Yellow; Get-PSSession | ?{$_.ComputerName -like "*lync.com"} | Remove-PSSession}
    if ($All -or $SharePoint) {Write-Host "Disconnecting SharePoint Online" -ForegroundColor Yellow;Disconnect-SPOService -ErrorAction SilentlyContinue}
}


