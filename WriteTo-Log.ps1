<#
.SYNOPSIS
    This function generates a log file to be used in a script and adds a new line to the log file each time
    it's called, including a date/time stamp of when the line was added to the log.

.DESCRIPTION
    This function generates a log file to be used in a script and adds a new line to the log file each time
    it's called, including a date/time stamp of when the line was added to the log.
   
.PARAMETER String
    Specifies the string of text that should be added to the log file

.PARAMETER Logfile
    Specifies the file name and path of the log file to be used.                      
        * If NO -Logfile parameter is passed, the log file created by default will match YYYYMMDD-HHmm-<Log File Name>.log
        * If -Logfile parameter is passed with a string value, that value will be used as the log file name
        * if a $Script:Logfile variable is defined in the main body of the script, there will be no need to pass
          the -Logfile parameter and the string value defined will be used each time the function is called.
   
.PARAMETER OutputToScreen
    Specifies that the new log file entry (including date/timestamp) should be output on the screen
   
.PARAMETER ForegroundColor
    Specifies the foreground color of the screen output and works as it would for Write-Host
   
.PARAMETER BackgroundColor
    Specifies the background color of the screen output and works as it would for Write-Host

.EXAMPLE 
    WriteTo-Log "Logging on to Azure Active Directory with user id $userid" -OutputToScreen -ForegroundColor White

.EXAMPLE 
    [string]$Logfile = 'c:\Logs\MyLogFile.txt'
    WriteTo-Log $Error[0].Exception -OutputToScreen -ForegroundColor Red -Logfile $Logfile

.EXAMPLE
    [string]$Script:Logfile = '.\'+(Get-Date).tostring('yyyyMMdd-HHmm')+'-LogFile.txt'
    WriteTo-Log -String "Logging off" -OutputToScreen

.NOTES 
    AUTHOR: David Smith
    LASTEDIT: 08/05/2014
    KEYWORDS: PowerShell, Scripting
    BLOG: texmx.net
    EMAIL: david.smith@Quisitive.com
    COMMENTS: 	-Modified 10/19/2015 - Added Foreground and Background color 
		         options for output to screen and changed OutputToScreen variable
		         to a switch.
                -Modified 1/12/2016 - Added functionality to set default log file
                 name include Date/Time string plus Script File name plus .log extension
                 Example: 2016-01-12-02.03-test-script.log
                 Also shortened the log entry date/time header to yyyy-MM-dd hh.mm.ss format
                -Modified 1/27/2016 - Changed log entry date/time header to yyyy-MM-dd HH:mm:ss format
                -Modified 3/8/2016 - Change the functionaly of the default log file naming.  
                    * If no -Logfile parameter is passed, the log file created by default will match YYYYMMDD-HHmm-<Log File Name>.log
                    * If -Logfile parameter is passed, the string value passed will be used as the log file name
                    * if a $Script:Logfile variable is defined in the main body of the script, there will be no need to pass
                      a -Logfile parameter and the string value defined will be used each time the function is called.
                Also modified the date/time format for log entries to match yyyyMMdd HH:mm:ss and use a | as timestamp/string separator

    The script is provided “AS IS” with no guarantees, no warranties, and confers no rights
        
.LINK
    RBA Consulting http://www.Quisitive.com/

.LINK 
    My Blog Link http://www.texmx.net/
   
#>

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

 