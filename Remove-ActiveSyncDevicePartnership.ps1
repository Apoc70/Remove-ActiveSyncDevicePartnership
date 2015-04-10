<#
    .SYNOPSIS
    Remove Exchange Server 2010 ActiveSync Device Partnerships 
   
   	Sebastian Rubertus / Thomas Stensitzki
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Send comments and remars to: support@granikos.eu
	
	Version 1.0, 2015-04-09
 
    .LINK  
    More information can be found at http://www.rubertus.net/Blog/tabid/85/EntryId/41/Scripted-removing-of-ActiveSync-Device-Partnerships.aspx 
	
    .DESCRIPTION

    THis script removes ActiveSync device association from user mailboxes
    that have been inactive for more than 150 days.

    .NOTES 
    Requirements 
    - Exchange Server 2010
    - Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
    
    .EXAMPLE
    Remove-ActiveSyncDevicePartnership
    	
    #>

### BEGIN SnapIns -------------------------------------------------------------

# Add Exchange SnapIn if not already loaded
if ( (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PsSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
   
    if ( (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue) -eq $null )
    {
        Write-Host "Microsoft.Exchange.Management.PowerShell.Admin could NOT be loaded!" -ForegroundColor Red
        Write-Host "Verify that the Exchange 2010 Management is installed on this computer!" -ForegroundColor Red
    }
}

### END SnapIns ---------------------------------------------------------------

### BEGIN Variables | EDIT ACCORDING TO YOUR NEEDS ----------------------------

# ScriptPath
$scriptPath = "C:\Scripts\Remove-ActiveSync-Devices\"

# Logfile
$logfile = "C:\Scripts\Remove-ActiveSync-Devices\Logs\$(get-date -format yyyy-MM-dd___HH-mm-ss)___Logname.log"

### END Variables -------------------------------------------------------------


### BEGIN Functions -----------------------------------------------------------

Function Log
{
   Param ([string]$logstring)
   Add-content $logfile -value "$(get-date -format yyyy-MM-dd___HH-mm-ss) $logstring "
}

### END Functions -------------------------------------------------------------

### BEGIN Main ----------------------------------------------------------------

# Create a new log file
Write-Host
Write-Host "Script started, creating Log File."
Log "Script started."
Write-Host

# Query User Mailboxes and Device Statistics
Write-Host "Querying User Mailboxes, please wait a few seconds..." -ForeGroundColor green
Log "Querying User Mailboxes."
Write-Host
$Mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -WarningAction SilentlyContinue
$NumberOfMailboxes = $Mailboxes.count
Write-Host "Number of Mailboxes: $NumberOfMailboxes "
Log "Number of Mailboxes: $NumberOfMailboxes "
Write-Host

# Iterate each User Mailbox
ForEach ($Mailbox in $Mailboxes)
{
    $MailboxAlias = $Mailbox.Alias
    Write-Host
    Write-Host "================================================================================="
    Write-Host
    Write-Host "Getting ActiveSync Devices from user $MailboxAlias..."
    Log "Getting ActiveSync Devices from user $MailboxAlias. "
    $AllDevicesFromSpecificUser = Get-ActiveSyncDevice -Mailbox $MailboxAlias -Result Unlimited  -WarningAction SilentlyContinue | Get-ActiveSyncDeviceStatistics -WarningAction SilentlyContinue
    $AllDevicesFromSpecificUserNotSynchronizedSince150Days = Get-ActiveSyncDevice -Mailbox $MailboxAlias -Result Unlimited  -WarningAction SilentlyContinue | Get-ActiveSyncDeviceStatistics  -WarningAction SilentlyContinue | Where {$_.LastSuccessSync -le (Get-Date).AddDays("-150")}
    Write-Host
    $CountAllDevicesFromSpecificUser = $AllDevicesFromSpecificUser.Count
    $CountAllDevicesFromSpecificUserNotSynchronizedSince150Days = $AllDevicesFromSpecificUserNotSynchronizedSince150Days.Count
   
    If ($CountAllDevicesFromSpecificUser -lt 5)
    {
        Write-Host "User $MailboxAlias has only $CountAllDevicesFromSpecificUser ActiveSync Devices. Nothing to delete!" -ForegroundColor Green
        Log "User $MailboxAlias has only $CountAllDevicesFromSpecificUser ActiveSync Devices. Nothing to delete!"
    }
   
    If (($CountAllDevicesFromSpecificUser -gt 4) -and ($CountAllDevicesFromSpecificUserNotSynchronizedSince150Days -gt 1))
    {
        Write-Host "User $MailboxAlias has $CountAllDevicesFromSpecificUser devices. $CountAllDevicesFromSpecificUserNotSynchronizedSince150Days have not synced for more than 150 days." -ForegroundColor Red
        Log "User $MailboxAlias has $CountAllDevicesFromSpecificUser devices. $CountAllDevicesFromSpecificUserNotSynchronizedSince150Days have not synced for more than 150 days."
       
        ForEach ($Device in $AllDevicesFromSpecificUserNotSynchronizedSince150Days)
        {
            $DeviceType = $Device.DeviceType
            $DeviceFriendly = $Device.FriendlyName
            $DeviceID = $Device.DeviceID
            $DeviceFirstSyncTime = $Device.FirstSyncTime
            $DeviceLastSuccessSync = $Device.LastSuccessSync
            Write-Host
            Write-Host "ActiveSync Device 2 delete Properties: "
            Write-Host "-------------------------------------- "
            Write-Host "Type         : $DeviceType "           
            Write-Host "Friendly Name: $DeviceFriendly "
            Write-Host "ID           : $DeviceID "
            Write-Host "Last Sync    : $DeviceLastSuccessSync " -ForegroundColor Red
            Log "Removing Device $DeviceType with ID $DeviceID ..."
            Write-Host
            Write-Host "Removing Device $DeviceID ..." -ForegroundColor Red
            $Device | Remove-ActiveSyncDevice -WarningAction SilentlyContinue
        }
    }
}

# Script finished
Write-Host
Write-Host "Script finished!"
Write-Host
Log "Script finished!"

### END Main ------------------------------------------------------------------