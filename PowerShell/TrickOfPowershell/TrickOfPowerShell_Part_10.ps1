# Reference site: http://powershell.com/cs/blogs/tips/
###########################################################################################################################################################################################

# Tip 1: Use a Lock Screen


# With WPF, PowerShell can create windows in just a couple of lines of code. Here's a funny example of a transparent screen overlay. 

# You can call Lock-Screen and submit a script block and a title. 
# PowerShell will then lock the screen with its overlay, execute the code and remove the lock screen again.

function Lock-Screen([ScriptBlock]$Payload = {Start-Sleep -Seconds 5}, $Title = "Busy, go away ...")
{
    try
    {
        $window = New-Object System.Windows.Window
        
        $label = New-Object Windows.Controls.Label
        $label.Content = $Title
        $label.FontSize = 60
        $label.FontFamily = "Consolas"
        $label.Background = "Transparent"
        $label.Foreground = "Red"
        $label.HorizontalAlignment = "Center"
        $label.VerticalAlignment = "Center"

        $window.AllowsTransparency = $true
        $window.Opacity = .7
        $window.WindowStyle = "None"
        $window.Content = $label
        $window.Left = $window.Top = 0
        $window.WindowState = "Maximized"
        $window.Topmost = $true

        $null = $window.Show()
        
        Invoke-Command -ScriptBlock $Payload
    }
    finally
    {
        $window.Close()
    }
}

$job = {

    Get-ChildItem C:\Windows -Recurse -ErrorAction SilentlyContinue
}

Lock-Screen -Payload $job -Title "I'm busy, go away and grab a coffee ..."

# As you will soon discover, the look screen does protect against mouse clicks, but it won't shield the keyboard. It's just a fun technique, no security lock.

###########################################################################################################################################################################################

# Tip 2: Setting Default Email Address for AD Users


# Scripting Active Directory does not necessarily require additional modules. With simple .NET Framework methods, you can achieve amazing things. 
# In fact, this technique is so powerful that you should not run the following example in your productive environment until you understand what it does.

# The next piece of code finds all users in your Active Directory that are located in CN=Users and have no mail address. 
# It then assigns a default mail address, consisting of first and last name plus "mycompany.com".

$SearchRoot = 'LDAP://CN=Users,{0}' -f ([ADSI]'').distinguishedName.ToString()

# adjust LDAPFilter. Example: (!mail=*) = all users with no defined mail attribute
$LdapFilter = "(&(objectClass=user)(objectCategory=person)(!mail=*))"

$Searcher = New-Object DirectoryServices.DirectorySearcher($SearchRoot, $LdapFilter)
$Searcher.PageSize = 1000

$Searcher.FindAll() | ForEach-Object {

    $User = $_.GetDirectoryEntry()

    try
    {
        $User.Put("mail", ("{0}.{1}@mycompany.com" -f $user.givenName.ToString(), $user.sn.ToString()))
        $User.SetInfo()
    }
    catch
    {
        Write-Warning "Problems with $user. Reason: $_"
    }
}

# This example code can read and change/set any attribute. This is especially useful for custom attributes that often cannot be directly set by cmdlets.

###########################################################################################################################################################################################

# Tip 3: Finding Disk Partition Details


Get-WmiObject -Class Win32_DiskPartition | Select-Object -Property *

# Next, pick the properties you are really interested in, then replace the "*" with a comma-separated list of these properties. For example:

Get-WmiObject -Class Win32_DiskPartition | Select-Object -Property Name,BlockSize,Description,BootPartition
# Output:
#        Name                       BlockSize Description                    BootPartition
#        ----                       --------- -----------                    -------------
#        Disk #0, Partition #0            512 Installable File System                 True
#        Disk #0, Partition #1            512 Installable File System                False
#        Disk #0, Partition #2            512 Installable File System                False

# If you pick four or less properties, the result is a neat table, otherwise a list.

# If you are hungry for more, use the parameter -List to search for other WMI classes, either related to "disk", or just something completely different:

Get-WmiObject -Class Win32_*Processor* -List

###########################################################################################################################################################################################

# Tip 4: Displaying Path Environment Variables


# The environment variable $env:Path lists all paths that are included in the Windows search path when you launch an application. 
# Likewise, $env:PSModulePath lists all paths PowerShell searches for PowerShell modules (and includes in its module auto-loading).

# These variables contain semicolon-separated information. Use the operator -split to view the paths separately:

$env:Path -split ";"

# The third entry (in Program Files) was added by PowerShell 4.0 by the way.

###########################################################################################################################################################################################

# Tip 5: Managing Office365 with PowerShell


# Did you know that you can manage your Office365 accounts with PowerShell, too? Provided you have an Office365 account, try this:

$officeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/poweshell/ `
                -Credential (Get-Credential) -Authentication Basic -AllowRedirection

$import = Import-PSSession $officeSession -Prefix Off365

Get-Command -Noun Off365*

# This will connect to Office 365 with your credentials, and then import the PowerShell cmdlets you can use to manage it. 
# You'll get roughly 400 new cmdlets. If you get an "Access Denied" instead, then your account may not have sufficient privileges, or you mistyped your password.

# You can choose the prefix yourself (see code above) which enables you to connect to multiple Office365 accounts at the same time and using different prefixes. 
# You can omit the prefix, too, when you run Import-PSSession.

# To view the new commands that were exported by Office365, use this:

$import.ExportedCmdlets

###########################################################################################################################################################################################

# Tip 6: Getting Local Group Members


# In PowerShell, local accounts and groups can be managed in an object-oriented way thanks to .NET Framework 3.51 and above. 
# This will list local administrator accounts:

Add-Type -AssemblyName System.DirectoryServices.AccountManagement

$type = New-Object DirectoryServices.AccountManagement.PrincipalContext("Machine", $env:COMPUTERNAME)

$group = [DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($type, "SAMAccountName", "Administrators")


# You can get a lot more, though. Try and query the group by itself:

$group
# Output:
#        IsSecurityGroup       : True
#        GroupScope            : Local
#        Members               : {Administrator, Silence...}
#        Context               : System.DirectoryServices.AccountManagement.PrincipalContext
#        ContextType           : Machine
#        Description           : Administrators have complete and unrestricted access to the computer/domain
#        DisplayName           : 
#        SamAccountName        : Administrators
#        UserPrincipalName     : 
#        Sid                   : S-1-5-32-544
#        Guid                  : 
#        DistinguishedName     : 
#        StructuralObjectClass : 
#        Name                  : Administrators


# Or try and view all properties for all members:

$group.Members

$group.Members | Select-Object -Property SAMAccountName, LastPasswordSet, LastLogon, Enabled

# Output:
#        SamAccountName                 LastPasswordSet           LastLogon                Enabled
#        --------------                 ---------------           ---------                -------
#        Administrator                  2012/10/11 20:59:55       2012/10/11 14:39:10        False
#        Silence                        2012/10/11 06:00:57       2012/10/11 15:41:03         True
#        sihe                           2014/10/15 01:18:09       2014/11/13 01:53:50         True                                                                                                                                   

###########################################################################################################################################################################################

# Tip 7: Searching for Local User Accounts


# Did you know that you can actually search for local user accounts, much like you can search for domain accounts?

Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$type = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalContext('Machine', $env:COMPUTERNAME)
$userPrincipal = New-Object System.DirectoryServices.AccountManagement.UserPrincipal($type)


# One: searches for all local accounts with a name that starts with "S" and are enabled

$userPrincipal.Name = "S*"
$userPrincipal.Enabled = $true

$searcher = New-Object System.DirectoryServices.AccountManagement.PrincipalSearcher
$Searcher.QueryFilter = $userPrincipal
$results = $Searcher.FindAll()
$results | Select-Object -Property Name, LastLogon, Enabled

# Output:
#        Name            LastLogon                    Enabled
#        ----            ---------                    -------
#        Silence         2012/10/11 15:41:03             True



# Two: to find all enabled local accounts with a password that never expires

$userPrincipal.PasswordNeverExpires = $true
$userPrincipal.Enabled = $true

$Searcher = New-Object System.DirectoryServices.AccountManagement.PrincipalSearcher
$Searcher.QueryFilter = $userPrincipal
$results = $Searcher.FindAll()

$results | Select-Object -Property Name, LastLogon, Enabled, PasswordNeverExpires

# Output:
#        Name           LastLogon                    Enabled           PasswordNeverExpires
#        ----           ---------                    -------           --------------------
#        Silence        2012/10/11 15:41:03             True                           True

###########################################################################################################################################################################################

# Tip 8: Managing Windows Defender in Windows 8.1


# Windows 8.1 ships with a new module called "Defender". The cmdlets found inside enable you to manage, 
# view and change all aspects of the Windows Defender anti-spyware application.

Get-Command -Module Defender     # To list all cmdlets available

# If you do not get back anything, then you're probably not running Windows 8.1, so the module is not available.

# Next, try and explore the cmdlets. 

Get-MpPreference            # lists all the current preference settings. Likewise, Set-MpPreference can change them.

Get-MpThreatDetection       # list all threats currently detected (or nothing it there are no current threats).

###########################################################################################################################################################################################

# Tip 9: Search and View PowerShell Videos


# PowerShell is amazing. It will let you search YouTube videos for keywords you select, 
# then offer the videos to you, and upon selection play the videos as well.

# Here's a little script that--Internet access assumed--lists the most recent "Learn PowerShell" videos from YouTube. 
# The list opens in a grid view window, so you can use the full text search at the top or sort columns until you find the video you want to give a try.

# Next, click the video to select it, and then click "OK" in the lower-right corner of the grid. 

# PowerShell will launch your web browser and play the video. Awesome!

$keyword = "Learn PowerShell"

Invoke-RestMethod -Uri "https://gdata.youtube.com/feeds/api/videos?v=2&q=$($keyword.Replace(' ','+'))" | 
    
    Select-Object -Property Title,@{N='Author';E={$_.Author.Name}}, @{N='Link';E={$_.Content.src}}, @{N='Updated';E={[DateTime]$_.Updated}} | 

    Sort-Object -Property Updated -Descending | Out-GridView -Title "Select your '$keyword' video, then click OK to view." -PassThru | ForEach-Object { Start-Process $_.Link }

# Simply change the variable $keyword in the first line to search for different videos or topics.

# Note that due to a bug in PowerShell 3.0, Invoke-RestMethod will only return half of the results. In PowerShell 4.0, this bug was fixed.

###########################################################################################################################################################################################

# Tip 10: Getting Yesterday’s Date - at Midnight!


# Getting relative dates (like yesterday or one week ahead) is easy once you know the Add…() methods every DateTime object supports. This would give you yesterday:

$today = Get-Date
$yesterday = $today.AddDays(-1)
$yesterday                                                         # Output: Thursday, November 20, 2014 09:37:08


# $yesterday will be exactly 24 hours before now. So what if you would like yesterday, but at a given time? Let's say, yesterday at midnight?

$today = Get-Date -Hour 0 -Minute 0 -Second 0 -Millisecond 0
$today                                                             # Output: Friday, November 21, 2014 00:00:00


# And if you want another date with a given time, use Get-Date again to override the parts of the date you need. This is yesterday at midnight:

$today = Get-Date
$yesterday = $today.AddDays(-1)
$yesterday | Get-Date -Hour 0 -Minute 0 -Second 0 -Millisecond 0   # Output: Thursday, November 20, 2014 00:00:00


# According to one's comment: it's slightly wrong above for "Get-Date -Hour 0 -Minute 0 -Second 0" if you're missing "-Millisecond 0" 
# and won't get you exactly midnight unless you are lucky, because you are not settings Milliseconds to be 0.
# In Powershell 3 and above there is Millisecond parameter available on the Get-Date cmdlet which can be used to do that. 
# In PowerShell v2 (Vista or not upgraded Windows 7) there is no such parameter and you have to work around to get exactly midnight. 
# You can for example do: Get-Date -Hour 0 -Minute 0 -Second 0 | foreach { $_.AddMilliseconds(-$_.Millisecond) }


#　Another way to do it:
[System.DateTime]::Today                                           # Output: Friday, November 21, 2014 00:00:00          
[System.DateTime]::Today.Subtract([System.TimeSpan]::FromDays(1))  # Output: Thursday, November 20, 2014 00:00:00

###########################################################################################################################################################################################

# Tip 11: Ordered Hash Tables and Changing Order


# Ordered hash tables are new in PowerShell 3.0 and great for creating new objects. 
# Unlike regular hash tables, ordered hash tables keep the order in which you add keys, 
# so you can control in which order these keys turn into object properties.

$hashTable = [Ordered]@{}
$hashTable.Name = "Silence"
$hashTable.ID = 1
$hashTable.Location = "China"

New-Object -TypeName PSObject -Property $hashTable
# Output:
#        Name           ID               Location                                                              
#        ----           --               --------                                                              
#        Silence         1               China 


# This produces an object with the properties defined in the exact order how they were specified.
# What if you wanted to add another property not at the end, but let's say at the beginning of the list? Try Insert():

$hashTable.Insert(0, "Position", "CSA")
# Output:
#        Name                           Value                                                                                                                                                                                 
#        ----                           -----                                                                                                                                                                                 
#        Position                       CSA                                                                                                                                                                                   
#        Name                           Silence                                                                                                                                                                               
#        ID                             1                                                                                                                                                                                     
#        Location                       China

###########################################################################################################################################################################################

# Tip 12: Getting Error Events from Multiple Event Logs


# Get-EventLog can read events only from one event log at a time. 
# If you want to find events in multiple event logs, you can append array information, though:

$events = @(Get-EventLog -LogName System -EntryType Error)
$events += Get-EventLog -LogName Application -EntryType Error

$events


# In these cases, it might be easier to use WMI in the first place - which can query any number of event logs at the same time.

# This will get you the first 100 error events from the application and system log 
# (cumulated, so if the first 100 errors are in the application log, no system log errors will be reported, of course):

Get-WmiObject -Class Win32_NTLogEvent -Filter 'Type="Error" and (LogFile="System" or LogFile="Application")' | 
    
    Select-Object -First 100 -Property TimeGenerated, LogFile, EventCode, Message

# When you replace Get-WmiObject with Get-CimInstance (which is new in PowerShell 3.0), 
# then the cryptic WMI datetime format is automatically converted to normal date and times:

Get-CimInstance -ClassName Win32_NTLogEvent -Filter 'Type="Error" and (LogFile="System" or LogFile="Application")' | 
    
    Select-Object -First 100 -Property TimeGenerated, LogFile, EventCode, Message

###########################################################################################################################################################################################

# Tip 13: Getting Most Recent Earthquakes


# Everything is connected these days. PowerShell can retrieve public data from web services. 
# So here's a one-liner that gets you a list of the most recently detected earthquakes and their magnitude:

Invoke-RestMethod -Uri "http://www.seismi.org/api/eqs"
# Output:
#        count     earthquakes                                                                                               
#        -----     -----------                                                                                               
#        21740     {@{src=us; eqid=c000is61; timedate=2013-07-29 22:22:48; lat=7.6413; lon=93.6871; magnitude=4.6; depth=4...

Invoke-RestMethod -Uri "http://www.seismi.org/api/eqs" | Select-Object -ExpandProperty Earthquakes -First 30 | Out-GridView

###########################################################################################################################################################################################

# Tip 14: PowerShell Remoting with Large Token Size


# The Kerberos token size depends on the number of group memberships. 
# In some corporate environments with heavy use of group memberships, the token size can grow beyond the limits allowed for PowerShell remoting. 
# In these scenarios, PowerShell remoting fails with a cryptic error message.

# To enable PowerShell remoting, you can set two Registry values and increase the supported token size:

New-ItemProperty HKLM:\SYSTEM\CurrentControlSet\services\HTTP\Parameters -Name "MaxFieldLength" -Value 65335 -PropertyType "DWORD"
New-ItemProperty HKLM:\SYSTEM\CurrentControlSet\services\HTTP\Parameters -Name "MaxRequestBytes" -Value 40000 -PropertyType "DWORD"

# Reference site: http://www.miru.ch/how-the-kerberos-token-size-can-affect-winrm-and-other-kerberos-based-services/

###########################################################################################################################################################################################

# Tip 15: Lowering PowerShell Process Priority


# When you run a PowerShell task, by default it has normal priority, 
# and if the things your script does are CPU intensive, the overall performance of your machine may be affected.

# To prevent this, you can assign your PowerShell process a lower priority and have it run only when CPU load allows. 
# This will ensure that your PowerShell task won't degrade performance of other tasks.

# This sample sets the priority to "below normal". 
# You can also set it to "Idle" in which case your PowerShell script would only run when the machine has nothing else to do:

$process = Get-Process -Id $pid
$process.PriorityClass = "BelowNormal"

<#
     The allowed PriorityClass names are:
     
     Normal, Idle, High, RealTime, BelowNormal, AboveNormal
     
     The scary, but recommended way to find this type of information, 
     that is not easy to find documentation about, is to just try something you know is absolutely wrong:
     
     $process.PriorityClass = 'trudellic'
     
     Exception setting "PriorityClass": "Cannot convert value "trudellic" to type "System.Diagnostics.ProcessPriorityClass". 
     Error: "Unable to match the identifier name trudellic to a valid enumerator name.  
     
     Specify one of the following enumerator names and try again: Normal, Idle, High, RealTime, BelowNormal, AboveNormal""
#>

###########################################################################################################################################################################################

# Tip 16: Using ICACLS to Secure Folders


# Console applications are equal citizens in the PowerShell ecosystem. 
# In this example, a function uses icacls.exe to secure a newly created folder:

function New-Folder
{
    param($Path, $UserName)

    if((Test-Path -Path $Path) -eq $false)
    {
        New-Item -ItemType Directory | Out-Null
    }

    icacls $Path /inheritance:r /grant '*S-1-5-32-544:(OI)(CI)R' ('{0}:(OI)(CI)F' -f $UserName)
}

# The function New-Folder will create a new folder (if it does not exist) and then use icacls.exe to
# turn off inheritance and grant read permissions to the Administrators group and full permissions to the user specified.

###########################################################################################################################################################################################

# Tip 17: Starting Services Remotely


# Since Start-Service has no -ComputerName parameter, you cannot use it easily to remotely start a service. 
# While you could run Start-Service within a PowerShell remoting session, an easier way may sometimes be Set-Service. 

# This would start the Spooler service on Server12:

# Method One:
Set-Service -Name Spooler -Status Running -ComputerName Server12


# Method Two:
Start-Service -InputObject $(Get-Service -ComputerName $Server12 -Name Spooler)

# if you want to start more process on remote machine one by one, then you can try like following:
do
{
    Start-Sleep -Milliseconds 650

}until((Get-Service -ComputerName $Server12 -Name Spooler).Status -eq "Running")

# Note: Unfortunately, there is no -Force switch. So while you can easily start services, you may not be able to stop them this way. 
# Once a service is dependent upon another one, it may not stop without "force".


# By using the -InputObject for Stop-Service you should be able to use the -Force Parameter.

Stop-Service -InputObject (Get-Service -ComputerName $Server12 -Name Spooler) -Force

###########################################################################################################################################################################################

# Tip 18: Getting System Information


# PowerShell plays friendly with existing console applications.
# One of the most useful is systeminfo.exe which gathers all kinds of useful system information. 
# By importing the information provided by systeminfo.exe as CSV, PowerShell can convert the text information into rich objects:

$header = 'Hostname','OSName','OSVersion','OSManufacturer','OSConfig','Buildtype',`
          'RegisteredOwner','RegisteredOrganization','ProductID','InstallDate','StartTime','Manufacturer',`
          'Model','Type','Processor','BIOSVersion','WindowsFolder','SystemFolder','StartDevice','Culture',`
          'UICulture','TimeZone','PhysicalMemory','AvailablePhysicalMemory','MaxVirtualMemory',`
          'AvailableVirtualMemory','UsedVirtualMemory','PagingFile','Domain','LogonServer','Hotfix','NetworkAdapter'

systeminfo.exe /FO CSV | Select-Object -Skip 1 | ConvertFrom-Csv -Header $header

# When you run this, it will take a couple of seconds for systeminfo.exe to gather the information. Then, you get a wealth of information

# Note $header: This variable defines the property names, and your list of headers is exchanged with the default headers. 
# This way, the headers are always the same, regardless of the language a system is using.


# You can also store the information in a variable and access the information individually:

$results = systeminfo.exe /FO CSV | Select-Object -Skip 1 | ConvertFrom-Csv -Header $header

$results.AvailablePhysicalMemory    # Output: 1,186 MB
$results.BIOSVersion                # Output: Hewlett-Packard 786G6 v01.03, 2009/08/25
$results.Hostname                   # Output: SIHE-01
$results.Buildtype                  # Output: Multiprocessor Free
$results.Processor                  # Output: 1 Processor(s) Installed.,[01]: AMD64 Family 16 Model 4 Stepping 2 AuthenticAMD ~2800 Mhz


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Getting System Information for Remote Systems


# systeminfo.exe has built-in remoting capabilities, so provided you have the proper permissions, you can also get system information from remote systems.

function Get-SystemInfo
{
    param($ComputerName = $env:COMPUTERNAME)

    $header = 'Hostname','OSName','OSVersion','OSManufacturer','OSConfig','Buildtype',`
          'RegisteredOwner','RegisteredOrganization','ProductID','InstallDate','StartTime','Manufacturer',`
          'Model','Type','Processor','BIOSVersion','WindowsFolder','SystemFolder','StartDevice','Culture',`
          'UICulture','TimeZone','PhysicalMemory','AvailablePhysicalMemory','MaxVirtualMemory',`
          'AvailableVirtualMemory','UsedVirtualMemory','PagingFile','Domain','LogonServer','Hotfix','NetworkAdapter'

    systeminfo.exe /FO CSV /S $ComputerName | Select-Object -Skip 1 | ConvertFrom-Csv -Header $header
}

$results = Get-SystemInfo -ComputerName Server002 
$results.OSManufacturer
$results.OSVersion

###########################################################################################################################################################################################

# Tip 19: Change Desktop Wallpaper


# To change the current desktop wallpaper and make this change effective immediately, 
# PowerShell can tap into the Windows API calls. Here is a function that changes the wallpaper immediately:

function Set-Wallpaper
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Path,

        [ValidateSet("Center", "Stretch")]
        $Style = "Stretch"
    )

    $code = @"

     using System;
     using System.Runtime.InteropServices;
     using Microsoft.Win32;
     
     namespace Wallpaper
     {
         public enum Style : int
         {
             Center, Stretch
         }
     
         public class Setter
         {
             public const int SetDesktopWallpaper = 20;
             public const int UpdateIniFile = 0x01;
             public const int SendWinIniChange = 0x02;
     
             [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
             private static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
     
             public static void SetWallpaper(string path, Wallpaper.Style style)
             {
                 SystemParametersInfo(SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange);
     
                 RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true);
     
                 switch(style)
                 {
                     case Style.Stretch :
                         key.SetValue(@"WallpaperStyle", "2") ; 
                         key.SetValue(@"TileWallpaper", "0") ;
                         break;
     
                     case Style.Center :
                         key.SetValue(@"WallpaperStyle", "1") ; 
                         key.SetValue(@"TileWallpaper", "0") ; 
                         break;
                 }
     
                 key.Close();
             }
         }
     }  
"@

    Add-Type -TypeDefinition $code

    [Wallpaper.Setter]::SetWallpaper($Path, $Style)
}

Set-Wallpaper -Path 'C:\Windows\Web\Wallpaper\Windows\img0.jpg'

# Question: can this be used to change a wallpaper on a remote machine?

# Answer: I am afraid not. It is changing the wallpaper for the account that runs the script, so it is the equivalent of using the 'personalize' context menu.

###########################################################################################################################################################################################

# Tip 20: Finding Logon Failures


# Whenever someone logs on with invalid credentials, there will be a log entry in the security log.

# Here is a function that can read these events from the security log (Admin privileges needed). 
# It will then list all the invalid logons found in the log:

function Get-LogonFailure
{
    param($Computer)

    try
    {
        Get-EventLog -LogName Security -EntryType FailureAudit -InstanceId 4625 -ErrorAction Stop @PSBoundParameters | ForEach-Object {
        
            $domain
            $user = $_.ReplacementStrings[5,6]
            $time = $_.TimeGenerated

            "Logon Failure: $domain\$user at $time"
        }
    }
    catch
    {
        if($_.CategoryInfo.Category -eq "ObjectNotFound")
        {
            Write-Host "No logon failures found." -ForegroundColor Green
        }
        else
        {
            Write-Warning "Error occured: $_"
        }
    }
}

# Note that this function can work remotely, too. Use the -ComputerName parameter to query a remote system. 
# The remote system needs the running RemoteRegistry service, and you need local administrator privileges on the target machine.

###########################################################################################################################################################################################

# Tip 21: Finding Logged-On User


# There is a helpful console application called quser.exe which will tell you who is logged on to a machine. 
# The executable returns plain text, but with the help of a little regular expression, this text can be converted to CSV and then imported into PowerShell.

# So this will give you rich objects for all users currently logged on to your machine:

(quser.exe) -replace "\s{2,}","," | ConvertFrom-Csv

# Output:
#         USERNAME    : >sihe
#         SESSIONNAME : console
#         ID          : 1
#         STATE       : Active
#         IDLE TIME   : none
#         LOGON TIME  : 2014/11/19 12:05 PM

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Finding Logged-On User on Remote Machine


# Here is now a function that also allows us to find the currently logged-on user on a remote machine.
# As an extra benefit, the returned information is appended by a property called "ComputerName", 
# so when you query multiple machines, you will know where the results came from:

function Get-LoggedOnUser
{
    param([string[]]$ComputerName = $env:COMPUTERNAME)

    $ComputerName | ForEach-Object{
        
        (quser.exe /SERVER:$_) -replace "\s{2,}","," | ConvertFrom-Csv | Add-Member -MemberType NoteProperty -Name ComputerName -Value $_ -PassThru
    }
}

Get-LoggedOnUser -ComputerName Server002

# Output:
#        USERNAME     : administrator
#        SESSIONNAME  : console
#        ID           : 1
#        STATE        : Active
#        IDLE TIME    : none
#        LOGON TIME   : 2014/11/19 07:18 PM
#        ComputerName : Server002
#        
#        USERNAME     : sihe
#        SESSIONNAME  : rdp-tcp#0
#        ID           : 2
#        STATE        : Active
#        IDLE TIME    : 1:30
#        LOGON TIME   : 2014/11/20 09:39 AM
#        ComputerName : Server002

Get-LoggedOnUser -ComputerName IIS-CTI5052 | Select-Object -Property UserName, SessionName, "Logon Time", "Idle Time"

###########################################################################################################################################################################################

# Tip 22: Hidden Array Extensions in PowerShell 4.0


@(1..10).Where({$_ % 2})                          # will get only uneven numbers from a list of numbers

@(Get-Service).Where({$_.Status -eq "Running"})   # will get only services that are actually running

# And there is more (undocumented) stuff. This line will get all numbers that are greater than 2, but only the first 4 of them:

@(1..10).Where({$_ -gt 2}, "skipuntil", 4)

# Finally, this line will do the same, but then converts them to the TimeSpan objects:

@(1..10).Where({$_ -gt 2}, "skipuntil", 5).Foreach([Timespan])

###########################################################################################################################################################################################

# Tip 23: Eliminating Empty Results


# To exclude results that have empty properties, you can easily use Where-Object. 
# For example, when you run Get-Hotfix, and you only want to see hotfixes that have a date for InstalledOn, here is the solution:
Get-HotFix | Where-Object InstalledOn

# Likewise, to get only network adapters from WMI that actually have an IP address, try this:
Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object IPAddress

# Note that in PowerShell 2.0 and below, you need to use the full syntax like this:
Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress }

# Where-Object will exclude any results that have either one of the following in the property you selected: a null value, an empty string, or the number 0.
#  This is because all of these will convert to $false when converted into a Boolean value.

###########################################################################################################################################################################################

# Tip 24: Unblocking Download Files


# Any file you download from the Internet or receive via email get marked by Windows as potentially unsafe. 
# If the file contains executables or binaries, they will not run until you unblock the file. 

# PowerShell 3.0 and better can identify files with a "download mark":

Get-ChildItem -Path $HOME\Downloads -Recurse | Get-Item -Stream Zone.Identifier -ErrorAction Ignore | Select-Object -ExpandProperty FileName | Get-Item

# Output:
#        Mode                LastWriteTime     Length Name                                                                                                                                                                    
#        ----                -------------     ------ ----                                                                                                                                                                    
#        -a---        2014/01/03  02:30 PM     311887 SAM_1557.JPG

# You may not receive any files (if there are none with a download mark), or you may get tons of files 
# (which can be an indication that you unpacked a downloaded ZIP file and forgot to unblock it before).

# To remove the blocking, use the Unblock-File cmdlet. 
# This would unblock all files in your download folder that are currently blocked (not touching any other files): 

Get-ChildItem -Path $HOME\Downloads -Recurse | Get-Item -Stream Zone.Identifier -ErrorAction Ignore | Select-Object -ExpandProperty FileName | Get-Item | Unblock-File

###########################################################################################################################################################################################

# Tip 25: Create New Local Admin Account on the Fly


# Ever needed a new local administrator account for testing purposes? Provided you are already Administrator, 
# and you opened a PowerShell with full Administrator privileges, adding such a user is a matter of just a couple of lines of code:

$object = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')

$Administrators = $object.Translate( [System.Security.Principal.NTAccount]).Value.Split('\')[1]

net localgroup $Administrators domain\username /add

# Note that this way try to translate "Administrators" to local culture via sid, so it can be create success on non-us OS.

###########################################################################################################################################################################################

# Tip 26: Reading and Writing NTFS Streams


# When a file is stored on a drive with NTFS file system, you can attach data streams to it to store hidden information.

# Here is a sample that hides PowerShell code in an NTFS stream of a script. 
# When you run this code, it creates a new PowerShell script file on your desktop, and opens the file in the ISE editor:

$path = "$home\Desktop\secret.ps1"

$secretCode = {

    Write-Host -ForegroundColor Red "This is a miracle!"

    [System.Console]::Beep(4000, 1000)
}

Set-Content -Path $path -Value '(Invoke-Expression ''[ScriptBlock]::Create((Get-Content ($MyInvocation.MyCommand.Definition) -Stream SecretStream))'').Invoke()'
Set-Content -Path $path -Stream SecretStream -Value $secretCode
ise $path

# The new file will expose code like this:
(Invoke-Expression '[ScriptBlock]::Create((Get-Content ($MyInvocation.MyCommand.Definition) -Stream SecretStream))').Invoke()

# When you run the script file, it will output a red text and beeps for a second. 
# So the newly created script actually executes the code embedded into the secret NTFS stream "SecretStream".

# To attach hidden information to (any) file stored on an NTFS volume, use Add-Content or Set-Content with the -Stream parameter. 

# To read hidden information from a stream, use Get-Content and again specify the -Stream parameter with the name of the stream used to store the data. 

###########################################################################################################################################################################################

# Tip 27: Getting DNS IP Address from Host Name


# There is a tiny .NET function called GetHostByName() that is vastly useful. It will look up a host name and return its current IP address:

[System.Net.DNS]::GetHostByName("IIS-CTI5052")

# Output:
#        HostName                                      Aliases                             AddressList                                                           
#        --------                                      -------                             -----------                                                           
#        machine.redmond.corp.company.com             {machine.corp.company.com}           {11.99.15.58} 



# With just a simple PowerShell wrapper, this is turned into a great little function that is extremely versatile:

function Get-IPAddress
{
    param
    (
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [String[]]
        $Name
    )

    process
    {
        $Name | ForEach-Object {
        
            try
            {
                [System.Net.Dns]::GetHostByName($_)
            }
            catch
            {
               $_
            }
        }
    }
}

# You can now run the function as-is (to get your own IP address). 
# You can submit one or more computer names (comma separated). You can even pipe in data from Get-ADComputer or Get-QADComputer.

Get-IPAddress
Get-IPAddress -Name TobiasAir1
Get-IPAddress -Name TobiasAir1, Server12, Storage1

'TobiasAir1', 'Server12', 'Storage1' | Get-IPAddress

Get-QADComputer | Get-IPAddress
Get-ADComputer -Filter * | Get-IPAddress 

# This is possible because the function has both a pipeline binding and an argument serializer.

# The -Name argument is fed to ForEach-Object, so no matter how many computer names a user specifies, they all get processed.

# The -Name parameter accepts value from the pipeline both by property and as a value. 
# So you can feed in any object that has a "Name" property, but you can also feed in any list of plain strings.

# If you submit a computer name that cannot be resolved, exception happens. Add code to the catch block if you want an error message instead.

###########################################################################################################################################################################################

# Tip 28: Playing a Sound on Error


# To catch a user’s attention, your script can easily play WAV sound files. Here is a simple function:

function Play-Alarm
{
    $path = "$PSScriptRoot\alarm06.wav"

    $playerStart = New-Object Media.SoundPlayer $path
    $palyerStart.Load()
    $palyerStart.PlaySync()
}

# It assumes the WAV file is located in the same folder the script is saved in. Note that $PSScriptRoot is not supported in PowerShell 2.0.

# Just make sure you rename the $path variable and have it point to any valid WAV file you want to use.

# By default, PowerShell will wait until the sound is played. Replace PlaySync() with Play() if you want PowerShell to continue and not wait.



# Alternatively we can use below sound:

[System.Media.SystemSounds]::Hand.Play()
[System.Media.SystemSounds]::Beep.Play()
[System.Media.SystemSounds]::Asterisk.Play()
[System.Media.SystemSounds]::Question.Play()
[System.Media.SystemSounds]::Exclamation.Play()

###########################################################################################################################################################################################

# Tip 29: Pinging Computers


# There are multiple ways how you can ping computers. 
# Here is a simple approach that uses the traditional ping.exe but can be easily integrated into your scripts:

function Test-Ping
{
    param([Parameter(ValueFromPipeline = $true)]$Name)

    process
    {
        $null = ping.exe $Name -n 1 -w 1000

        if($LASTEXITCODE -eq 0)
        {
            $Name
        }
    }
}

# Test-Ping accepts a computer name or IP address and will return it if the ping was successful. 
# This way, you can feed in large lists of computer names or IP addresses, and get back only those that were online:

'??','127.0.0.1','localhost','notthere',$env:COMPUTERNAME | Test-Ping 


# It's also possible in .NET. Try:
(New-Object System.Net.NetworkInformation.Ping).Send("localhost")

###########################################################################################################################################################################################

# Tip 30: Multiple Assignments in One Line


$a = Get-Service

($a = Get-Service)
# Output:
#        Status   Name               DisplayName                           
#        ------   ----               -----------                           
#        Running  AdobeARMservice    Adobe Acrobat Update Service          
#        Stopped  AdobeFlashPlaye... Adobe Flash Player Update Service     
#        Stopped  AeLookupSvc        Application Experience                
#        Stopped  ALG                Application Layer Gateway Service 


$b = ($a = Get-Service).Name

$a
# Output:
#        Status   Name               DisplayName                           
#        ------   ----               -----------                           
#        Running  AdobeARMservice    Adobe Acrobat Update Service          
#        Stopped  AdobeFlashPlaye... Adobe Flash Player Update Service     
#        Stopped  AeLookupSvc        Application Experience                
#        Stopped  ALG  

$b
# Output:
#        AdobeARMservice
#        AdobeFlashPlayerUpdateSvc
#        AeLookupSvc
#        ALG

$c = ($b = ($a = Get-Service).Name).ToUpper()     
# Now $c will also contain the service names in all uppercase letters. Pretty freaky stuff

###########################################################################################################################################################################################

# Tip 31: Speaking English and German (and Spanish, and you name it)


# Windows 8 is the first operating system that comes with fully localized text-to-speech engines. 
# So you can now have PowerShell speak (and curse) in your mother tongue.

# At the same time, there is always an English engine, so your computer is now bilingual.

# Here is a sample script for German systems (that is easily adaptable to your locale). 
# Simply change the language IDs (like "de-de" for German), and have Windows speak in another language.

# Note: Before Windows 8, only English engines shipped. With Windows 8, you will also get your local language. No other languages.

$speaker = New-Object -ComObject SAPI.SpVoice

$speaker.Voice = $speaker.GetVoices() | Where-Object { $_.ID -like "*de-de*" }
$null = $speaker.Speak('Ich spreche Deutsch')

$speaker.Voice = $speaker.GetVoices() | Where-Object { $_.ID -like "*en-us*" }
$speaker.Speak("But I can of course also speak english.")

###########################################################################################################################################################################################

# Tip 32: Testing for Valid Date


# If you need to test whether some information resembles a valid date format, here is a test function:

function Test-Date
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Date
    )

    (($Date -as [DateTime]) -ne $null)
}

# It uses the -as operator to try and convert the input to a DateTime format. 
# If this fails, the result is $null, so the function checks for this and returns $true or $false. 
# Note that the -as operator uses your local DateTime format.

###########################################################################################################################################################################################

# Tip 33: Using Comma as Decimal Delimiter


# You may not be aware of this, but PowerShell uses a different decimal delimiter for input and output - which may cause confusions to script users.

# When you enter information, PowerShell expects culture-neutral format (using "." as decimal delimiter). 
# When outputting information, it uses your regional setting (so in many countries, a "," is used).

$a = 1.5
$a                # Output: 1.5

# This is good practice, because by using culture-neutral input format, scripts will always run the same, regardless of the culture settings. 
# However, if you want users to be able to use a comma as delimiter, take a look at this script:

function Multiply-LocalNumber
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Number1,

        $Number2 = 10
    )

    [Double]$Number1 = ($Number1 -join ".")
    [Double]$Number2 = ($Number2 -join ".")

    $Number1 * $Number2
}

Multiply-LocalNumber -Number1 1.5 -Number2 9.233    # Output: 13.8495
Multiply-LocalNumber -Number1 1,5 -Number2 9,233    # Output: 13.8495

# When a user picks the comma, PowerShell actually interprets this as array. 
# That’s why the script joins any array by a ".", effectively converting an array to a number. 
# Since the result of -join is a string, the string needs to be converted to a number, and all is fine.

# Of course, this is a hacky trick, and it is always better to educate your users to always use the "." delimiter in the first place.

###########################################################################################################################################################################################

# Tip 34: Mandatory Parameter with a Dialog


# Typically, when you mark a function parameter as "mandatory", PowerShell will prompt the user when the user omits the parameter:

function Get-Something
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Path
    )

    "You entered $Path "
}

Get-Something
# Output:
#        cmdlet Get-Something at command pipeline position 1
#        Supply values for the following parameters:
#        Path: 



# Here is an alternative: if the user omits -Path, the function opens an OpenFile dialog:

function Get-Something
{
    param
    (
        $Path = $(
            
            Add-Type -AssemblyName System.Windows.Forms

            $dialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog

            if($dialog.ShowDialog() -eq "OK")
            {
                $dialog.FileName
            }
            else
            {
                Throw 'No Path submitted'
            }        
        )
    )

    "You entered $Path"
}

Get-Something

###########################################################################################################################################################################################

# Tip 35: Setting (And Deleting) Environment Variables


# PowerShell can read environment variables easily. This returns the current windows folder:
$env:windir

# However, if you want to make permanent changes to user or machine environment variables, you need to access .NET functionality. 
# Here is a simple function that makes setting or deleting environment variables a snap:

function Set-EnvironmentVariable
{
    param
    (
        [Parameter(Mandatory = $true, HelpMessage = "Help note")]
        $Name,

        [System.EnvironmentVariableTarget]
        $Target,

        $Value = $null
    )

    [System.Environment]::SetEnvironmentVariable($Name, $Value, $Target)
}

Set-EnvironmentVariable -Name TestVar -Value 123 -Target User         # To create a permanent user environment variable

# Note that new user variables are visible only to newly launched applications. 
# Applications that were already running will keep their copied process set unless they explicitly ask for changed variables

Set-EnvironmentVariable -Name TestVar -Value "" -Target User          # deletes the variable

###########################################################################################################################################################################################

# Tip 36: Reading StringExpand Registry Values


$path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion'
$key = Get-ItemProperty -Path $path
$key.DevicePath                                                    # Output: C:\Windows\inf


# The result will be an actual path. 
# That is OK unless you wanted to get the original (unexpanded) Registry value. Here is the example that reads the original value:

$path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion'
$key = Get-Item -Path $path 
$key.GetValue('DevicePath', '', 'DoNotExpandEnvironmentNames')     # Output: %SystemRoot%\inf


# Accessing Registry values this way will get you another valuable piece of information: you can now also find out the value's data type:

$path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion'
$key = Get-Item -Path $path 
$key.GetValueKind("DevicePath")                                    # Output: ExpandString

###########################################################################################################################################################################################

# Tip 37: Importing Certificates from PFX Files


# You can use Get-PfxCertificate to load digital certificates from PFX files, and then use the certificate to sign script files, for example:

$pfxPath = 'C:\PathToPfxFile\testcert.pfx'
$cert = Get-PfxCertificate -FilePath $pfxPath
$cert

Get-ChildItem -Path C:\MyScripts -Filter *.ps1 | Set-AuthenticodeSignature -Certificate $cert


# However, Get-PfxCertificate will interactively ask for the password that you specified when you originally exported the certificate to the PFX file.

# To import a certificate unattended, try this code instead:

$pfxPath = 'C:\PathToPfxFile\testcert.pfx'
$password = "topsecret"

Add-Type -AssemblyName System.Security
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import($pfxPath, $password, "Exporttable")
$cert

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Importing Multiple Certificates from PFX Files


# Get-PfxCertificate can import digital certificates from a PFX file. However, it can only retrieve one certificate. 
# So if your PFX file contains more than one certificate, you cannot use this cmdlet to get the others.

# To import multiple certificates from a PFX file, simply use the code below:

$pfxPath = 'C:\PathToPfxFile\testcert.pfx'
$password = "topsecret"

Add-Type -AssemblyName System.Security
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
$cert.Import($pfxPath, $password, "Exportable")
$cert


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Signing VBScript Files with PowerShell


# You probably know that Set-AuthenticodeSignature can be used to digitally sign PowerShell scripts. 
# But did you know that this cmdlet can sign anything that 

# This piece of code would load a digital certificate from a PFX file, 
# then scan your home folders for VBScript files, and apply a digital signature to the scripts:

# change path to point to your PFX file:
$pfxpath = 'C:\Users\Tobias\Documents\PowerShell\testcert.pfx'

# change password to the password needed to read the PFX file:
# (this password was set when you exported the certificate to a PFX file)
$password = 'topsecret'

Add-Type -AssemblyName System.Security
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import($pfxpath, $password, "Exportable")

Get-ChildItem -Path $HOME -Filter *.vbs -Recurse -ErrorAction SilentlyContinue | Set-AuthenticodeSignature -Certificate $cert -WhatIf

###########################################################################################################################################################################################

# Tip 38: Ensuring Backward Compatibility


# Let's assume you have created this function:

function Test-Function
{
    param([Parameter(Mandatory = $true)]$ServerPath)

    "You selected $ServerPath"
}

# It works fine, but half a year later in a code review, your boss wants you to use standard parameter names, 
# and rename "ServerPath" to "ComputerName". So you change your function appropriately:

function Test-Function
{
    param([Parameter(Mandatory = $true)]$ComputerName)

    "You selected $ComputerName"
}

# What you cannot easily control, though, is who else calls your function, and may still use the old parameter. 
# So to ensure backward compatibility, make sure your function can still work with the old parameter name, too:

function Test-Function
{
    param
    (
        [Parameter(Mandatory = $true)]
        [Alias("ServerPath", "CN")]     # we can have more than one alias for a single parameter
        $ComputerName
    )
   
    "You selected $ComputerName"
}

Test-Function -ServerPath Server1       # Output: You selected Server1
Test-Function -ComputerName Server1     # Output: You selected Server1
Test-Function -CN Server1               # Output: You selected Server1

# Old code can now still run, and new code (and code completion) will use the new name:

###########################################################################################################################################################################################

# Tip 39: Using Fully Qualified Names in Remoting


# When you try PowerShell Remoting, you may run into connection errors just because your machine name was not fully qualified. 
# Kerberos authentication may or may not needs this, depending on your DNS configuration.

Enter-PSSession -ComputerName Storage1

# When that happens, try and ask DNS for the fully qualified name:

[System.Net.Dns]::GetHostByName("IIS-CTI5052").HostName     # Output: Storage1.domain.corp.company.com

# Then, try this name instead. If remoting is enabled and set up correctly, you should now be able to connect.

###########################################################################################################################################################################################

# Tip 40: Use Fresh Testing Environment in PowerShell ISE


<#
    When you develop PowerShell scripts in the PowerShell ISE editor, you should run the final tests in a clean environment,
    making sure that no leftover variables or functions from previous runs interfere.
    
    The easiest way to create a completely fresh testing environment is this: choose the menu File, then "New PowerShell Tab". 
    This gets you a new tab, and this tab actually represents a completely new PowerShell host. Perfect for testing!
#>

# Note: shortcuts here should be "Ctrl + T" but not "Ctrl + N".

###########################################################################################################################################################################################

# Tip 41: Use $PSScriptRoot to Load Resources


# Beginning in PowerShell 3.0, there is a new automatic variable available called $PSScriptRoot. 
# This variable previously was only available within modules. 
# It always points to the folder the current script is located in (so it only starts to be useful once you actually save a script before you run it).

# You can use $PSScriptRoot to load additional resources relative to your script location. 
# For example, if you decide to place some functions in a separate "library" script that is located in the same folder, 
# this would load the library script and import all of its functions:

# this loads the script "library1.ps1" if it is located in the very
# same folder as this script.
# Requires PowerShell 3.0 or better.

. "$PSScriptRoot\library1.ps1" 

# Likewise, if you would rather want to store your library scripts in a subfolder, 
# try this (assuming the library scripts have been placed in a folder called "resources" that resides in the same folder as your script:

# this loads the script "library1.ps1" if it is located in the subfolder
# "resources" in the folder this script is in.
# Requires PowerShell 3.0 or better.

. "$PSScriptRoot\resources\library1.ps1" 

###########################################################################################################################################################################################

# Tip 42: Keeping a Handle to a Process


# When you launch an EXE file, PowerShell will happily start it, then continue and not care about it anymore:
notepad


# If you'd like to keep a handle to the process, for example to find out its process ID, 
# or to be able to check back later how the process performs, or to kill it, use Start-Process and the –PassThru parameter. This returns a process object:

$process = Start-Process -FilePath notepad -PassThru

$process.Id                        # Output: 7636
$process.CPU                       # Output: 0.1404009
$process.CloseMainWindow()         # Output: True

###########################################################################################################################################################################################

# Tip 43: Filtering Text-Based Command Output


# Comparison operators act like filters when applied to arrays. 
# So any console command that outputs multiple text lines can be used with comparison operators.

@(netstat) -like "*establ*"

@(netstat) -like "*stor*establ*"

@(ipconfig) -like "*IPv4*"         # Output: IPv4 Address. . . . . . . . . . . : 172.18.32.43

# The trick is to enclose the console command in @() which makes sure that the result is always an array.



netstat | Select-String "establ"
ipconfig | Select-String "ipv4"

# Select-String is powerful but entirely different. It does not return strings but instead MatchInfo objects. 
# Pipe the results to Get-Member to discover the differences.

@(ipconfig) -like "*IPv4*" | Get-Member *
ipconfig | Select-String "IPv4" | Get-Member *

###########################################################################################################################################################################################

# Tip 44: Using Aliases to Launch Windows Components


# PowerShell is not just an automation language but also an alternate user interface. 
# If you do not like the graphical interface, educate PowerShell to open the tools you need via easy alias names.

# For example, to open the device manager, you could use its original name:
devmgmt.msc

# If you do not want to remember this name, use an alias:
Set-Alias -Name DeviceManager -Value devmgmt.msc
DeviceManager

# As you can see, to open the device manager, all you now need is enter "DeviceManager". 
# You can also just enter "Device" and press TAB to use auto-completion.

# Aliases will vanish when PowerShell closes, so to keep your aliases, add the Set-Alias command(s) to your profile script. 
# The path can be found in $profile. You may have to create this file (and its parent folders) first if it does not yet exist. 
# Test-Path can check if it is already present or not.

$profile  # Output: C:\Users\sihe\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1

Test-Path $profile   # Output: False

###########################################################################################################################################################################################

# Tip 45: Expanding Variables in Strings


# To insert a variable into a string, you probably know that you can use double quotes like this:

$domain = $env:USERDOMAIN
$username = $env:USERNAME

"$domain\$username"                       # Output: far\sihe
'$domain\$username'                       # Output: $domain\$username

"$username`: located in domain $domain"   # Output: sihe: located in domain far

"Current Background Color: $host.UI.RawUI.BackgroundColor" 
# Output: Current Background Color: System.Management.Automation.Internal.Host.InternalHost.UI.RawUI.BackgroundColor

# Token colors indicate that double quoted strings only resolve the variable and nothing else 
# (nothing that follows the variable name, like accessing object properties).

# To solve this problem, you must use one of these techniques:

"Current Background Color: $($host.UI.RawUI.BackgroundColor)"
'Current Background Color: ' + $host.UI.RawUI.BackgroundColor
'Current Background Color: {0}' -f $host.UI.RawUI.BackgroundColor

###########################################################################################################################################################################################

# Tip 46: Save Time With Select-Object -First!


# Select-Object has a parameter called -First that accepts a number. It will then return only the first x elements. 

# This gets you the first 4 PowerShell scripts in your Windows folder:

Get-ChildItem -Path C:\Windows -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue | Select-Object -First 4

# Beginning in PowerShell 3.0, -First not only selects the specified number of results.
# It also informs the upstream cmdlets in the pipeline that the job is done, effectively stopping the pipeline.

# So if you have a command where you know that after a certain number of results, you are done, 
# then you should always add Select-Object -First x - this can speed up your code dramatically in certain cases.

# Let's assume you are looking for a file called "test.txt" somewhere in your home folder, and let's assume there is only one such file. 
# You just do not know where exactly it is located, so you use Get-ChildItem and -Recurse to recursively search all folders:
Get-ChildItem -Path $HOME -Filter test.txt -Recurse -ErrorAction SilentlyContinue

# When you run this, Get-ChildItem will eventually find your file - and then continue to search your folder tree. 
# Maybe for minutes. It cannot know whether or not there may be additional files.

# You know, though, so if you know the number of expected results beforehand, try this:
Get-ChildItem -Path $HOME -Filter test.txt -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

# This time, Get-ChildItem will stop immediately once the file is found.

###########################################################################################################################################################################################

# Tip 47: Tag Your Objects with Additional Information


# There may be the need to add additional information to command results. 
# Maybe you get data from different machines and want to keep a reference where the data came from. 
# Or, you want to add a date so you know when the data was created.

Get-Service | Add-Member -MemberType NoteProperty -Name SourceComputer -Value $env:COMPUTERNAME -PassThru | 
              Add-Member -MemberType NoteProperty -Name Date -Value (Get-Date) -PassThru | 
              Select-Object -Property Name, Status, SourceComputer, Date
# Output:
#        Name                             Status SourceComputer          Date                                                
#        ----                             ------ --------------          ----                                                
#        AdobeARMservice                 Running SIHE-01                 2014/11/24 15:02:49                                 
#        AdobeFlashPlayerUpdateSvc       Stopped SIHE-01                 2014/11/24 15:02:49                                 
#        AeLookupSvc                     Stopped SIHE-01                 2014/11/24 15:02:49                                 
#        ALG                             Stopped SIHE-01                 2014/11/24 15:02:49  

# Just remember that your added properties may not show up in the result until you use Select-Object and explicitly ask to show them.  

# With PowerShell 3 (and above), there are several more convenient forms, especially when adding multiple properties. For example:
Get-Service | Add-Member -PassThru @{ SourceComputer = $env:COMPUTERNAME; Date = (Get-Date) } | Select-Object -Property Name, Status, SourceComputer, Date

###########################################################################################################################################################################################

# Tip 48: Eliminating Duplicates


# Sort-Object has an awesome feature: with the parameter -Unique, you can remove duplicates:

1,3,2,3,2,1,2,3,4,7 | Sort-Object              # Output: 1,1,2,2,2,3,3,3,4,7
                                               
1,3,2,3,2,1,2,3,4,7 | Sort-Object -Unique      # Output: 1,2,3,4,7

# This can be applied to object results, as well. Check out this example: it will get you the latest 40 errors from your system event log:

Get-EventLog -LogName System -EntryType Error -Newest 40 | Sort-Object -Property InstanceID, Message | Out-GridView

# This may be perfectly fine, but depending on your event log, you may also get duplicate entries.
# With -Unique, you can eliminate duplicates, based on multiple properties:

Get-EventLog -LogName System -EntryType Error -Newest 40 | Sort-Object -Property InstanceID, Message -Unique | Out-GridView

# You will no longer see more than one entry with the same InstanceID AND Message.
# You can then sort the result again, to get back the chronological order:

Get-EventLog -LogName System -EntryType Error -Newest 40 | Sort-Object -Property InstanceID, Message -Unique | Sort-Object -Property TimeWritten -Descending | Out-GridView

# So bottom line is: Sort-Objects parameter -Unique can be applied to multiple properties at once. 

###########################################################################################################################################################################################

# Tip 49: Formatting Numbers Easily


# Often, users need to format numbers and limit the number of digits, or add leading zeros. 
# There is one simple and uniform strategy for this: the operator "-f"!

$number = 68
"{0:d7}" -f $number                  # Output: 0000068

# This will produce a 7-digit number with leading zeros. Adjust the number after "d" to control the number of digits.



# To limit the number of digits, use "n" instead of "d". This time, the number after "n" controls the number of digits:
$number = 35553568.67826738
"{0:n1}" -f $number                  # Output: 35,553,568.7



# Likewise, use "p" to format percentages:
$number = 0.32562176536
"{0:p2}" -f $number                  # Ouput: 32.56 %

###########################################################################################################################################################################################

# Tip 50: Padding Strings Left and Right


$mytext = "Test"

$paddedText = $mytext.PadLeft(15)
"Here is the text: '$paddedText'"    # Output: Here is the text: '           Test'

$paddedText = $mytext.PadRight(15)
"Here is the text: '$paddedText'"    # Output: Here is the text: 'Test           '


# You can even add a padding character yourself (if you do not want to pad with spaces):

"Silence".PadLeft(20, ".")           # Output: .............Silence
"Silence".PadRight(20, "-")          # Output: Silence-------------

###########################################################################################################################################################################################

# Tip 51: Auto-Connecting with Public Hotspot


# Many mobile phone service providers offer public hotspots at airports and public places. 
# To connect, you typically need to browse to a logon page, and then manually enter your credentials.

# Here is a script that does this automatically. 
# It is tailored to t-mobile.de but in the script, you can see the places that need adjustment for other providers:

function Start-Hostspot
{
    param
    (
        [System.String]
        $UserName = "XYZ@t-mobile.de",

        [System.String]
        $Password = "topsecred"
    )

    $url = 'https://hotspot.t-mobile.net/wlan/start.do'
    $r = Invoke-WebRequest -Uri $url -SessionVariable fb

    $form = $r.Forms[0]

    $form.Fields["username"] = $UserName
    $form.Fields["password"] = $Password

    $r = Invoke-WebRequest -Uri ("https://hotspot.t-mobile.net" + $form.Action) -WebSession $fb -Method POST -Body $form.Fields

    Write-Host 'Connected' -ForegroundColor Green

    Start-Process "http://www.google.de"
}

# In a nutshell, Invoke-WebRequest can navigate to a page, fill out form data, and then send the form back. 
# To do this right, you will want to look at the source code of the logon web page 
# (browse to the page, then right-click the page in your browser and display the source HTML code).

# Next, identify the form that you want to fill out, 
# and change the form field names and action according to the form that you identified in the HTML code.

###########################################################################################################################################################################################

# Tip 52: Finding Events around A Date


# Often, you might want to browse all system events around a given date. 
# Let's say a machine crashed at 08:47, and you'd like to see all events +/− 2 minutes around that time. 

$deltaminutes = 2
$delta = New-TimeSpan -Minutes $deltaminutes

$time = Read-Host -Prompt 'Enter time of event (yyyy-MM-dd HH:mm:ss or HH:mm)'

$datetime = Get-Date -Date $time
$start = $datetime - $delta
$end = $datetime + $delta

$result = @(Get-EventLog -LogName System -Before $end -After $start)
$result +=  Get-EventLog -LogName Application -Before $end -After $start

$result | Sort-Object -Property TimeGenerated -Descending | Out-GridView -Title "Events +/− $deltaminutes minutes around $datetime"

# When you run it, it asks for a time or a date and time. 
# Next, you get back all events that occurred within 2 minutes before and after in the system and application log.
# If you do not get back anything, then there were no events in the given time frame.

# The code illustrates how you can get events within a given time frame, and it illustrates how you can query multiple event logs.

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Finding Errors since Yesterday


# Relative dates are important to get data within a special time frame, avoiding hard-coded dates and times.

# This script will get all error and warning events from the System log since yesterday (24 hours ago):

$today = Get-Date
$oneday = New-TimeSpan -Days 1

$yesterday = $today - $oneday

Get-EventLog -LogName System -EntryType Error, Warning -After $yesterday

# you can use the AddDays method as shown in the following command to retrieve events for the last 24 hours.
Get-EventLog -LogName system -EntryType Error, Warning -After (Get-Date).AddDays(-1)

# To retrieve events from yesterday starting at midnight, use the Date method prior to AddDays.
Get-EventLog -LogName system -EntryType Error, Warning -After (Get-Date).Date.AddDays(-1)

###########################################################################################################################################################################################

# Tip 53: Exporting Data to Excel


# You can easily convert object results to CSV files in PowerShell. This generates a CSV report of current processes:
Get-Process | Export-Csv $env:TEMP\report.csv -UseCulture -Encoding UTF8 -NoTypeInformation

# To open the CSV file in Microsoft Excel, you could use Invoke-Item to open the file,
# but this would only work if the file extension CSV is actually associated with the Excel application.

# This script will open the CSV file always in Microsoft Excel. 
# It illustrates a way to find out the actual path to your Excel application 
# (it assumes it is installed and does not check if it is missing altogether though):

$report = "$env:temp\report.csv"
$excelPath = 'C:\Program Files*\Microsoft Office\OFFICE*\EXCEL.EXE'
$realExcelPath = Resolve-Path $excelPath | Select-Object -First 1 -ExpandProperty Path
& $realExcelPath $report


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# This tip provides a nice demonstration of using wildcards to resolve the path to a file or executable. 
# Another method to open a CSV file in Excel takes advantage of the Start-Process cmdlet and the fact that 
# Excel is registered in the App Paths key in the registry. 
# In fact, any application that's registered in App Paths can be lauched using Start-Process in the same manner.

$report = "$env:temp\report.csv"
Start-Process excel -ArgumentList $report

# Note that if the path to the CSV file includes spaces, it must be passed as a quoted string, so the following would work.
Start-Process Excel -ArgumentList '"C:\Users\Some User\Documents\Import File.csv"'

# This won't work, however, if you require variable expansion. Instead, use:
Start Start-Process Excel -ArgumentList """$report"""

# You can use the following command to determine if Excel is registered in App Paths.
Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'          # Output: True


# To list all applications registered in App Paths, use:
Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"

###########################################################################################################################################################################################

# Tip 54: Finding Hard Drives Running Low on Storage


# WMI can retrieve information about drives easily. 
# This will get you the drive information for your local machine (use -ComputerName to access a remote system):

Get-WmiObject -Class Win32_LogicalDisk
# Output:
#        DeviceID     : C:
#        DriveType    : 3
#        ProviderName : 
#        FreeSpace    : 11271933952
#        Size         : 104752738304
#        VolumeName   : 


# To limit the results to only hard drives, and only those hard drives that have less than a given amount of free space:
$limit = 80GB
Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3 and Freespace<$limit" | Select-Object -Property VolumeName, Freespace, DeviceID

# Output:
#        VolumeName           Freespace           DeviceID                                                              
#        ----------           ---------           --------                                                              
#                           11391197184           C:      

###########################################################################################################################################################################################

# Tip 55: Drive Data in GB and Percent


# When a cmdlet returns raw data, you may want to convert the data into a better format. 
# For example, WMI can report the free space of a drive but reports bytes. 

# You can then use Select-Object and provide hash tables to take the raw data and convert it to anything you want. 
# This sample illustrates how to convert Freespace into GB and also calculate the percentage of free space:

$FreeSpace = @{

    Name = "Free Space(GB)"    
    Expression = { [int]($_.Freespace / 1GB) }
}

$PercentFree = @{

    Name = "Free (%)"
    Expression = { [int]($_.Freespace * 100 / $_.Size)}
}

Get-WmiObject -Class Win32_LogicalDisk | Select-Object -Property DeviceID, $FreeSpace, $PercentFree

# Output:
#        DeviceID      Free Space(GB)          Free (%)
#        --------      --------------          --------
#        C:                        11                11
#        D:                         0                  
#        E:                       348                95

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Here's a query that you can use which will return info about your volumes (mounts).  
# Notice that VolumeName is replaced with Name and Size becomes Capacity, 
# then plug this into an array and out to an email, database table, csv, or whatever.

$ServerName = 'YOURSERVER'
$Freespace = @{

    Expression = {[int]($_.Freespace/1GB)}
    Name = 'Free Space (GB)'
}

$PercentFree = @{

    Expression = {[int]($_.Freespace*100/$_.Capacity)}
    Name = 'Free (%)'
}

Get-WmiObject -namespace "root/cimv2" -computername $ServerName -query "SELECT Name, Capacity, FreeSpace FROM Win32_Volume WHERE Capacity > 0 and (DriveType = 2 OR DriveType = 3)" |
   Select-Object -Property Name, $Freespace, $PercentFree

###########################################################################################################################################################################################

# Tip 56: Finding Wireless Network Adapters


# There are many ways of finding network adapters, but apparently none to identify active wireless adapters.

# All information about your network adapters can be found right in the Registry,
# and here is a one-liner that provides all the information you may need:
Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Network\*\*\Connection" -ErrorAction SilentlyContinue | 
    Select-Object -Property Name, DefaultNameIndex, MediaSubType

# Output:
#        Name                              DefaultNameIndex MediaSubType                                                          
#        ----                              ---------------- ------------                                                          
#        Local Area Connection* 2                         2                                                                       
#        Local Area Connection* 8                         8                                                                       
#        Local Area Connection* 7                         7                

# The interesting part is the MediaSubType value. Wireless adapters always are marked with a MediaSubType of 2.
# So this line will always return wireless adapters only:

Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Network\*\*\Connection" -ErrorAction SilentlyContinue | 
    Where-Object { $_.MediaSubType -eq 2 } | Select-Object -Property Name, PnpInstanceID


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Get-WirelessAdapter

# Here is now a function Get-WirelessAdapter that returns all wireless adapters in your system:

function Get-WirelessAdapter
{
    Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Network\*\*\Connection" -ErrorAction SilentlyContinue | 
        Select-Object -Property MediaSubType, PnpInstanceID | Where-Object { $_.MediaSubType -eq 2 -and $_.PnpInstanceID } | 
        Select-Object -ExpandProperty PnpInstanceID | ForEach-Object {
            
            $wmipnpID = $_.Replace("\", "\\")
            Get-WmiObject -Class Win32_NetworkAdapter -Filter "PNPDeviceID='$wmipnpID'"
        }
}

Get-WirelessAdapter

# Output:
#       ServiceName:       BCM43xx
#       MACAddress:        68:A8:6D:0B:5F:CC
#       AdapterType:       Ethernet 802.3
#       DeviceID: 7
#       Name:              Broadcom 802.11n Network Adapter
#       NetworkAddresses:
#       Speed:             26000000

# Since the function returns a true WMI object, you can then determine whether the adapter is currently active, and enable or disable it. 
# This would identify the adapter, then disable it, then enable it again:

$adapter = Get-WirelessAdapter
$adapter.Disable().ReturnValue
$adapter.Enable().ReturnValue

# Note that a return code of 5 indicates that you do not have sufficient privileges. Run the script as an Administrator.

###########################################################################################################################################################################################

# Tip 57: Running Commands Elevated in PowerShell


# Sometimes, a script may need to run a command that needs elevation (Administrator privileges).

# Instead of requiring the script to run with full privileges altogether, you can also send individual commands to an elevated shell.

# This will restart the Spooler driver (which requires elevated privileges), and sends the command to another PowerShell process. 
# It will automatically start elevated if the current process does not have Administrator privileges.

$command = "Restart-Service -Name spooler"
Start-Process -FilePath powershell.exe -ArgumentList "-noprofile -command $command" -Verb runas

###########################################################################################################################################################################################

# Tip 58: Profiling Systems


# If you just want to profile a local or remote system and get back the most commonly used pieces of information, 
# then do not waste time for your own solutions. Simply reuse systeminfo.exe, and feed the data into PowerShell:

function Get-SystemInfo
{
    param($ComputerName = $env:COMPUTERNAME)

    $header = 'Hostname','OSName','OSVersion','OSManufacturer','OSConfiguration','OS Build Type','RegisteredOwner',
              'RegisteredOrganization','Product ID','Original Install Date','System Boot Time','System Manufacturer',
              'System Model','System Type','Processor(s)','BIOS Version','Windows Directory','System Directory',
              'Boot Device','System Locale','Input Locale','Time Zone','Total Physical Memory','Available Physical Memory',
              'Virtual Memory: Max Size','Virtual Memory: Available','Virtual Memory: In Use','Page File Location(s)',
              'Domain','Logon Server','Hotfix(s)','Network Card(s)'

    systeminfo.exe /FO CSV /S $ComputerName | Select-Object -Skip 1 | ConvertFrom-Csv -Header $header
}

Get-SystemInfo
# Output:
#        Hostname                  : SIHE-01
#        OSName                    : Microsoft Windows 7 
#        OSVersion                 : 6.1.7601 Service Pack 1 Build 7601
#        OSManufacturer            : Microsoft Corporation
#        OSConfiguration           : Member Workstation
#        OS Build Type             : Multiprocessor Free
#        RegisteredOwner           : Silence
#        ...                         

# When you store the result in a variable, you can easily access individual pieces of information:

$sysInfo = Get-SystemInfo

$sysInfo.HostName                  # Output: SIHE-01
$sysInfo."Logon Server"            # Output: \\SHA-02
$sysInfo."System Boot Time"        # Output: 2014/11/25, 11:50:49

# If you like the information to be called differently, simply change the list of property names to your liking. 
# So if you do not like "System Boot Time", simply rename this label in the script to "BootTime", for example.

###########################################################################################################################################################################################

# Tip 59: Applying NTFS Access Rules


# There are many ways to add or change NTFS permissions. One is to reuse existing tools such as icacls.exe.

# This function will create new folders that have default permissions. 
# The script uses icacls.exe to explicitly add full permissions to the current user and read permisssions to local Administrators:

function New-Folder
{
    param([String]$Path, [String]$UserName = "$env:userdomain\$env:username")

    if((Test-Path -Path $Path) -eq $false)
    {
        New-Item $Path -ItemType Directory | Out-Null
    }

    icacls $Path /inheritance:r /grant '*S-1-5-32-544:(OI)(CI)R' ('{0}:(OI)(CI)F' -f $UserName)
}

New-Folder -Path C:\Users\v-sihe\Desktop\testnew 

###########################################################################################################################################################################################

# Tip 60: Submitting Arguments to EXE Files


# Running applications such as robocopy.exe from PowerShell sometimes is not trivial. 
# How do you submit arguments to the EXE so that PowerShell won't change them?

# It really is simple: make sure all arguments are strings (so quote the arguments if they are no strings or contain spaces or other special characters). 
# And, make sure you submit one string per argument, not one big string. 


# This would invoke robocopy.exe from PowerShell and copy all JPG pictures from the Windows folder
# to another folder c:\jpegs recursively, not retrying on errors, and skipping the folder *winsxs*:

$arguments = "$env:windir\", "C:\jepgs\", "*.jpg", "/R:0", "/S", "/XD", "*winsxs*"
robocopy.exe $arguments

# As you can see, the arguments are all strings, and they are submitted as a string array. 

# That works beautifully for each and every exe file you may ever want to invoke from PowerShell.


# Anyone comments: Why you did not include all the argument in single quote?
$arguments = '$env:windir\ c:\jpegs\ *.jpg /R:0 /S /XD'

###########################################################################################################################################################################################

# Tip 61: Finding Expired Certificates


# PowerShell grants access to your certificate stores by using the cert: drive.

# You can use this drive to find certificates based on given criteria. 
# This would list all certificates that have a date in NotAfter that is before today (indicating expired certificates):

$today = Get-Date

Get-ChildItem -Path Cert:\ -Recurse | Where-Object { $_.NotAfter -ne $null } | Where-Object { $_.NotAfter -lt $today } | 
    Select-Object -Property FriendlyName, NotAfter, PSParentPath, Thumbprint | Out-GridView

###########################################################################################################################################################################################

# Tip 62: Finding Time Servers (And Reading All RegKey Values)


# Maybe you'd like to get a list of timeservers registered in the Registry database. Then you probably run code like this:

$timeServerPath = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime\Servers' 

Get-ItemProperty -Path $timeServerPath
# Output:
#        (default)    : 1
#        1            : time.windows.com
#        2            : time.nist.gov
#        3            : time-nw.nist.gov
#        4            : time-a.nist.gov
#        5            : time-b.nist.gov
#        PSPath       : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime\Servers
#        PSParentPath : Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime
#        PSChildName  : Servers
#        PSProvider   : Microsoft.PowerShell.Core\Registry

# This code accesses the Registry key, and then uses its methods to get the value names, then dumps the values:

$key = Get-Item -Path $timeServerPath
foreach($valuename in $key.GetValueNames())
{
    if($valuename -ne "")
    {
        $key.GetValue($valuename)
    }
}

# Output:
#        time.windows.com
#        time.nist.gov
#        time-nw.nist.gov
#        time-a.nist.gov
#        time-b.nist.gov

###########################################################################################################################################################################################

# Tip 63: Finding USB Stick Information


function Get-USBInfo
{
    param($FriendlyName = "*")

    $usbPath = 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\USBSTOR\*\*\'

    Get-ItemProperty -Path $usbPath | Where-Object { $_.FriendlyName } | Where-Object { $_.FriendlyName -like $FriendlyName } | 
        Select-Object -Property FriendlyName, Mfg | Select-Object FriendlyName
}

Get-USBInfo

Get-USBInfo -FriendlyName "*King*"  # You can query for specific manufacturers, too

###########################################################################################################################################################################################

# Tip 64: Exporting and Importing Credentials in PowerShell


# Credential objects contain a username and a password. 
# You can create them using Get-Credential, and then supply this object to any cmdlet that has the -Credential parameter. 

# However, what do you do if you want your scripts to run without user intervention yet securely? 
# You do not want a credentials dialog to pop up, and you do not want to store the password information inside the script.


# Here's a solution: use the function Export-Credential to save the credential to file:

function Export-Credential
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Path,

        [System.Management.Automation.Credential()]
        [Parameter(Mandatory = $true)]
        $Credential
    )

    $CredentialCopy = $Credential | Select-Object *
    $CredentialCopy.Password = $CredentialCopy.Password | ConvertFrom-SecureString
    $CredentialCopy | Export-Clixml $Path
}

Export-Credential -Path E:\Test\mycred -Credential domain\sihe           # This would save a credential for the user tobias to a file

# mycred file:
#   <Objs Version="1.1.0.1" xmlns="http://schemas.microsoft.com/powershell/2004/04">
#     <Obj RefId="0">
#       <TN RefId="0">
#         <T>Selected.System.Management.Automation.PSCredential</T>
#         <T>System.Management.Automation.PSCustomObject</T>
#         <T>System.Object</T>
#       </TN>
#       <MS>
#         <S N="UserName">domain\sihe</S>
#         <S N="Password">01000000d08c9ddf0115d1118c7a00c04f...</S>
#       </MS>
#     </Obj>
#   </Objs>


# Note that while you do this, the credentials dialog pops up and securely asks for your password. 
# The resulting file contains XML, and the password is encrypted.

# Now, when you need the credential, use Import-Credential to get it back from file:

function Import-Credential
{
    param([Parameter(Mandatory = $true)]$Path)

    $CredentialCopy = Import-Clixml $Path
    $CredentialCopy.Password = $CredentialCopy.Password | ConvertTo-SecureString

    New-Object System.Management.Automation.PSCredential($CredentialCopy.UserName, $CredentialCopy.Password)
}

$cred = Import-Credential -Path E:\Test\mycred
$cred

# Output:
#        UserName                                      Password
#        --------                                      --------
#        domain\sihe               System.Security.SecureString

# The "secret" used for encryption and decryption is your identity, 
# so only you (the user that exported the credential) can import it again. No need to hard-code secrets into your script.

###########################################################################################################################################################################################

# Tip 65: Enabling Classic Remoting


# Many cmdlets have built-in remoting capabilities, 
# for example Get-Service and Get-Process both have the parameter -ComputerName, and so does Get-WmiObject.

# However, to actually use these cmdlets remotely, the remoting technique employed by the cmdlets must be present.
# Most cmdlets that use classic remoting require the "Remote Administration" firewall rule to be enabled on the target side.
# It allows DCOM traffic. Some also require the remote Registry service to run on the target side.

# So in most scenarios, when you have Administrator privileges on the target machine and run these commands, 
# the machine will then be accessible for Administrators via classic cmdlet remoting:

netsh firewall set service remoteadmin enable

Start-Service RemoteRegistry

Set-Service -Name RemoteRegistry -StartupType Automatic

# Note that the netsh firewall command is considered obsolete on newer Windows versions but still works. 
# This command is much easier to use than the newer netsh advfirewall command.

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Enabling PowerShell Remoting


# If you'd like to use PowerShell Remoting to execute commands and scripts on another machine, 
# then you need to enable Remoting on the target side with full Admin privileges:

Enable-PSRemoting -SkipNetworkProfileCheck -Force

# On the client side, you do not need to do anything special when you are in the same domain and are using a domain account. 

# If you want to access a target computer without Kerberos authentication 
# (target is in another domain, or you want to use IP addresses or not fully qualified DNS names), 
# then you need to do this once with full Admin privileges:

Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value * -Force

# By setting trusted hosts to “*”, PowerShell will allow you to contact any IP or machine name you want, 
# and if the connection cannot be secured by Kerberos, NTLM authentication is used. 
# So this setting is not affecting who can contact a machine (use firewall rules for this). 
# It is simply telling PowerShell that you are willing to use the (less secure) NTLM authentication when Kerberos is unavailable.
# NTLM is less secure because it cannot tell whether the target computer is really the computer you want to reach. 
# While Kerberos does a mutual authentication, NTLM does not. Your credentials directly go to the machine you specify. 
# An attacker could have replaced the machine with his own and taken the machine IP address, for example, and you would not notice with NTLM.

# Note: Run Disable-PSRemoting only if you enabled remoting just to change the trusted hosts list and want to turn this off again afterwards. 
# Do not disable remoting if it was turned on before and possibly is in use by someone else.

Disable-PSRemoting

# Once Remoting is in place, you can visit remote systems using Enter-PSSession, 
# and you can run commands and scripts on these machines using Invoke-Command.

###########################################################################################################################################################################################

# Tip 66: Testing UNC Paths


# Test-Path can test whether or not a given file or folder exists. 
# This works fine for paths that use a drive letter, but can fail with pure UNC paths.

# At its simplest, this should return $true, and it does (provided you did not disable your administrative shares):

$path = "\\127.0.0.1\C$"
Test-Path -Path $path                    # Output: True

# Now, the very same code can also return $false:

Set-Location -Path HKCU:\ 
Test-Path -Path $path                    # Output: False 

# If a path does not use a drive letter, PowerShell falls back to the current path, and if that path happens to point to a non-file system location, 
# Test-Path interprets the UNC path in the context of that provider. Since there is no such path in your Registry, Test-Path returns $false.

# To make Test-Path work reliably with UNC paths, make sure you prepend the UNC path with the FileSystem provider. 
# Now, the result is valid regardless of current drive location:

Set-Location -Path HKCU:\ 
$path = "FileSystem::\\127.0.0.1\C$"
Test-Path -Path $path                    # Output: True

###########################################################################################################################################################################################

# Tip 67: Using Encrypting File System (EFS) to Protect Passwords


# If you absolutely need to hardcode passwords and other secrets into your scripts (which you should avoid for obvious reasons), 
# then you might still be safe when you encrypt the script with the EFS (Encrypting File System). 
# Encrypted scripts can only be read (and run) by the one that encrypted it, 
# so this works only if you are running the script yourself, and if you are running it from your machine.

# Here's an easy way of encrypting a PowerShell script:

$path = "$env:temp\test.ps1"
"Write-Host 'I run only for my master.'" > $path

$file = Get-Item -Path $path
$file.Encrypt()

# Once you run this, it will create a new PowerShell script in your temp folder that is encrypted by EFS 
# (if you get an error message instead, then EFS might either not be available or disabled on your machine).

# Once encrypted, the file will appear in green when viewed in Windows Explorer, 
# and only you will be able to run it. No one else can even see the source code.

# Note that in many corporate environments, EFS is set up with recovery keys that allow specific recovery personnel to decrypt files with a master key. 
# If no such master key exists, once you lose your EFS certificate, even you will not be able to view or run the script anymore.

###########################################################################################################################################################################################

# Tip 68: Storing Secret Data


# If you wanted to store sensitive data in a way that only you could retrieve it, you can use a funny approach: 
# convert some plain text into a secure string, then convert the secure string back, and save it to disk:

$storage = "$env:temp\secretdate.txt"
$mysecret = "Hello, I'm safe!"
$mysecret | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File -FilePath $storage

Get-Content -Path $storage
# Output:  01000000d08c9ddf0115d1118c7a00c04fc297eb010000007d6977d6bcc317408a0ded1e667e84d20000000002000000000003660000c0000000
#          1000000058ce42dacc9fa0c809528f37ddc767aa0000000004800000a000000010000000f19cf25b88843eb45a11f2dbe05baa7e2800000095732
#          b3e1c88e877a2941419f7f25fd8785f2e9c03363476cc754c2459b0988382f99941f78a3d4d14000000f40122cc9605fdc264b15dfd0126783de307c723

# Your secret was automatically encrypted by the built-in Windows data protection API (DPAPI), using your identity and your machine as encryption key. 
# So only you (or any process that runs on your behalf) can decipher the secret again, and only on the machine where it was encrypted.

$secureString = Get-Content -Path $storage | ConvertTo-SecureString
$ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($secureString)
$mysecret = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)

$mysecret       # Output: Hello, I'm safe!

# It works--you get back the exact same text that you encrypted before.

# Now, try the same as someone else. You will see that any other user cannot decrypt the secret file. 
# And you won't be able to, either, when you try it from a different machine.

###########################################################################################################################################################################################

# Tip 69: Reading All Text


# You can use Get-Content to read in any plain text file. 
# However, Get-Content will return the file content line by line, and you get back a string array. The line endings are consumed.

# To read in a text file in one big chunk, beginning with PowerShell 3.0, 
# you can use the –Raw parameter (which coincidentally speeds up reading the file quite a bit, too).

# So this gives you an array of string lines:

$content = Get-Content -Path $env:windir\WindowsUpdate.log    
$content -is [array]                                               # Output: True


# And this reads in the entire content in one big chunk, returning a single string:

$content = Get-Content -Path $env:windir\WindowsUpdate.log -Raw
$content -is [array]                                               # Output: False

# This time, the length is the number of overall bytes, and reading the file is much faster (albeit consumes more memory, too). 

# Which approach is right? It depends on what you want to do with the data.

###########################################################################################################################################################################################

# Tip 70: Logging Script Runtime


# Method One: Using Get-Date

$start = Get-Date
Get-HotFix
$end = Get-Date
Write-Host -ForegroundColor Red ('Total Runtime: ' + ($end - $start).TotalSeconds)                             

# Output: Total Runtime: 8.299




# Method Two: Using .Net

$stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
$stopWatch.Start()
Get-HotFix
$stopWatch.Stop()
Write-Host "Elapsed Runtime:" $stopWatch.Elapsed.Minutes "minutes and" $stopWatch.Elapsed.Seconds "seconds."

# Output: Elapsed Runtime: 0 minutes and 10 seconds.




# Method Three:

$start = Get-Date
Get-HotFix
$end = Get-Date
$ts = New-TimeSpan -Start $start -End $end
Write-Output "`nTotal runtime: $($TS.Seconds) seconds"

# Output: Total runtime: 8 seconds

###########################################################################################################################################################################################

# Tip 71: Converting Ticks into Real Date


# Internally, Active Directory uses ticks (100 nanosecond units since 1601) to represent date and time. 
# It has been hard in the past to convert these huge numbers into human readable date and time. Here is a much easier way:

[DateTime]::FromFileTime(635526118800000000)        # Output: Wednesday, November 26, 3614 23:18:00

(Get-Date -Date "2014/11/26 15:18:00").Ticks        # Output: 635526118800000000



# Note that Get-Date will accept a value representing ticks as an argument
Get-Date 635526118800000000                         # Output: Wednesday, November 26, 2014 15:18:00

[DateTime]::FromBinary(635526118800000000)          # Output: Wednesday, November 26, 2014 15:18:00

###########################################################################################################################################################################################

# Tip 72: Parallel Processing in PowerShell


# If a script needs some speed-up, you might find background jobs helpful. 
# They can be used if a script does a number of things that also could run concurrently.

# PowerShell is single-threaded and can only do one thing at a time. 
# With background jobs, additional PowerShell processes run in the background and can share the workload. 
# This works well only if the jobs you need to do are completely independent from each other, 
# and if your background job does not need to produce a lot of data. 
#　Sending back data from a background job is an expensive procedure that can easily eat up all the saved time, resulting in an even slower script.

# Here are three tasks that all can run concurrently:

$start = Get-Date

$task1 = { Get-HotFix }
$task2 = { Get-Service | Where-Object Status -EQ Running }
$task3 = { Get-Content -Path $env:windir\WindowsUpdate.log | Where-Object { $_ -like '*successfully installed*' }}

# run 2 tasks in the background, and 1 in the foreground task
$job1 = Start-Job -ScriptBlock $task1
$job2 = Start-Job -ScriptBlock $task2
$result3 = Invoke-Command -ScriptBlock $task3

# wait for the remaining tasks to complete (if not done yet)
$null = Wait-Job -Job $job1,$job2

# now they are done, get the results
$result1 = Receive-Job -Job $job1
$result2 = Receive-Job -Job $job2

# discard the jobs
Remove-Job -Job $job1,$job2

$end = Get-Date

Write-Host -ForegroundColor Red ($end - $start).TotalSeconds         # Output: 7.771

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

$start = Get-Date

$task1 = { Get-HotFix }
$task2 = { Get-Service | Where-Object Status -EQ Running }
$task3 = { Get-Content -Path $env:windir\WindowsUpdate.log | Where-Object { $_ -like '*successfully installed*' }}

Invoke-Command -ScriptBlock $task1
Invoke-Command -ScriptBlock $task2
Invoke-Command -ScriptBlock $task3

$end = Get-Date

Write-Host -ForegroundColor Red ($end - $start).TotalSeconds         # Output: 9.445

# So background jobs really pay off for long running tasks that all take almost the same time. 
# Since the three sample tasks returned a lot of data, the benefit of executing them concurrently was eliminated 
# by the overhead that it took to serialize the return data and transport it back to the foreground process.

###########################################################################################################################################################################################

# Tip 73: Running Background Jobs Efficiently


# Using background jobs to run tasks concurrently often is not very efficient as you might have seen. 
# Background job performance worsens with the amount of data that is returned by a background job.

# A much more efficient way uses in-process tasks. 
# They run as separate threads inside the very same PowerShell, so there is no need to serialize return values.

# Here is a sample that runs two processes in the background, and one in the foreground, using PowerShell threads. 
# To create some really long-running tasks, each task uses Start-Sleep in addition to some other command:

$start = Get-Date

$task1 = { Start-Sleep -Seconds 3; Get-Service }
$task2 = { Start-Sleep -Seconds 4; Get-Service }
$task3 = { Start-Sleep -Seconds 5; Get-Service }

$thread1 = [Powershell]::Create()
$job1 = $thread1.AddScript($task1).BeginInvoke()

$thread2 = [Powershell]::Create()
$job2 = $thread2.AddScript($task2).BeginInvoke()

$result3 = Invoke-Command -ScriptBlock $task3

do
{
    Start-Sleep -Milliseconds 100
}until($job1.isCompleted -and $job2.isCompleted)

$result1 = $thread1.EndInvoke($job1)
$result2 = $thread2.EndInvoke($job2)

$thread1.Runspace.Close()
$thread1.Dispose()

$thread2.Runspace.Close()
$thread2.Dispose()

$end = Get-Date

Write-Host -ForegroundColor Red ($end - $start).TotalSeconds     # Output: 5.5181034

# Running these three tasks consecutively would take at least 12 seconds for the Start-Sleep statements alone. 
# In reality, the script only takes a bit more than 5 seconds. The result can be found in $result1, $result2, and $result3. 
# In contrast to background jobs, there is almost no time penalty for returning large amounts of data. 

###########################################################################################################################################################################################

# Tip 74: Getting Events From All Event Logs


# Recently, a reader asked how to retrieve all events from all event logs from a local or remote system, and optionally save them to file.

# Here is a potential solution:

$start = (Get-Date) - (New-TimeSpan -Hours 1)
$ComputerName = $env:COMPUTERNAME

Get-EventLog -AsString -ComputerName $ComputerName | ForEach-Object {

    Write-Progress -Activity "Checking Eventlogs on \\$ComputerName" -Status $_

    Get-EventLog -LogName $_ -EntryType Error, Warning -After $start -ComputerName $ComputerName -ErrorAction SilentlyContinue | 
        Add-Member NoteProperty -TypeName NoteProperty -Name EventLog -Value  $_ -PassThru
         
} | Sort-Object -Property TimeGenerated -Descending | 
    Select-Object EventLog, TimeGenerated, EntryType, Source, Message | 
    Out-GridView -Title "All Errors & Warnings from \\$Computername" 

# At the top of this script, you can set the remote system you want to query, and the number of hours you want to go back.

# Next, the script gets all log files available on that machine, 
# and then uses a loop to get the errors and warnings from each log within the timeframe. 
# To be able to know which event came from which log file, it also tags the events with a new property called "EventLog", using Add-Member.

# The result is a report with all error and warning events within the last hour, shown in a grid view window. 
# Replace "Out-GridView" with "Out-File" or "Export-Csv" to write the information to disk.

# Note that remote access requires Administrator privileges. Remote access might require additional security settings. 
# Note also that you will receive red error messages if you run this code as a non-Administrator 
# (because some logs like "Security" require special access privileges).

###########################################################################################################################################################################################

# Tip 75: Hiding Terminating Errors


# Occasionally, you may have noticed that cmdlets throw errors although you specified "SilentlyContinue" as -ErrorAction.

# The -ErrorAction parameter can only hide non-terminating errors (errors that originally were handled by the cmdlet). 
# Any error that was not handled by the cmdlet is called "terminating". These errors are typically security-related and never covered by -ErrorAction.

# So if you are a non-Administrator, the following call will raise an exception even though -ErrorAction asked to suppress errors:

Get-EventLog -LogName Security -Newest 10 -ErrorAction SilentlyContinue  # Exception here: Get-EventLog : Requested registry access is not allowed.


# To suppress terminating errors, you must use an error handler:

try
{
    Get-EventLog -LogName Security -Newest 10

}catch{}



# A nice addition from comment:

Catch { $Error.RemoveAt(0) }      # This will remove the latest error from the $Error array.

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Catching Non-Terminating Errors


# Non-terminating errors are errors that are handled by a cmdlet internally. Most errors that can occur in cmdlets are non-terminating. 

# You cannot catch these errors in an error handler. So although there is an error handler in this sample, it will not catch the cmdlet error:

try
{
    Get-WmiObject -Class Win32_BIOS -ComputerName offlineServer
}
catch
{
    Write-Warning "Oops, error $_"
}

# Exception here: Get-WmiObject : The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)


# To catch non-terminating errors, you must turn them into terminating errors. That is done by setting the -ErrorAction to "Stop":

try
{
    Get-WmiObject -Class Win32_BIOS -ComputerName offlineServer -ErrorAction Stop
}
catch
{
    Write-Warning "Oops, error $_"
}

# You can temporarily set $ErrorActionPreference to "Stop" if you do not want to add
# a -ErrorAction Stop parameter to all cmdlets within your error handler. 
# The preference is used if a cmdlet does not explicitly specify an -ErrorAction setting.

###########################################################################################################################################################################################

# Tip 76: Logging All Errors


# Cmdlet errors can only be caught by an error handler if the -ErrorAction is set to "Stop". 
# Doing this alters the way the cmdlet works, though. It will then stop at the first error encountered.

# Take the next example: it scans the windows folder for PowerShell scripts recursively. 
# If you wanted to catch errors (accessing protected subfolders, for example), this would not work:

try
{
    Get-ChildItem -Path $env:windir -Filter *.ps1 -Recurse -ErrorAction Stop
}
catch
{
    Write-Warning "Error: $_"
}

# The code would catch the first error, but the cmdlet would stop and not continue to scan the remaining subfolders.

# If you just suppressed errors, you would get complete results, but now your error handler would not catch anything:

try
{
    Get-ChildItem -Path $env:windir -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue
}
catch
{
    Write-Warning "Error: $_"
}

# So if you want a cmdlet to run uninterruptedly, and still get a list of all folders that you were unable to access, 
# do not use an error handler at all. Instead, use -ErrorVariable and log the errors silently to a variable.

# After the cmdlet has finished, you can evaluate the variable and get an error report:
Get-ChildItem -Path $env:windir -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue -ErrorVariable myErrors

foreach($incidence in $myErrors)
{
    Write-Warning ("Unable to access " + $incidence.CategoryInfo.TargetName)
} 

###########################################################################################################################################################################################

# Tip 77: Writing Events to Own Event Logs


# Often, there is a need to log information when a script runs. 
# Instead of writing log information to a text file that you would have to maintain and manage yourself, 
# you can use the built-in Windows logging system with all of its benefits, too.

# To do this, you just need to initialize your own log. 
# This needs to be done once by an Administrator, so launch an elevated PowerShell, and then enter a line like this:
New-EventLog -LogName ScriptIncidents -Source LogonScript, MaintenanceScript, Miscellaneous

# That's it. You now have a log file that can log incidents with the sources "LogonScript", "MaintenanceScript", and "Miscellaneous". 
# You might just want to configure it a bit more and tell the logging system how large the file can grow and what should occur when it runs full:
Limit-EventLog -LogName ScriptIncidents -RetentionDays 30 -OverflowAction OverwriteOlder -MaximumSize 500MB

# Now, your new log file could grow as large as 500MB, and entries would be kept for 30 days until they get overwritten by newer entries.


# You can now close your elevated shell. 
# Writing to your log file does not require special privileges and can occur from any vanilla script or logon script. 
# So switch to a regular PowerShell console, and try this:

Write-EventLog -LogName ScriptIncidents -Source LogonScript -EntryType Information -EventId 123 -Message "Logonscript stared"
Write-EventLog -LogName ScriptIncidents -Source LogonScript -EntryType Error -EventId 999 -Message "Unable to do something really import"
Write-EventLog -LogName ScriptIncidents -Source MaintenanceScript -EntryType Warning -EventId 522 -Message "Running low on disk space"

# Logging incidents is now a snap, and you can pick any event ID or message you want. You are only limited to the registered event sources.

# Using Get-EventLog, you can now easily analyze scripting problems on that machine:
Get-EventLog -LogName ScriptIncidents

Remove-EventLog -LogName ScriptIncidents

###########################################################################################################################################################################################

# Tip 78: Finding Registered Event Sources


# Each Windows log file has a list of registered event sources. 
# To find out which event sources are registered to which event log, you can directly query the Windows Registry.

# This will dump all registered sources for the "System" event log:

$logName = "System"
$path = "HKLM:\System\CurrentControlSet\services\eventlog\$logName"

Get-ChildItem -Path $path -Name

# Output:
#        ACPI
#        AeLookupSvc
#        AmdK8
#        amdkmdag
#        amdkmdap
#        AmdPPM
#        APPHOSTSVC
#        ...

###########################################################################################################################################################################################

# Tip 79: Getting Picture URLs from Google Picture Search


# Invoke-WebRequest is your friend whenever you want to download information from the Internet. 
# You could, for example, send a search request to Google and have PowerShell examine the results.

# Google knows about this, too, so when you send a search query from PowerShell, Google responds with encrypted links. 
# To get the real links, you would have to tell Google that you are not PowerShell but just a regular browser. 
# This can be done by setting a browser agent string.

# This script takes any keyword and returns the original picture sources for any picture 
# that matches your keyword and has at least a 2 megapixel resolution:

$keyword = "Powershell"

$url = "https://www.google.com/search?q=$keyword&espv=210&es_sm=93&source=lnms&tbm=isch&sa=X&tbm=isch&tbs=isz:lt%2Cislt:2mp"

$browserAgent = 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/33.0.1750.146 Safari/537.36'

$page = Invoke-WebRequest -Uri $url -UserAgent $browserAgent

$page.Links | Where-Object { $_.href -like "*imgres*" } | ForEach-Object { ($_.href -split "imgurl=")[-1].Split("&")[0] } | Out-GridView



# Download the pictures you just searched (get the first 5 results):

New-Item -Path "$env:temp\IMGSearched" -ItemType Directory
Push-Location "$env:temp\IMGSearched"

$page.Links | Where-Object { $_.href -like "*imgres*" } | Select-Object -First 5 | ForEach-Object {
    
    $picutureLink = ($_.href -split "imgurl=")[-1].Split("&")[0]
     
    $fileName = [System.IO.Path]::GetFileName($picutureLink)

    Invoke-WebRequest -Uri $picutureLink -OutFile $fileName   
}

Pop-Location

Invoke-Item "$env:temp\IMGSearched"

###########################################################################################################################################################################################

# Tip 80: Open MsgBox with Random Sound


# You may have seen script code that opens a MsgBox dialog box. 
# Today, you get a piece of code that opens a MsgBox and plays a random sound, adding extra attention and fun. 
# The sound stops when the user responds to the MsgBox:

$randomWAV = Get-ChildItem -Path C:\Windows\Media -Filter *.wav | Get-Random | Select-Object -ExpandProperty FullName

Add-Type -AssemblyName System.Windows.Forms

$player = New-Object Media.SoundPlayer $randomWAV
$player.Load()
$player.PlayLooping()

$result = [System.Windows.Forms.MessageBox]::Show("We will reboot your machine now. Ok?", "PowerShell", "YesNo", "Exclamation")
$player.Stop()

###########################################################################################################################################################################################

# Tip 81: Getting Executable from Command Line


# Sometimes it becomes necessary to extract the command name from a command line. Here is a function that can do this:

function Remove-Argument
{
    param($CommandLine)

    $divider = " "

    if($CommandLine.StartsWith('"'))
    {
        $divider = '"'
        $CommandLine = $CommandLine.SubString(1)
    }

    $CommandLine.Split($divider)[0]
}

Remove-Argument -CommandLine 'explorer.exe c:\windows\somfolder'       # Output: explorer.exe
Remove-Argument -CommandLine '"explorer.exe" c:\windows\somfolder'     # Output: explorer.exe

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Getting Arguments from Command Line


function Get-Argument
{
    param($CommandLine)

    $result = 1 | Select-Object -Property Command, Argument

    if($CommandLine.StartsWith('"'))
    {
        $index = $CommandLine.IndexOf('"', 1)
        
        if($index -gt 0)
        {
            $result.Command = $CommandLine.SubString(0, $index).Trim('"')
            $result.Argument = $CommandLine.SubString($index + 1).Trim()

            $result
        }
    }
    else
    {
        $index = $CommandLine.IndexOf(" ")

        if($index -gt 0)
        {
            $result.Command = $CommandLine.SubString(0, $index)
            $result.Argument = $CommandLine.SubString($index + 1).Trim()

            $result
        }
    }
}

Get-Argument -CommandLine 'notepad c:\test'
Get-Argument -CommandLine '"notepad.exe" c:\test'
Get-Argument -CommandLine '"shutdown.exe" -r -t'

# Output:
#        Command                        Argument                                                                                                  
#        -------                        --------                                                                                                  
#        notepad                        c:\test                                                                                                   
#        notepad.exe                    c:\test                                                                                                   
#        shutdown.exe                   -r -t 


# And here is a real world example: it takes all running processes, and returns the commands and arguments:

Get-WmiObject -Class Win32_Process | Where-Object { $_.CommandLine } | ForEach-Object { Get-Arguemnt -CommandLine $_.CommandLine }

# Output:
#        Command                                                                               Argument                                                                                                  
#        -------                                                                               --------                                                                                                  
#        taskhost.exe                                                                                                                                                                                    
#        C:\Windows\system32\Dwm.exe                                                                                                                                                                     
#        C:\Program Files\Zune\ZuneLauncher.exe                                                                                                                                                          
#        C:\Program Files\Microsoft Security Client\msseces.exe                                -hide -runkey                                                                                             
#        C:\Windows\System32\StikyNot.exe                                                                                                                                                                
#        C:\Program Files (x86)\Skype\Phone\Skype.exe                                          /minimized /regrun                                                                                        
#        C:\Program Files\Microsoft Office 15\root\office15\ONENOTEM.EXE                       /tsr                                                                                                      
#        C:\Program Files (x86)\Common Files\Microsoft Shared\IME14WR\SHARED\IMECMNT.EXE       -Embedding 

# Now that command and argument are separated, you could also group information like this:

Get-WmiObject -Class Win32_Process | Where-Object { $_.CommandLine } | ForEach-Object { Get-Arguemnt -CommandLine $_.CommandLine } |
     Group-Object -Property Command | Sort-Object -Property Count -Descending | Out-GridView

###########################################################################################################################################################################################

# Tip 82: Finding Default MAPI Client


# Your MAPI client is the email client that by default is used with URLs like "mailto:". 
# To find out if there is a MAPI client, and if so, which one it is, 
# here is a function that retrieves this information from the Windows Registry. 

function Get-MAPIClient
{
    function Remove-Argument
    {
        param($CommandLine)

        $divider = " "

        if($CommandLine.StartsWith('"'))
        {
            $divider = '"'
            $CommandLine = $CommandLine.Substring(1)
        }

        $CommandLine.Split($divider)[0]
    }

    $path = 'Registry::HKEY_CLASSES_ROOT\mailto\shell\open\command'
    
    $returnValue = 1 | Select-Object -Property HasMAPIClient, Path, MailTo

    $returnValue.HasMAPIClient = Test-Path -Path $path
    
    if($returnValue.HasMAPIClient)
    {
        $values = Get-ItemProperty -Path $path

        $returnValue.MailTo = $values."(Default)"
        $returnValue.Path = Remove-Argument -CommandLine $returnValue.MailTo

        if((Test-Path $returnValue.Path) -eq $false)
        {
            $returnValue.HasMAPIClient = $false
        }
    }

    $returnValue
}

Get-MAPIClient

# Output:
#    HasMAPIClient Path                                                            MailTo                                                                
#    ------------- ----                                                            ------                                                                
#             True C:\Program Files\Microsoft Office 15\Root\Office15\OUTLOOK.EXE  "C:\Program Files\Microsoft Office 15\Root\Office15\OUTLOOK.EXE" -c...

###########################################################################################################################################################################################

# Tip 83: Sending Email via Outlook


# Of course you can send emails directly via SMTP server using Send-MailMessage. 
# But if you want to prefill an email form in your default MAPI client, this is not very much harder either:

$subject = "Sending via MAPI client"
$body = "My Message"
$to = "hhbstar@hotmail.com"

$mail = "mailto:$to&subject=$subject&body=$body"

Start-Process -FilePath $mail

# This script takes advantage of the mailto: moniker. Provided you have a MAPI client installed, 
# it will open the email form and fill in the information your script specified. You do have to send the mail manually, though.

###########################################################################################################################################################################################

# Tip 84: Showing WPF Info Message


# WPF (Windows Presentation Foundation) is a technology that enables you to create windows and dialogs. 
# The advantage of WPF is that the window design can be separated from program code.

# Here is a sample that displays a catchy message. The message is defined in XAML code which works similar to HTML (but is case-sensitive). 
# You can easily adjust font size, text, colors, etc. without the need to touch any code:

$xaml = @"

    <Window
     xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'>
    
     <Border BorderThickness="20" BorderBrush="Yellow" CornerRadius="9" Background='Red'>
      <StackPanel>
       <Label FontSize="50" FontFamily='Stencil' Background='Red' Foreground='White' BorderThickness='0'>
        System will be rebooted in 15 minutes!
       </Label>
    
       <Label HorizontalAlignment="Center" FontSize="15" FontFamily='Consolas' Background='Red' Foreground='White' BorderThickness='0'>
        Worried about losing data? Talk to your friendly help desk representative and freely share your concerns!
       </Label>
      </StackPanel>
     </Border>
    </Window>
"@

Add-Type -AssemblyName PresentationFramework     # need add this if you run it in real powershell console but not in ISE

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader] $xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$window.AllowsTransparency = $true
$window.SizeToContent = "WidthAndHeight"
$window.ResizeMode = "NoResize"
$window.Opacity = .7
$window.Topmost = $true
$window.WindowStartupLocation = "CenterScreen"
$window.WindowStyle = "None"

$null = $window.Show()

Start-Sleep -Seconds 5

$window.Close()

###########################################################################################################################################################################################

# Tip 85: Automation via Keystroke and Mouse Click


# Occasionally, the only way of automating processes is to send keystrokes or mouse clicks to UI elements. 
# A good and free PowerShell extension is called "WASP" and is available here:

# Reference: http://wasp.codeplex.com/ 

# Once you install the module (do not forget to unblock the ZIP file before you unpack it, 
# via right-click, Properties, Unblock), the WASP module provides the following cmdlets:

Get-WindowPosition
Remove-Window
Select-ChildWindow
Select-Window
Send-Click
Send-Keys
Set-WindowActive
Set-WindowPosition

# Here is a simple automation example using the Windows calculator:

Import-Module WASP 

# launch Calculator
$process = Start-Process -FilePath calc -PassThru
$id = $process.Id
Start-Sleep -Seconds 2
$window = Select-Window | Where-Object { $_.ProcessID -eq $id }

# send keys
$window | Send-Keys 123
Start-Sleep -Seconds 1

$window | Send-Keys '{+}'
Start-Sleep -Seconds 1

$window | Send-Keys 999
Start-Sleep -Seconds 1

$window | Send-Keys =
Start-Sleep -Seconds 1

# send CTRL+c
$window | Send-Keys '^c'

# Result is now available from clipboard


#  And here are the caveats:
#  
#  Once you launch a process, allow 1-2 seconds for the window to be created before you can use WASP to find the window 
#  Sending keys follows the SendKeys API. Some characters need to be "escaped" by placing braces around them. 
#  More details here: http://msdn.microsoft.com/en-us/library/system.windows.forms.sendkeys.send(v=vs.110).aspx/ 
#  When sending control key sequences such as CTRL+C, make sure you use a lowercase letter. 
#  "^c" would send CTRL+c whereas "^C" would send CTRL+SHIFT+C 
#  Access to child controls like specific textboxes or buttons is supported for WinForms windows only (Select-ChildWindow, Select-Control). 
#  WPF windows can receive keys, too, but with WPF you have no control over the UI element on the window that receives the input. 


# Details: http://wasp.codeplex.com/releases/view/22118?RateReview=true

###########################################################################################################################################################################################

# Tip 86: Updating Windows Defender Signatures


# Windows 8.1 comes with a ton of new cmdlets.
# One of them can automatically download and install the latest antivirus signatures for Windows Defender:

Update-MpSignature

Get-MpComputerStatus            # returns information about the state of your signatures. 

# These cmdlets are not part of PowerShell but rather part of Windows 8.1, 
# so on previous OS versions, you will receive an error message complaining about a missing command.

###########################################################################################################################################################################################

# Tip 87: Getting Variable Value in Parent Scope


# If you define variables in a function, then these variables live in function scope. To find out what the value of the variable is in the parent scope, use Get-Variable with the parameter -Scope:


$a = 1

function Test
{
    $a = 2

    $parentVar = Get-Variable -Name a -Scope 1
    $parentVar.Value
}

Test    # Output: 1

# When the script calls "test", the function defines $a and sets it to 2.
# In the caller scope, variable $a is 1. By using Get-Variable, the function can find out the variable value in the parent scope.

###########################################################################################################################################################################################

# Tip 88: Use JSON to Create Objects


# JSON is describing objects, similar to XML--but a lot easier. 
# JSON allows for nested object properties, so you can retrieve information from various sources and consolidate them into one custom object.

# Have a look. This creates an inventory item containing various computer details:

$json = @"

{
    "ServerName": "$env:ComputerName",
    "UserName": "$env:UserName",
    "BIOS": 
        {
            "Manufacturer" : "$((Get-WmiObject -Class Win32_BIOS).Manufacturer)",
            "Version" : "$((Get-WmiObject -Class Win32_BIOS).Version)",
            "Serial" : "$((Get-WmiObject -Class Win32_BIOS).SerialNumber)"
        },
    "OS" : "$([Environment]::OSVersion.VersionString)"
}

"@

$info = ConvertFrom-Json -InputObject $json

$info.ServerName
$info.BIOS.Version
$info.OS

# Output:
#        SIHE-01
#        HPQOEM - 20090825
#        Microsoft Windows NT 6.1.7601 Service Pack 1



# You can then manipulate the resulting objects - retrieve information, or add/update details. 

# If you made changes to the object, you can use ConvertTo-Json to convert it back into JSON format for serialization:

$info.ServerName = "test"

$info | ConvertTo-Json

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


$json = @"

{
    "Name": "Silence",
    "ID": 123
}
"@

$info = ConvertFrom-Json -InputObject $json
$info.Name
$info.ID

$info | Get-Member

# Output:
#        Name        MemberType   Definition                    
#        ----        ----------   ----------                    
#        Equals      Method       bool Equals(System.Object obj)
#        GetHashCode Method       int GetHashCode()             
#        GetType     Method       type GetType()                
#        ToString    Method       string ToString()             
#        ID          NoteProperty System.Int32 ID=123           
#        Name        NoteProperty System.String Name=Silence

###########################################################################################################################################################################################

# Tip 89: Creating Excel Reports


# PowerShell objects can easily be opened in Microsoft Excel. 
# Simply export the objects to CSV, then open the CSV file with the associated program (which should be Excel if it is installed).

# This creates a report of running processes and opens in Excel:

$path = "$env:temp\$(Get-Random).csv"
$title = 'Name', 'Id', 'Company', 'Description', 'WindowTitle'

Get-Process | Select-Object -Property $title | Export-Csv -Path $path -Encoding UTF8 -NoTypeInformation -UseCulture

Invoke-Item $path

# Note how -UseCulture automatically picks the correct delimiter based on your regional settings.


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Fixing Display in Excel Reports

$Path = "$env:temp\$(Get-Random).csv"

Get-EventLog -LogName System -EntryType Error -Newest 10 | 
  Select-Object EventID, MachineName, Data, Message, Source, ReplacementStrings, InstanceId, TimeGenerated |
  Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation -UseCulture

Invoke-Item -Path $Path 

# The columns "Data" and "ReplacementStrings" is unusable. Since both properties contain arrays, 
# the auto-conversion simply displays the name of the data type. This is a phenomenon found often with Excel reports created from object data.

# To improve the report, you can explicitly use the PowerShell engine to convert object to text, 
# then turn multiple lines of text in one single line of text. 

Get-EventLog -LogName System -EntryType Error -Newest 10 | 
    Select-Object EventID, MachineName, Data, Message, Source, ReplacementStrings, InstanceId, TimeGenerated | ForEach-Object {

        $_.Message = ($_.Message | Out-String -Stream) -join ' '
        $_.Data = ($_.Data | Out-String -Stream) -join ', '
        $_.ReplacementStrings = ($_.ReplacementStrings | Out-String -Stream) -join ', '

        $_

  } | Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation -UseCulture

Invoke-Item -Path $Path 

# Now all columns show correct results. Note how the problematic properties were first sent to Out-String 
# (using PowerShell's internal mechanism to convert the data to meaningful text), then -join was used to turn the information to a single line of text.

# Note also how the property "Message" was processed. Although this property seemed to be OK, in reality it can be a multiline text. 
# Multiline messages would only show the first line in Excel, appended by "…". By joining the lines with a space, Excel shows the full message.

###########################################################################################################################################################################################

# Tip 90: Bulk Renaming Object Properties


# Occasionally, it may become necessary to bulk rename object properties to create better reports. 
# For example, if you retrieve process objects, you may need to create a report with column headers that are different from the original object properties.

# Here is a filter called Rename-Property that can rename any property for you. In the example, a process list is generated, and some properties are renamed:

filter Rename-Property ([Hashtable]$PropertyMapping)
{
    Foreach ($key in $PropertyMapping.Keys)
    {
        $_ = $_ | Add-Member -MemberType AliasProperty -Name $PropertyMapping.$key -Value $key  -PassThru
    }

    $_
}

$newProps = @{

    Company = 'Manufacturer'
    Description = 'Purpose'
    MainWindowTitle = 'TitlebarText'
}

# get raw data
Get-Process | Rename-Property $newProps | Select-Object -Property Name, Manufacturer, Purpose, TitlebarText 

# Rename-Property automatically adds all the properties specified in $newProps. 
# The resulting objects have new properties named "Manufacturer", "Purpose", and "TitlebarText". 
# You can then use Select-Object to select the properties you want in your report. 
# You can choose from the existing original properties and the newly added alias properties.

# So actually, the properties are not renamed (which is technically impossible). 
# Rather, the filter adds alias properties with a new name that point back to the original properties.

###########################################################################################################################################################################################

# Tip 91: Compiling Binary Cmdlets


$code = @"

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Management.Automation;

namespace CustomCmdlet
{
    [Cmdlet("Get", "Magic", SupportsTransactions = false)]
    public class test: PSCmdlet
    {
        private int _Age;
       
        [Alias(new String[]
        {
            "HowOld", "YourAge"
        }), Parameter(Position = 0, ValueFromPipeline = true)]

        public int Age
        {
            get { return _Age; }
            set { _Age = value; }
        }

        private String _Name;

        [Parameter(Position = 1)]
        public String Name
        {
            get { return _Name; }
            set { _Name = value; }
        }

        protected override void BeginProcessing()
        {
            this.WriteObject("Good Morning ...");
            base.BeginProcessing();
        }

        protected override void ProcessRecord()
        {
            this.WriteObject("Your name is " + Name + " and your age is " + Age);
            base.ProcessRecord();
        }

        protected override void EndProcessing()
        {
            this.WriteObject("That's it for now.");
            base.EndProcessing();
        }
    }  
}
"@

$datetime = Get-Date -Format yyyyMMddHHmmssffff
$dllPath = "$env:temp\myCmdlet($datetime).dll"

Add-Type -TypeDefinition $code -OutputAssembly $dllPath

Import-Module -Name $dllPath -Verbose

# Output:
#        VERBOSE: Loading module from path 'C:\Users\sihe\AppData\Local\Temp\myCmdlet(201411271650278304).dll'.
#        VERBOSE: Importing cmdlet 'Get-Magic'.


Import-Module -Name $dllPath


# Now you are ready to use the new Get-Magic cmdlet. 
# It features all the things a cmdlet can do, including parameters, parameter aliases, and even pipeline support:

Get-Magic 


Get-Magic -Age 12 -Name "silence"
# Output:
#        Good Morning ...
#        Your name is silence and your age is 12
#        That's it for now.


1..3 | Get-Magic -Name "Tom"
# Output:
#        Good Morning ...
#        Your name is Tom and your age is 1
#        Your name is Tom and your age is 2
#        Your name is Tom and your age is 3
#        That's it for now.


# Note that the majority of PowerShell code in the example is needed only to create and compile the DLL. 
# Once the DLL exists, all you need is this line (for example, in your shipping product):

Import-Module -Name $dllPath

# To develop sophisticated binary cmdlets, you will want to work in a C# development environment such as Visual Studio. 
# All you need is adding a reference to the PowerShell assembly. The path to the PowerShell assembly can be easily retrieved with this line:

[PSObject].Assembly.Location | clip     # It will place the path to the PowerShell assembly into your clipboard

[PSObject].Assembly.Location
# Output:
#    C:\Windows\Microsoft.Net\assembly\GAC_MSIL\System.Management.Automation\v4.0_3.0.0.0__31bf3856ad364e35\System.Management.Automation.dll

# Note that compiling a C# code by itself will not add much additional protection to your intellectual property as it can be decompiled. 
# So do not use this to "protect" secret information such as passwords. With binary cmdets, 
# you do get the option to use professional copy protection software and obfuscators. 
# These extra layers of protection are not available with plain PowerShell code. 

###########################################################################################################################################################################################

# Tip 92: Converting Text Arrays to String


# Occasionally, text from a text file needs to be read and processed by other commands. 
# Typically, you would use Get-Content to read the text file content, then pass the result on to other commands. This may fail, though.

# And here is the caveat: Always remember that Get-Content returns an array of text lines, not a single text line. 
# So whenever a command expects a string rather that a bunch of text lines (a string array), you need to convert text lines to text.

# Beginning in PowerShell 3.0, Get-Content has a new switch parameter called -Raw. 
# It will not only speed up reading large text files but also return the original text file content in one chunk, without splitting it into text lines.

$info = Get-Content "$env:windir\windowsupdate.log"
$info -is [array]                                          # Output: True

$info = Get-Content "$env:windir\windowsupdate.log" -Raw
$info -is [array]                                          # Output: False


# If you already have text arrays and would like to convert them to a single text, use Out-String:

$info = "One", "Two", "Three"
$info -is [array]                                          # Output: True

$all = $info | Out-String
$all -is [array]                                           # Output: False

###########################################################################################################################################################################################

# Tip 93: Adding and Resetting NTFS Permissions


# Whether you want to add a new NTFS access rule to a file or turn off inheritance and add new rules, 
# here is a sample script that illustrates the trick and can serve you as a template.

# The script creates a test file, then defines a new access rule for the current user. 
# This rules allows read and write access. The new rule is added to the existing security descriptor. In addition, inheritance is turned off. 

$path = "$env:temp\exampleFile.txt"
$null = New-Item -Path $path -ItemType File -ErrorAction SilentlyContinue

$username = "$env:userdomain\$env:username"

$colRights = [System.Security.AccessControl.FileSystemRights]"Read, Write"
$inheritanceFlag = [System.Security.AccessControl.InheritanceFlags]::None
$propagationFlag = [System.Security.AccessControl.PropagationFlags]::None
$objType = [System.Security.AccessControl.AccessControlType]::Allow
$objUser = New-Object System.Security.Principal.NTAccount($username)

$objACE = New-Object System.Security.AccessControl.FileSystemAccessRule($objUser, $colRights, $inheritanceFlag, $propagationFlag, $objType)
$objACL = Get-Acl -Path $path

$objACL.AddAccessRule($objACE)
$objACL.SetAccessRuleProtection($true, $false)

Set-Acl -Path $path -AclObject $objACL

explorer.exe "/Select,$path"      # Note: have no blank or space between "," for the arguments

# Once completed, the script opens the test file in the File Explorer and selects it.
# You can then right-click the file and choose Properties > Security to view the new settings.

# To find out the available access rights, in the ISE editor type in this line:    [System.Security.AccessControl.FileSystemRights]::    

###########################################################################################################################################################################################

# Tip 94: Start to Look at DSC


# Desired State Configuration (DSC) is a new feature in PowerShell 4.0. 
# With DSC, you can write simple configuration scripts and apply them to the local or a remote machine. 

# Here is a sample script to get you started:

Configuration MyConfig
{
    param($MachineName)  # Parameters are optional

    Node $MachineName    # A Configuration block can have one or more Node blocks
    {
        Registry RegistryExample
        {
            Ensure = "Present"              # You can also set Ensure to "Absent"
            Key = "HKEY_LOCAL_MACHINE\SOFTWARE\ExampleKey"
            ValueName = "TestValue"
            ValueDate = "TestData"
        }
    }
}

MyConfig -Machine $env:COMPUTERNAME -OutputPath C:\dsc
Start-DscConfiguration -Path c:\dsc -Wait 

# The configuration "MyConfig" uses the resource "Registry" to make sure that a given Registry key is present. 
# There are many more resources you can use in your DSC script, like adding (or removing) a local user or files, 
# unpacking an MSI package or ZIP file, or starting or stopping a service, to name just a few.

# Running the configuration will only create a MOF file. To apply the MOF file, use the Start-DSCConfiguration cmdlet. 
# Use -Wait to wait for the configuration to take place. Else, the configuration will be done in the background using a job.

###########################################################################################################################################################################################

# Tip 95: Checking Windows Updates


# To check all installed updates on a Windows box, there is a COM library you can use. 
# Unfortunately, this library isn't very intuitive to use, nor does it work remotely.

# So here is a PowerShell function called Get-WindowsUpdate. 
# It gets the locally installed updates by default, but you can also specify one or more remote machines and then retrieve their updates.

# Remote access is done through PowerShell remoting, so it will work only if the remote machine has PowerShell remoting enabled 
# (Windows Server 2012 enables PS remoting by default, for example), and you need to have local Administrator privileges on the remote machine.

function Get-WindowsUpdate
{
    [CmdletBinding()]
    param
    (
        [String[]]
        $ComputerName,
        $Title = "*",
        $Description = "*",
        $Operation = "*"
    )

    $code = {
    
        param($Title, $Description)

        $Type = @{
        
            Name = "Operation"
            Expression = {
            
                switch($_.Operation)
                {
                    1 {"Installed"}
                    2 {"Uninstalled"}
                    3 {"Other"}
                }
            }       
        }
          
        $session = New-Object -ComObject "Microsoft.Update.Session"
        $searcher = $session.CreateUpdateSearcher()
        $historyCount = $searcher.GetTotalHistoryCount()
        
        $searcher.QueryHistory(0, $historyCount) | Select-Object Title,Description,Date,$type | 
            Where-Object { $_.Title -like $Title } |
            Where-Object { $_.Description -like $Description } | 
            Where-Object { $_.Operation -like $Operation }
    }

    $null = $PSBoundParameters.Remove("Title")
    $null = $PSBoundParameters.Remove("Description")
    $null = $PSBoundParameters.Remove("Operation")

    Invoke-Command -ScriptBlock $code @PSBoundParameters -ArgumentList $Title, $Description
}

Get-WindowsUpdate -Title *Office* -Operation Installed | Out-GridView

# Note: Looks following will return nothing, bug here?
Get-WindowsUpdate -ComputerName $env:COMPUTERNAME -Title *Office* -Operation Installed | Out-GridView
Get-WindowsUpdate -ComputerName IIS-CTI5052 -Title *Office* -Operation Installed | Out-GridView

# Reference site: http://powershell.com/cs/blogs/tips/archive/2014/05/26/checking-windows-updates.aspx

###########################################################################################################################################################################################

# Tip 96: PowerShell God Mode


# Before you can run a PowerShell script, the execution policy needs to allow this. 
# Typically, you would use this line to enable script execution:

Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned


# However, if group policy has disabled script execution, then this line will not do you any good. 
# In this case, you can re-enable script execution with this code (per PowerShell session):

$context = $executioncontext.GetType().GetField('_context','nonpublic,instance').GetValue($executioncontext)
$field = $context.GetType().GetField('_authorizationManager','nonpublic,instance')
$field.SetValue($context,(New-Object Management.Automation.AuthorizationManager 'Microsoft.PowerShell'))

# Note that this is a hack, effectively resetting the authorization manager, 
# which may or may not have other side effects. Use at your own risk.


# This technique is not a security issue by the way. Execution policy generally is not a security boundary. 
# It’s not designed to keep bad people away. It is solely meant to protect you from yourself.
# So whether you enable script execution via cmdlet or via this code, 
# you are in both cases consenting to taking the responsibility of executing PowerShell code into your own hands.

###########################################################################################################################################################################################

# Tip 97: Removing Selected NTFS Permissions


# Maybe you need to remove some permission settings from NTFS permissions. 
# Let's assume you want to remove all permissions for a specific user because the user left the department.

# Note: Of course you can manage NTFS permissions per group, and setting permissions per user is typically not a good idea. 
# Still, often permissions are set per user, and the following example script can not only remove such permissions 
# but with minor adjustments also be used as an audit tool to find such permissions.

# Here is a simple example script. Adjust $Path and $Filter. 
# The script will then scan the folder $Path and all of its subfolders for access control entries 
# that match the $Filter string. It will only process non-inherited ACEs.

# The output states in red the ACEs that will be removed, and in green all ACEs that do not match the filter. 
# If the script does not return anything, then there are no direct ACEs in the folder you scanned.

$path = "C:\somfolder"
$filter = "S-1-5-*"

Get-ChildItem -Path C:\Obfuscated -Recurse -ErrorAction SilentlyContinue | ForEach-Object {

    $ACL = Get-Acl -Path $path
    
    $found = $false
    foreach($acc in $ACL.Access)
    {
        if($acc.IsInherited -eq $false)
        {
            $value = $acc.IdentityReference.Value

            if($value -like $filter)
            {
                Write-Host "Remove $value from $path " -ForegroundColor Red

                $null = $ACL.RemoveAccessRule($acc)
                $found = $true
            }
            else
            {
                Write-Host "Skipped $value from $path " -ForegroundColor Green
            }
        }
    }

    if($found)
    {
        Set-Acl -Path $Path -AclObject $ACL -ErrorAction Stop      
    }
}

###########################################################################################################################################################################################

# Tip 98: Setting Registry Permissions


# Setting permissions for Registry keys isn't trivial. With a little trick, though, it is no big deal anymore.

# First, open REGEDIT and create a sample key. Next, right click the key and use the UI to set the permissions you want.

# Now, run this script (adjust the path to the Registry key you just defined):

$path = "HKCU:\silencetest"
$sd = Get-Acl -Path $path
$sd.Sddl | clip

# It will read the security information from your key and copies it to the clipboard.

# Now, use this script to apply the exact same security settings to any new or existing Registry key you want. 
# Simply select the SDDL definition in this script, and replace it with the one you just created:

$sddl = "O:BAG:DUD:(A;OICIID;KA;;;S-1-5-21-2146773085-903363285-719344707-1385424)(A;OICIID;KA;;;SY)(A;OICIID;KA;;;BA)(A;OICIID;KR;;;RC)"

$path = "HKCU:\silencetest2"
$null = New-Item -Path $path -ErrorAction SilentlyContinue

$sd = Get-Acl -Path $path
$sd.SetSecurityDescriptorSddlForm($sddl)
Set-Acl -Path $path -AclObject $sd

# You may need to run this script with full Administrator privileges. 
# As you can see, the first script and your sample Registry key were only needed to generate the SDDL text. 
# Once you have it, you simply paste it into the second script. The second script does not need any sample key anymore.

###########################################################################################################################################################################################

# Tip 99: Getting Group Membership - Fast


[System.Security.Principal.WindowsIdentity]::GetCurrent()

# Outout:
#        AuthenticationType : Kerberos
#        ImpersonationLevel : None
#        IsAuthenticated    : True
#        IsGuest            : False
#        IsSystem           : False
#        IsAnonymous        : False
#        Name               : domain\sihe
#        Owner              : S-1-5-...
#        User               : S-1-5-...
#        Groups             : {S-1-5-...}
#        Token              : 4812
#        UserClaims         : {http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name:...}
#        DeviceClaims       : {}
#        Claims             : {http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name:...}
#        Actor              : 
#        BootstrapContext   : 
#        Label              : 
#        NameClaimType      : http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name
#        RoleClaimType      : http://schemas.microsoft.com/ws/2008/06/identity/claims/groupsid

[System.Security.Principal.WindowsIdentity]::GetCurrent().Groups.Value | ForEach-Object {

    $sid = $_

    $objSID = New-Object System.Security.Principal.SecurityIdentifier($sid)

    $objUser = $objSID.Translate([System.Security.Principal.NTAccount])

    $objUser.Value + "   =   " + $sid
}


# Enable PowerShell remoting, use Invoke-Command and put all that code in a script block as an argument for the -ScriptBlock parameter.

###########################################################################################################################################################################################

# Tip 100: Submitting Parameters through Splatting


# Splatting was introduced in PowerShell 3.0, but many users still never heard of this. 
# It is a technique to programmatically submit parameters to a cmdlet. Have a look:

$infos = @{}
$infos.Path = "C:\Windows"
$infos.Recures = $true
$infos.Filter = "*.log"
$infos.ErrorAction = "SilentlyContinue"
$infos.Remove("Recures")

dir @infos

dir -Path C:\Windows -Filter "*.log" -ErrorAction SilentlyContinue

# This example defines a hash table with key-value pairs. 
# Each key corresponds to a parameter found in the dir command, 
# and each value is the argument that should be submitted to that parameter.

# Splatting can be extremely useful if your code needs to decide which parameters should be forwarded to a given cmdlet. 
# Your code would then simply manipulate a hash table, then submit it to the cmdlet of choice.

###########################################################################################################################################################################################