# Reference site: http://powershell.com/cs/blogs/tips/
###########################################################################################################################################################################################

# Tip 1: Extract Paths from Strings Like Environment Variables

$env:Path

$env:Path += ";C:\newpath"                                         # Add C:\newpath to $env:path

$env:Path

(($env:Path -split ";") -ne "C:\Windows\System32") -join ";"       # Remove C:\Windows\System32 from $env:path

$env:Path -split ";"                                               # Show all env path in a friendly and clear way

###########################################################################################################################################################################################

# Tip 2: Copying Large Files with BITS

# You can use Copy-Item or the console applications xcopy and robocopy to copy large files. But did you know that you can also use the BITS service to do this? 
# The BITS service is supposed to download large update files and does so in a very robust way. It can download files across multiple restarts. BITS can copy local files as well.

Import-Module BitsTranser

$source = "C:\Images\*.iso"
$destination = "C:\BackupFolder\"

if(-not (Test-Path $destination))
{
    $null = New-Item -Path $destination -ItemType Directory
}

Start-BitsTransfer -Source $source -Destination $destination -Description "Backup" -DisplayName "Backup" 

# A couple of things to note: BitsTransfer cannot copy recursively, and it cannot copy files that are currently in use by other programs. 
# It does display a neat progress bar, though, and it is designed to copy large files in a robust way.

###########################################################################################################################################################################################

# Tip 3: Using Splatting for Better Formatting

Get-ChildItem -Path $env:windir -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue -Force


Get-ChildItem -Path $env:windir `
-Filter *.ps1 -Recurse `
-ErrorAction SilentlyContinue -Force


$MyParameter = @{

    Path = "$env:windir"
    filter = "*.ps1"
    Recurse = $true
    ErrorAction = "SilentlyContinue"
    Force = $true
}

Get-ChildItem @MyParameter                                                           # Note: should be @MyParameter here but not $MyParameter

# As a nice side effect, you can submit your set of parameters defined in $MyParameter multiple times. 
# Splatting consists of a hash table ($MyParameter) that is submitted via '@' to a cmdlet. Each key in your hash table turns into a parameter, 
# and each value becomes the parameter argument. For switch parameters like -Force or -Recurse, you explicitly set the key to $true or $false.

###########################################################################################################################################################################################

# Tip 4: Converting Letters to Numbers and Bit Masks

#　Sometimes, you may need to turn characters such as drive letters into numbers or even bit masks. With a little bit of math, this is easily doable. 

$driverList = 'a', 'b:', 'd', 'Z', 'x:'

$driverList | ForEach-Object {$_.toUpper()[0]} | Sort-Object        　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　# sort and normalize a dirty and unsorted list of drive letters
# Output:
#         A
#         B
#         D
#         X
#         Z

$driverList | ForEach-Object {$_.toUpper()[0]} | Sort-Object | ForEach-Object {([byte]$_) - 65} 　# get back drive indices, taking advantage of the ASCII code of characters
# Output:
#         0
#         1
#         3
#         23
#         25

$driverList | ForEach-Object {$_.toUpper()[0]} | Sort-Object | ForEach-Object {[Math]::Pow(2, (([byte]$_) - 65))} 　　　#To turn this into a bit mask, use the Pow() function
# Output:
#         1
#         2
#         8
#         8388608
#         33554432

###########################################################################################################################################################################################

# Tip 5: Hiding Drive Letters

function Hide-Drive($dirveLetter)
{
    $key = @{

        Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer'
        Name = 'NoDrives'
    }

    if($dirveLetter -eq $null)
    {
        Remove-ItemProperty @key
    }
    else
    {
        $mask = 0
        $dirveLetter | ForEach-Object {$_.toUpper()[0]} | Sort-Object | ForEach-Object {$mask += [Math]::Pow(2, (([byte]$_) - 65))}

        Set-ItemProperty @key -Value $mask -type DWORD
    }
}

Hide-Drive A,B,E,Z                                  # to hide drives A, B, E, and Z, a user can still open files and folders on hidden drives if he knows the path.
                                                    
Hide-Drive                                          # To display all drives, omit arguments

# Note that you need to have administrative privileges to change policies, and that policies may be overridden by group policy settings set up by your corporate IT.
# For the changes to take effect, you need to log off and on again or kill your explorer.exe and restart it.

###########################################################################################################################################################################################

# Tip 6: Locking Drive Content

# Below function will not hide drive letters but prohibit access to drive content. You need administrative privileges to set this setting. 
function Hide-DriveContent($driveLetter)
{
    $key = @{

        Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer'
        Name = 'NoViewOnDrive'
    }
    
    if($driveLetter -eq $null)
    {
        Remove-ItemProperty @key
    } 
    else
    {
        $mask = 0
        $driveLetter | ForEach-Object {$_.toUpper()[0]} | Sort-Object | ForEach-Object {$mask += [Math]::Pow(2, (([byte]$_) - 65))}

        Set-ItemProperty @key -Value $mask -type DWORD
    }
}

Hide-DriveContent D,Z                                                      # prohibits access to drives D: and Z:
Hide-DriveContent                                                          # removes all restrictions

# Note: Log off and on again, or kill explorer.exe for changes to take effect.

###########################################################################################################################################################################################

# Tip 7: No Reboots After Updates

# If you have set Windows Update to automatic mode, it takes care of detecting, downloading, and installing all necessary updates - fine. 
# However, it also automatically takes care of rebooting if required. 

$key = @{

    Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate\AU'
    Name = 'NoAutoRebootWithLoggedOnUsers'
}

Set-ItemProperty @key -Value 1 -type DWORD                                   # To turn off automatic rebooting while a user is logged on 


# To change it back to defaults, use either one of these:
Set-ItemProperty @key -Value 0 -type DWORD
Remove-ItemProperty @key

# Note that this code is a generic template and illustrates how you can set or remove any Registry value easily. 
# Of course, you may need admin privileges to do the change, and you should know exactly what you are doing because messing up the Registry can destroy your Windows installation.

###########################################################################################################################################################################################

# Tip 8: Implicit Foreach in PSv3

# When you work with array data, in previous PowerShell versions you had to loop through all elements to get to their properties and methods. No more in PowerShell v3. 
# PowerShell now detects that you are working with an array and applies your call to all array elements. So with this line you could gracefully close all running notepad.exe. 

(Get-Process notepad).CloseMainWindow()                                      # gracefully close all running notepad.exe
                                                                             
Get-Process notepad | ForEach-Object {$_.CloseMainWindow()}                  # With previous powershell version, you should do like this

###########################################################################################################################################################################################

# Tip 9: Running a Script Block with Parameters

# To run a submitted script block, use Invoke-Command. This cmdlet can run script blocks locally and remote, and it has a built-in mechanism to supply parameters to the script block.

$code = {

    param($name)

    Get-EventLog -LogName $name -EntryType Error
}

Invoke-Command -ScriptBlock $code -ArgumentList System


# On powershell V3, you can:
$name = "System"
Invoke-Command -ScriptBlock {Get-EventLog -LogName $using:name -EntryType Error} -ArgumentList $name   # Exception here because using only accessable on remote machine

# PowerShell 3.0中新增了前缀using，可以标记本地变量，让它提供远程支持。这样一来远程执行可以顺利过关, 但是现在在本地执行又会失败。 
# 因为 “using” 只有远程执行时才会识别。针对这个问题有一个解决方案是将本地变量已参数列表的形式传递给远程会话，下面的修改的代码终于可以支持本地执行和也可兼容远程执行。

function Get-Log($logName = "System", $computerName = $env:COMPUTERNAME)
{
    $code = {param($logName) Get-EventLog -LogName $logName -EntryType Error -Newest 5}
    $null = $PSBoundParameters.Remove("logName")
    Invoke-Command -ScriptBlock $code.GetNewClosure() @PSBoundParameters -ArgumentList $logName
} 

Get-Log
Get-Log -computerName IIS-CTI5052

# About ScriptBlock.GetNewClosure() Behavior: http://stackoverflow.com/questions/4058721/scriptblock-getnewclosure-behavior

###########################################################################################################################################################################################

# Tip 10: Check PowerShell Speed

# Method One:
function Test
{
    $codeText = $args -join " "
    $codeText = $ExecutionContext.InvokeCommand.ExpandString($codeText)
    
    $code = [ScriptBlock]::Create($codeText)

    $timeSpan = Measure-Command $code

    "Your code took {0:0.000} secondes to run" -f $timeSpan.TotalSeconds
}

Test Get-Service
Test Get-WmiObject Win32_Service
Test dir $home -Include *.ps1 -Recurse
# Note: To test drive more complex commands, make sure you place them in quotes.



# Method Two:
function Test-CmdSpeed([ScriptBlock]$command)
{
    $measure = Measure-Command {Invoke-Command -Command $command}

    New-Object PSObject -Property @{
    
        Command = $command
        Seconds = $measure.TotalSeconds
        Milliseconds = $measure.TotalMilliseconds
    }
}

Test-CmdSpeed {dir} | Format-List
# Output:
#        Command      : dir
#        Milliseconds : 221.3807
#        Seconds      : 0.2213807



# Method Three:
function Test-CmdSpeed([ScriptBlock]$command)
{
    ((Measure-Command $command).TotalSeconds).toString("0.000")
}

Test-CmdSpeed {dir}

###########################################################################################################################################################################################

# Tip 11: Synchronizing Current Folder

[IO.Path]::GetFullPath(".")                                                        # Output: C:\Windows\system32
cd $env:windir                                                                     # PS C:\Windows> 
Get-Location                                                                       # Path: C:\Windows
[IO.Path]::GetFullPath(".")                                                        # Output: C:\Windows\system32


# whenever you want to use one of the .NET Framework methods from IO.Path, you need to sync the current location 

[IO.Path]::GetFullPath(".")                                                        # Output: C:\Windows\system32
cd $env:windir                                                                     # PS C:\Windows> 

[IO.Directory]::SetCurrentDirectory((Get-Location -PSProvider FileSystem).ProviderPath)

Get-Location                                                                       # Path: C:\Windows
[IO.Path]::GetFullPath(".")                                                        # C:\Windows

###########################################################################################################################################################################################

# Tip 12: Resolving Paths

# Paths can be relative, such as ". \file.txt". To resolve such a path and display its full path, you could use Resolve-Path:

Resolve-Path .\file.txt    # Exception here: Resolve-Path : Cannot find path 'C:\Windows\file.txt' because it does not exist.

# Method One:
$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(".\file.txt")          # Output: C:\Windows\file.txt

# Method Two:
[IO.Path]::GetFullPath(".\file.txt")                                                           # Output: C:\Windows\file.txt

###########################################################################################################################################################################################

# Tip 13: Checking User Privileges

whoami /All /FO CSV | ConvertFrom-Csv | Sort-Object UserName


# There are numerous ways to find out if a script runs elevated. 
(whoami /all | Select-String "S-1-16-12288") -ne $null                                         # Output: True

###########################################################################################################################################################################################

# Tip 14: Finding Files Owned by a User

filter Get-Owner($account = "$env:UserDomain\$env:UserName")       # filter that will show only those files and folders that a specific user is owner of:
{
    if((Get-Acl $_.FullName).Owner -like $account)
    {
        $_
    }
}

dir $home | Get-Owner

# You can specify a specific user name (including domain part) or use wildcards as well:
dir $home | Get-Owner *system*

###########################################################################################################################################################################################

# Tip 15: Lunch Time Alert

# Here's a fun prompt function that turns your input prompt into a short prompt and displays the current path in your PowerShell window title bar. 
# In addition, it has a lunch count down, displaying the minutes to go. 
# Three minutes before lunch time, the prompt also emits a beep tone so you never miss the perfect time for lunch anymore.

function prompt
{
    $lunchTime = Get-Date -Hour 11 -Minute 30
    $timeSpan = New-TimeSpan -End $lunchTime
    [int]$minutes = $timeSpan.TotalMinutes

    switch($minutes)
    {
        { $_ -lt 0 }   { $text = 'Lunch is over. {0}'; break }
        { $_ -lt 3 }   { $text = 'Prepare for lunch!  {0}'; break }
        default        { $text = '{1} minutes to go... {0}'; break }
    }

    "PS> "

    $host.UI.RawUI.WindowTitle = $text -f (Get-Location), $minutes

    if($minutes -lt 3)
    {
        [System.Console]::Beep()
    }
}

###########################################################################################################################################################################################

# Tip 16: Easier ForEach/Where-Object in PSv3

# using Where-Object and ForEach-Object becomes a lot simpler in PSV3. No longer do you need a script block and code.
Get-ChildItem $env:windir | Where-Object Length -GT 1MB
Get-ChildItem $env:windir | ForEach-Object Length


# Previously, you would have had to write:
Get-ChildItem $env:windir | Where-Object {$_.Length -gt 1MB}

###########################################################################################################################################################################################

# Tip 17: Stopping Services Remotely

# Stop-Service cannot stop services remotely. One easy way of doing so is Set-Service:
Set-Service -Name Spooler -Status Stopped -ComputerName remoteMachine

# However, unlike Stop-Service, Set-Service has no -Force parameter, so you cannot stop services if they, for example, have running dependent services.

# If your infrastructure supports PowerShell Remoting, you could use Invoke-Command instead:
Invoke-Command {Stop-Service -Name Spooler -Force} -ComputerName remoteMachie

###########################################################################################################################################################################################

# Tip 18: Using Advanced Breakpoints

# PowerShell supports dynamic breakpoints. They trigger when certain requirements are met. Like regular breakpoints, they all require that your script has been saved to a file.

# set a breakpoint for script c:\test\script.ps1 that always triggers when the script accesses the variable $Server, read or write: 
Set-PSBreakpoint -Script "c:\test\test.ps1" -Variable Server -Mode ReadWrite

###########################################################################################################################################################################################

# Tip 19: Finding Built-In Administrators Group

# Using System User or group names like 'Administrators' in scripts may not always be a good idea 
# because they are localized and may not work on machines that use a different UI language.

$id = [System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid
$Account = New-Object System.Security.Principal.SecurityIdentifier($id, $null)
$Account.Translate([System.Security.Principal.NTAccount]).Value

[System.Enum]::GetNames("System.Security.Principal.WellKnownSidType")        # Get useful info form [System.Security.Principal.WellKnownSidType]
# Output:
#        RemoteLogonIdSid
#        LogonIdsSid
#        LocalSystemSid
#        LocalServiceSid
#        BuiltinDomainSid
#        BuiltinAdministratorsSid
#        AccountAdministratorSid
#        AccountGuestSid
#        AccountDomainAdminsSid
#        AccountDomainUsersSid
#        AccountDomainGuestsSid
#        AccountComputersSid
#        AccountControllersSid
#        AccountCertAdminsSid
#        ...


###########################################################################################################################################################################################

# Tip 20: Making netstat.exe Object-Oriented

netstat -an | ForEach-Object {
    
    $i = $_ | Select-Object -Property Protocol, Source, Destination, Mode

    $null, $i.Protocol, $i.Source, $i.Destination, $i.Mode = ($_ -split "\s{2,}")

    if($i.Protocol.Length -eq 3)                                                # filter the null value or invalid protocol like "Proto"
    {
      $i
    }
} | Out-File C:\Users\v-sihe\silencetest\2.txt

# Since the result is converted to objects, you can now easily use other PowerShell cmdlets to sort or filter the results.
netstat -an | ForEach-Object {
    
    $i = $_ | Select-Object -Property Protocol, Source, Destination, Mode

    $null, $i.Protocol, $i.Source, $i.Destination, $i.Mode = ($_ -split "\s{2,}")

    if($i.Protocol.Length -eq 3)
    {
        $i
    }

} | Where-Object {$_.Mode -eq "LISTENING"}

# Like always with console applications, the data may be localized and different on different UI cultures.




# The similer with command: route print -4

$ip = '\b(([01]?\d?\d|2[0-4]\d|25[0-5])\.){3}([01]?\d?\d|2[0-4]\d|25[0-5])\b'

route print -4 | foreach-object { 

     $i = $_ | Select -Property Destination , Netmask , Gateway , Interface, Metric

     $null, $i.destination, $i.netmask, $i.gateway, $i.interface, $i.metric= ($_ -split '\s{2,}')

     if ($i.Destination -match $ip) 
     { 
         $i 
     }

} | Format-Table

###########################################################################################################################################################################################

# Tip 21: Jagged Arrays

$array = 1,2,3,(1,('a','b'),3),5        # This may not be for everyone: have a look at how you can create "jagged arrays". 
$array

$array[2]                                                      # Output: 3
                                                              
$array[3]                                                     
# Output:                                                     
#        1                                                    
#        a                                                    
#        b                                                    
#        3                                                    
                                                              
$array[3][0]                                                   # Output: 1
$array[3][1]                                                   # Output: a,b
$array[3][1][0]                                                # Output: a

$array = "hello",(Get-Date),(Get-Process),(Get-ChildItem)
$array[0]
$array[1]
$array[2] -is [array]                                          # Output: True
$array[2][0]

###########################################################################################################################################################################################

# Tip 22: Check Out Nested Hash Tables

# Nested hash tables may look even more confusing but can be highly useful. 

$hash = @{
    
    id = 1
    age = 99
    group = "PS Programmers"
}

$person = @{
    
    name = "Silence"
    gender = "Male"
    more = $hash
}

$person
# Output:
#        Name                           Value                                                                                                                                                                       
#        ----                           -----                                                                                                                                                                       
#        name                           Silence                                                                                                                                                                     
#        more                           {group, age, id}                                                                                                                                                            
#        gender                         Male 


$person.name                                               # Output: Silence

$person.more
# Output:
#        Name                           Value                                                                                                                                                                       
#        ----                           -----                                                                                                                                                                       
#        group                          PS Programmers                                                                                                                                                              
#        age                            99                                                                                                                                                                          
#        id                             1

$person.more.group                                         # Output: PS Programmers
$person.more.id                                            # Output: 1

###########################################################################################################################################################################################

# Tip 23: Copying Command History to Clipboard

Get-History | Select-Object -ExpandProperty CommandLine | clip                    # Copy all commands from your command history to the clipboard

# Note that clip.exe is supported on Windows 7 and above only

###########################################################################################################################################################################################

# Tip 24: Creating New Scripts in ISE

# Takes all your interactive commands entered in the PowerShell ISE editor and opens them all as a new script inside ISE.
function New-Script
{
    $text = Get-History | Select-Object -ExpandProperty CommandLine | Out-String
    $file = $psISE.CurrentPowerShellTab.Files.Add()
    $file.Editor.Text = $text
}


# Note that this function runs only inside PowerShell ISE. Note also that you can "extend" the command line buffer so ISE remembers more commands:
$MaximumHistoryCount = 20000

###########################################################################################################################################################################################

# Tip 25: Finding Current Script Paths

# Method One:
function Get-ScriptDirectory
{
    $invocation = (Get-Variable MyInvocation -Scope 1).Value

    try
    {
        Split-Path $invocation.MyCommand.Path -ErrorAction SilentlyContinue
    }
    catch
    {
        Write-Warning 'You need to call this function from within a saved script.'
    }
}

Get-ScriptDirectory 




# Method Two:
Split-Path -Parent $MyInvocation.MyCommand.Definition

###########################################################################################################################################################################################

# Tip 26: Finding True WMI Properties

# get you all BIOS information including the PowerShell-added properties
Get-WmiObject Win32_BIOS | Select-Object * | Out-GridView          




# limits output to only the true WMI properties
$wmiObject = Get-WmiObject Win32_BIOS

$props = $wmiObject | Get-Member -MemberType Property | Select-Object -ExpandProperty Name | Where-Object { -not $_.StartsWith("__") }

$wmiObject | Select-Object -Property $props | Out-GridView

###########################################################################################################################################################################################

# Tip 27: Finding All Object Properties

Get-WmiObject Win32_Bios
# Output:
#         SMBIOSBIOSVersion : 786G6 v01.03
#         Manufacturer      : Hewlett-Packard
#         Name              : Default System BIOS
#         SerialNumber      : CNG9460LHL
#         Version           : HPQOEM - 20090825


Get-WmiObject Win32_Bios | Select-Object * -First 1

# The Select-Object statement selects all available columns and returns only the first element—just in case the initial command returned more than one object. 
# Since you only need one sample object to investigate the available properties, this is a good idea to not get overwhelmed with redundant data.


# To check out the formal data types and also methods present in objects, pipe the result to Get-Member instead:
Get-WmiObject Win32_Bios | Get-Member

###########################################################################################################################################################################################

# Tip 28: Investigating USB Drive Usage

# Dump the USB storage history from your registry and check which devices were used in the past

$key = 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\USBSTOR\*\*'

Get-ItemProperty $key | Select-Object -ExpandProperty FriendlyName | Sort-Object

###########################################################################################################################################################################################

# Tip 29: Discarding Unwanted Information

# If you want to dump results from a command, there are a number of ways. While they all do the same, they have tremendous performance differences:

(Measure-Command { 1..100000 | Out-Null }).TotalMilliseconds              # Output: 1156.5148

(Measure-Command { $null = 1..100000 }).TotalMilliseconds                 # Output: 168.2157

(Measure-Command { 1..100000 > $null }).TotalMilliseconds                 # Output: 15.5189

(Measure-Command { [void](1..100000) }).TotalMilliseconds                 # Output: 5.0103


# Note: As you can see, piping results to Out-Null is the worst approach. Instead, simply assign unwanted results to the special variable $null or cast the call to [void]. 

###########################################################################################################################################################################################

# Tip 30: Create System Restore Points 

# Provided system restore points are enabled on your machine and your script has admin privileges, 
# PowerShell can easily create restore points and restore the system if anything went bad. 
# Note though some of the limitations: Restore points only work past Windows XP. They only work for clients, not servers. And you can only create one restore point every 24hrs.

Checkpoint-Computer -Description "Before script execution"                                      # Create System Restore Points

Get-ComputerRestorePoint                                                                        # Returns all the restore points available on your system

$id = Get-ComputerRestorePoint | Where-Object {$_.Description -eq "Before script execution"}    # Filter out your new restore point according to its description

Restore-Computer $id -WhatIf                                                                    # Restore system with the specified restore points




# Enabling and Disabling Computer Restore

# you can use PowerShell to create restore points and restore your system state in case something went bad. 
# There are two more cmdlets that control which drives are monitored by system restore points: Enable-ComputerRestore and Disable-ComputerRestore.

Enable-ComputerRestore -Drive "C:\", "D:\"




# Verifying Restore Points

# When you create a new restore point with Checkpoint-Computer, you do not get back any feedback telling you whether the operation succeeded. 
# Fortunately, you can easily lookup the corresponding event log entry:

Get-EventLog -LogName Application -InstanceId 8194 -Newest 1 | Select-Object *
# Output:
#         EventID            : 8194
#         MachineName        : xx.xxx.xxxxx
#         Data               : {0, 0, 0, 0...}
#         Index              : 436806
#         Category           : (0)
#         CategoryNumber     : 0
#         EntryType          : Information
#         Message            : Successfully created restore point (Process = C:\Windows\system32\svchost.exe -k netsvcs; Description = Windows Update).
#         Source             : System Restore
#         ReplacementStrings : {C:\Windows\system32\svchost.exe -k netsvcs, Windows Update}
#         InstanceId         : 8194
#         TimeGenerated      : 2014/09/22 09:13:15
#         TimeWritten        : 2014/09/22 09:13:15
#         UserName           : 
#         Site               : 
#         Container          :

###########################################################################################################################################################################################

# Tip 31: Rename PowerShell Scripts

# finds all PowerShell script files in your home folder and checks whether they contain the phrase "Untitled" followed by a number. 
# If so, the name is replaced by that number plus "_Untitled Script - Check":

dir $home *.ps1 -Recurse | Rename-Item -NewName { $_.Name -replace 'Untitled(\d{1,})', '$1_Untitled Script - Check' }



# About $1, $2, $3 in regex expression, see details: http://www.pstips.net/regex-back-reference.html

###########################################################################################################################################################################################

# Tip 32: Adding Type Accelerators

# Type accelerators are shortcut names that represent .NET types. For example, you can use [XML] instead of [System.Xml.XmlDocument]. 
# By default, PowerShell only contains type accelerators for the most common .NET types but you can easily add more.

# Adds a type [SuperArray] which is a shortcut to [System.Collections.ArrayList] and creates a super-charged array: 
# it has methods to insert and add new elements, speeding up array-manipulation tremendously:

$accelerator = [type]::GetType("System.Management.Automation.TypeAccelerators")
$accelerator::Add("SuperArray", [System.Collections.ArrayList])                    # # Add "Alias" for .Net types


$arr = [SuperArray](1..5)
$arr.Insert(1, 12)


# Note that this trick will no longer work in PowerShell v3 unfortunately. It will get exception on PS V3.

###########################################################################################################################################################################################

# Tip 32: Identifying PowerShell Host: powershell_ise or powershell

# If your script requires a real console, or if your script requires PowerShell ISE features, it may be a wise thing to check which host is actually running your script.

# Method One:

(Get-Process -Id $pid).ProcessName
# PowerShell ISE:                  powershell_ise
# PowerShell Console:              powershell




# Method Two:

(Get-Host).Name
# PowerShell ISE:                  Windows PowerShell ISE Host
# PowerShell Console:              ConsoleHost

###########################################################################################################################################################################################

# Tip 33: PowerShell ISE v3 Keyboard Shortcuts

<#
    PowerShell v3 comes with a new PowerShell ISE script editor. Here are some of the most useful keyboard shortcuts:

    
    CTRL+N:                 creates a new script
                            
    CTRL+O:                 opens an existing script
                            
    CTRL+T:                 adds a new PowerShell tab
                            
    CTRL+SHIFT+R:           opens a remote PowerShell tab                                                  <=> Enter-PSSession -ComputerName remoteMachine
                            
    F1:                     gets help for the cmdlet that the cursor is on. 

    F3:                     finds next occurrence
                            
    CTRL+F1:                opens input assistant to provide arguments for the cmdlet the cursor is on     <=> Show-Command
    
    SHIFT+F3:               finds previous occurrence
                            
    CTRL+J:                 opens code snippet list to quickly insert predefined code blocks
                            
    CTRL+M:                 toggles regions
                            
    CTRL+F:                 finds text in a script                       
                                                     
    CTRL+H:                 replaces text
                            
    CTRL+G:                 go to line
                            
    CTRL+U:                 to lowercase
                            
    CTRL+SHIFT+U:           to uppercase
                            
    CTRL+SPACE:             start IntelliSense
    
    CTRL+1, CTRL+2, CTRL+3: changes position of script pane
                  
    ALT+SHIFT+T：           可以调换相邻的两行代码交换位置
                            
    CTRL+S：                暂停输出，再按任意键恢复并继续输出
                            
    CTRL+I:                 设置焦点到脚本面板

    CTRL+R:                 隐藏或显示脚本面板
                            
    Shift+Alt+Arrow:        批量水平移动或注释Code
#>


###########################################################################################################################################################################################

# Tip 34: Getting Relative Dates

# Here's a quick and fast way of generating relative dates in any format:

(Get-Date).AddDays(-1).ToString("yyyy-MM-dd")                                  # Output: 2014-09-23

Get-EventLog -LogName System -EntryType Error -After (Get-Date).AddDays(-2)    # Returns all error events from the System event log in the past 48 hour

###########################################################################################################################################################################################

# Tip 35: Using Code Snippets in PowerShell ISE v3

$code = @"
    
    $ErrorActionPreference = 'Stop'
    
    trap
    {
        $Global:g = $_
        $script = $_.InvocationInfo.ScriptName
        $line = $_.InvocationInfo.ScriptLineNumber
        $message = $_.Exception.Message
    
        Write-Warning ('LINE {0:000} "{1}" Error: {2}' -f $line, $script, $message)
    
        Continue
    }    

"@

New-IseSnippet -Title "Trap" -Text $code -Description "Inserts a global error handler" -CaretOffset $code.Length -Force           # Create contents to Ctrl+J list


# Note: All custom snippets are stored as XML files in this location:
Join-Path (Split-Path $profile) 'Snippets'


<#
    Note how New-ISESnippet cmdlet sets the caret position to the length of the code snippet text. 
    By default, after you insert a snippet, the cursor is positioned right before the snippet. 
    If you'd rather want to position the cursor at the end of the code snippet, use the approach showed in this sample.
    
    The snippet name is determined by the parameter -Title. To insert the error handler code, 
    simply press CTRL+J and then type the snippet name, for example "trap". Then press ENTER.
    
    New-ISESnippet actually produces snippet files, so the snippets you define do persist. Close PowerShell ISE and reopen it again. 
    Then press CTRL+J. Your custom snippets are still available. Use the parameter -Force to overwrite and replace existing snippets.    
#>




# Removing Code Snippets in PowerShell ISE v3

Get-IseSnippet                        # Returns all snippets that you registered previously with New-ISESnippet, the build-in ise snippets not included here.
                                      
Get-IseSnippet | Remove-Item          # Remove all custom snippets


dir (Join-Path (Split-Path $profile) "Snippets") Tr* | Remove-Item                        # Filter the particular one and delete it.
# Output:
#             Directory: C:\Users\v-sihe\Documents\WindowsPowerShell\Snippets
#         
#         
#         Mode                LastWriteTime     Length Name                                                                                                                                                                    
#         ----                -------------     ------ ----                                                                                                                                                                    
#         -a---        2014/09/24  02:00 PM        957 Trap.snippets.ps1xml

###########################################################################################################################################################################################

# Tip 36: Using ActiveDirectory Module without AD Drive

# By default, when you import Microsoft’s ActiveDirectory PowerShell module which ships with Server 2008 R2 and is a part of the free RSAT tools, 
# it will import AD cmdlets and also install an AD: PowerShell drive.

# If you do not want to install that drive, set a special environment variable:

$env:ADPS_LoadDefaultDrive = 0

# Once this variable is present, importing the ActiveDirectory module will no longer auto-mount the AD: drive.

###########################################################################################################################################################################################

# Tip 37: Testing URLs for Proxy Bypass

function Test-ProxyBypass                          # To find out whether a given URL goes through a proxy or is accessed directly
{ 
    param
    (
        [Parameter(Mandatory = $true)]
        [String]$url
    )

    $webClient = New-Object System.Net.WebClient
    
    return $webClient.Proxy.IsBypassed($url)
}

# To test a URL, simply submit it to the function Test-ProxyBypass. The result is either $true or $false.

Test-ProxyBypass http://www.baidu.com                               # Output: True
Test-ProxyBypass http://www.microsoft.com                           # Output: True

###########################################################################################################################################################################################

# Tip 38: Customize PowerShell ISE Editor


# Controlling IntelliSense in PowerShell ISE v3

# The new PowerShell ISE script editor that ships with PowerShell v3 has a much improved IntelliSense (code completion) that you can even fine-tune from PowerShell. 

$psISE.Options.IntellisenseTimeoutInSeconds                         # Output: 3
$psISE.Options.ShowIntellisenseInConsolePane                        # Output: True
$psISE.Options.ShowIntellisenseInScriptPane                         # Output: True
$psISE.Options.UseEnterToSelectInConsolePaneIntellisense            # Output: True
$psISE.Options.UseEnterToSelectInScriptPaneIntellisense             # Output: True




# Colorizing PowerShell ISE v3

# The new PowerShell ISE script editor in PowerShell v3 lets you customize a lot of colors, so if a particular color does not show well on a projector, 
# for example, simply change it. You can do that via GUI, but you can also do it programmatically.

$psISE.Options.RestoreDefaultConsoleTokenColors 
$psISE.Options.RestoreDefaultTokenColors 
$psISE.Options.RestoreDefaultXmlTokenColors 
$psISE.Options.ConsolePaneBackgroundColor 
$psISE.Options.ConsolePaneForegroundColor 
$psISE.Options.ConsolePaneTextBackgroundColor 
$psISE.Options.ConsoleTokenColors 
$psISE.Options.DebugBackgroundColor 
$psISE.Options.DebugForegroundColor 
$psISE.Options.ErrorBackgroundColor 
$psISE.Options.ErrorForegroundColor 
$psISE.Options.ScriptPaneBackgroundColor 
$psISE.Options.ScriptPaneForegroundColor 
$psISE.Options.TokenColors 
$psISE.Options.VerboseBackgroundColor 
$psISE.Options.VerboseForegroundColor 
$psISE.Options.WarningBackgroundColor 
$psISE.Options.WarningForegroundColor 
$psISE.Options.XmlTokenColors 

###########################################################################################################################################################################################

# Tip 39: Disabling Console Apps in PowerShell ISE

# In PowerShell v3, the new PowerShell ISE script editor has improved a lot. Yet, it still has no real console but instead sports a useful console simulation. 
# Some applications do require a real console, though. If you run those in PowerShell ISE, the editor may become unresponsive.

choice.exe                          # This command halts PowerShell ISE because choice.exe waits for input it never receives

# To prevent this, PowerShell ISE maintains a list of unsupported console applications and won't run them. 
# The list is stored in the variable $psUnsupportedConsoleApplications (which does not exist in the regular PowerShell console).

$psUnsupportedConsoleApplications        # Get all commands that are not supported in powershell ISE
# Output:
#         wmic
#         wmic.exe
#         cmd
#         cmd.exe
#         diskpart
#         diskpart.exe
#         edit.com
#         netsh
#         netsh.exe
#         nslookup
#         nslookup.exe
#         powershell
#         powershell.exe

# Note: You can improve this list and add applications that you find won't run well in PowerShell ISE

$psUnsupportedConsoleApplications.Add("choice.exe")         # add choice.exe to the list
$psUnsupportedConsoleApplications                           # Check this variable you will find choice.exe was added to the list

choice.exe
# Output:
#        Cannot start "choice.exe". Interactive console applications are not supported. 
#        To run the application, use the Start-Process cmdlet or use "Start PowerShell.exe" from the File menu.
#        To view/modify the list of blocked console applications, use $psUnsupportedConsoleApplications, or consult online help.
#        At line:0 char:0

###########################################################################################################################################################################################

# Tip 40: Get-Content Can Now Read Raw Text

<#
    In PowerShell v3, Get-Content was reading text files line by line. 
    This was great for pipeline processing but could take a long time and also changed/removed the original line endings.

    
    In PowerShell v3, Get-Content now has a parameter -Raw. 
    When specified, the text file will be read in as one single string, keeping line endings exactly the way they were.
 
    
    Use -Raw for example if you want to read in a text file, change or replace parts and then write it back (for example, by using Set-Content).
#>




# Note: Before PS 3.0 you could easily get the same result with the following line:

Get-Content -Path MyFile.txt -Delimiter "`0"

###########################################################################################################################################################################################

# Tip 41: Out-GridView Grows Up

Get-Service | Out-GridView                         # Displaying results in a separate window then exit the command

Get-Service | Out-GridView -Wait                   # Displaying results in a separate window,  the script pauses until the window is closed.

Get-Service | Out-GridView -PassThru               # -PassThru will allow the user to send back his selection to the script. So now, Out-GridView can work as a selector. 
Get-Service | Out-GridView -OutputMode Multiple    # Basically, the parameter -PassThru is just a shortcut for -OutputMode Multiple.

Get-Service | Out-GridView -OutputMode Single      # -OutputMode Single allows the user to select only one service. By default, multi selection is enabled. 

# OutputMode Supported Parameters: Multiple, Single, None

###########################################################################################################################################################################################

# Tip 42: Restarting Computers in PowerShell

<#
    -Wait: Halts the script until the machine has rebooted
    
    -Timeout: Seconds to wait for the machine to restart
    
    -For: Considers the computer to have restarted when the specified resources are available. Valid values: WMI, WinRM, and PowerShell.
    
    -Delay: Interval in seconds used to query the remote computer to determine its availability specified by -For.
#>

Restart-Computer -ComputerName Server01 -Wait -For PowerShell -Timeout 300 -Delay 2
      
# This command restarts the Server01 remote computer and then waits up to 5 minutes (300 seconds) for Windows PowerShell to be available on the restarted computer before continuing.

# The command uses the Wait, For, and Timeout parameters to specify the conditions of the wait. 
# It uses the Delay parameter to reduce the interval between queries to the remote computer that determine whether it is restarted.

###########################################################################################################################################################################################

# Tip 43: Sophisticated Directory Filtering in PowerShell v3

# In PowerShell v3, Get-ChildItem now supports sophisticated filtering through its –Attribute parameter. 

# To get all files in your Windows folder or one of its subfolders that are not system files but are encrypted or compressed
Get-ChildItem $env:windir -Attributes !Directory+!System+Encrypted,!Directory+!System+Compressed -Recurse -ErrorAction SilentlyContinue     # Note how "!" negates the filter


# The -Attributes parameter supports these attributes: 
#    Archive, Compressed, Device, Directory, Encrypted, Hidden, Normal, NotContentIndexed, Offline, ReadOnly, ReparsePoint, SparseFile, System, and Temporary.

# Use the following operators to combine attributes:
#                                                      !          NOT
#                                                      +          AND
#                                                      ,          OR

# See details:
Get-Help Get-ChildItem -Parameter Attributes

###########################################################################################################################################################################################

# Tip 44: Finding Files Only or Folders Only

# In PowerShell v2, to list only files or only folders you had to do filtering yourself:

Get-ChildItem $env:windir | Where-Object {$_.PSIsContainer -eq $true}
Get-ChildItem $env:windir | Where-Object {$_.PSIsContainer -eq $false}




# In PowerShell v3, Get-ChildItem is smart enough to do that for you:

Get-ChildItem $env:windir -File 
Get-ChildItem $env:windir -Directory

###########################################################################################################################################################################################

# Tip 45: Use Select-Object's With -ExpandProperty

Get-Process | Select-Object -Property *                            # Use -Property * when you want to see maximum information

Get-Process | Select-Object -Property Name, Description, Company   # Use -Property a,b,c to select more than one column

Get-Process | Select-Object -ExpandProperty Name                   # Use -ExpandProperty Column to select exactly one column, return results without column title

###########################################################################################################################################################################################

# Tip 46:　Installing PowerShell v3 Help

Update-Help -Force                       # To download the help files

Update-Help -UICulture en-us -Force      # To use English help on non-English systems

explorer $pshome                         # Go to the downloaded help files 

###########################################################################################################################################################################################

# Tip 47: Easier Parameter Attributes in PowerShell v3

param
(
    [Parameter(Mandatory = $true)]            # In PowerShell v2, to declare a function parameter as mandatory
    $p
)

param
(
    [Parameter(Mandatory)]                    # In PowerShell v3, to declare a function parameter as mandatory
    $p
)

###########################################################################################################################################################################################

# Tip 48: Finding Process Owners and Sessions

# Get-Process returns a lot of information about running tasks but it does not return the process owners or the session a process is logged on to. 
# There are built-in console tools like tasklist that do provide this information. By asking these tools to output their information as comma-separated values, 
# PowerShell can pick up the information and make it reusable with your scripts

tasklist

tasklist /V /FO CSV

tasklist /V /FO CSV | ConvertFrom-Csv | Out-GridView

tasklist /V /FO CSV | ConvertFrom-Csv | Select-Object -Property "Image Name", "Session Name", 'Session#', 'User Name'  # to get a list of tasks, their session and the process owner

###########################################################################################################################################################################################

# Tip 49: Normalizing Localized Data

# Many console-based tools like driverquery, whoami, or tasklist provide useful information but the column names are localized and may differ, depending on the language your system uses.

# One way of accessing columns regardless of localization: access columns by position rather than name. 
# To do this, you can use PSObject to find out the localized names, and then pick the names based on a numeric index.

systeminfo /FO CSV | ConvertFrom-Csv | Out-GridView

# To pick just the columns "OS Name", "Product ID", "Original Installation Date" and "System Model" in a language-neutral way, you'd want to pick columns 1,8,9 and 12. 

$data = @(systeminfo /FO CSV | ConvertFrom-Csv)
$columns = $data[0].PSObject.Properties | Where-Object {$_.MemberType -eq "NoteProperty"} | Select-Object -ExpandProperty Name
$data | Select-Object -Property $columns[1,8,9,12]



$data = @(tasklist /V /FO CSV | ConvertFrom-Csv)
$columns = $data[0].PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name
$data | Select-Object -Property $columns[0,1,6,5]
$data | Select-Object -Property $columns[0,3,2,6]



$data | Select-Object -Property * -First 1                     # Look at all available columns anytime to pick the ones you want

###########################################################################################################################################################################################

# Tip 50: Deleting Certificates

Get-ChildItem Cert:\CurrentUser\My                             # Lists all your personal certificates


Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.Subject -like "CN=*test*"} | ForEach-Object {             # Try to delete a certificate

    $store = Get-Item $_.PSParentPath
    $store.Open("ReadWrite")
    $store.Remove($_)
    $store.Close()
}

# Warning: deleting certificates that are needed can damage your computer, so be careful to select only certificates you really want to delete.

###########################################################################################################################################################################################

# Tip 51: Renaming Object Columns

#　To present a solid solution that also allows you to reassign new column names, making the data truly culture-neutral across different platforms.

$data = @(tasklist /V /FO CSV | ConvertFrom-Csv)
$columns = $data[0].PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name
$columnNames = 'Name', 'ID', 'Session', 'SessionID', 'Memory', 'Status', 'Owner', 'CPU', 'Title'
$customProperties = $columns | ForEach-Object {$i=0}{ @{Name = $columnNames[$i]; Expression=[ScriptBlock]::Create(('$_.''{0}''' -f $columns[$i]))}; $i++ }
$data | Select-Object -Property $customProperties
# Output:
#        Name      : powershell_ise.exe
#        ID        : 10404
#        Session   : Console
#        SessionID : 1
#        Memory    : 109,348 K
#        Status    : Running
#        Owner     : FAREAST\v-sihe
#        CPU       : 0:03:32
#        Title     : Administrator: Windows PowerShell ISE



# As a result, you get extensive information about all running processes (that goes beyond what Get-Process can do), 
# and object properties are no longer localized but instead use the names you provided.


$data = @(schtasks /FO CSV | ConvertFrom-Csv)
$columnNames = 'Name', 'NextCall', 'Status'
$columns = $data[0].PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name
$customProperties = $columns | ForEach-Object {$i=0}{ @{Name = $columnNames[$i]; Expression=[ScriptBlock]::Create(('$_.''{0}''' -f $columns[$i]))}; $i++ }
$data | Select-Object -Property $customProperties


###########################################################################################################################################################################################

# Tip 52: Launching Commands That Start With Numbers

# In PowerShell v2, if command names started with a number, you had to use the call operator "&":

& 7z

# In PowerShell v3, this is no longer necessary.

###########################################################################################################################################################################################

# Tip 53: In PowerShell v3, this is no longer necessary

# PowerShell v3 comes with a hugely useful new cmdlet called Invoke-WebRequest. You can use it to interact with websites which also includes downloading files. 



$source = 'http://download.sysinternals.com/files/SysinternalsSuite.zip'
$destination = "$home\silencetest\SysinternalsSuite.zip"

Invoke-WebRequest -Uri $source -OutFile $destination                                                     # download the SysInternals suite of tools to your computer
Unblock-File $destination 

# Since downloaded files are blocked by Windows, PowerShell v3 comes with yet another new cmdlet: Unblock-File removes the block. Now you're ready to unzip the file.

# If your Internet connection requires proxy settings or authentication, take a look at the parameters supported by Invoke-WebRequest. 

Get-Help Invoke-WebRequest -Parameter *

###########################################################################################################################################################################################

# Tip 54: Unzipping Files

# Unfortunately, there is no built-in cmdlet to unzip files. There are plenty of 3rd party tools, many of which are free. 
# If you cannot use these tools, here's how native Windows components can unzip files, too. 

# Note that this requires the built-in Windows ZIP support to be present and not replaced with other ZIP tools.

function Extract-Zip
{
    param([string]$zipfileName, [string]$destination)

    if((Test-Path $zipfileName) -and (Test-Path $destination))           # Both $zipfileName and $destination should exist here
    {
        $shellApplication = New-Object -ComObject Shell.Application
        $zipPackage = $shellApplication.NameSpace($zipfileName)
        $destinationFolder = $shellApplication.NameSpace($destination)
        $destinationFolder.CopyHere($zipPackage.Items())
    }
}

Extract-Zip "$home\silencetest\test.zip" $home\silencetest\test

###########################################################################################################################################################################################

# Tip 55: Manage Windows License Keys

# To automatically manage Windows license keys, use slmgr which is a VBScript that you can call from PowerShell. Just make sure that cscript.exe is your default VBScript host.

wscript.exe //H:cscript               # Set the default script host to "cscript.exe"

# Next, run wscript.exe and disable the option "Show Logo". Now you are set.
# 
# In Windows Server 2012, for example, you can use slmgr /ipk to register a new one.
# 
# If you did register a license key, find out how long it is valid:

slmgr /xpr

# Output:
#         Microsoft (R) Windows Script Host Version 5.8
#         Copyright (C) Microsoft Corporation. All rights reserved.
#         
#         Windows(R) 7, Enterprise edition:
#             Volume activation will expire 2015/04/06 09:57:03
#         



# Check Windows License Status

Get-WmiObject SoftwareLicensingService                                                                                      # access the raw licensing data 
                                                                                                                            
Get-WmiObject SoftwareLicensingProduct | Select-Object -Property Description, LicenseStatus | Out-GridView                  # check the license status of your copy of Windows
                                                                                                                           
Get-WmiObject SoftwareLicensingProduct | Where-Object {$_.LicenseStatus -eq 1} | Select-Object -ExpandProperty Description  # find out which Windows SKU you are actually using
# Output: Windows Operating System - Windows(R) 7, VOLUME_KMSCLIENT channel


# To investigate all the other logic found in slmgr, have a look at the VBScript source code:
Get-Command slmgr
# Output:
#         CommandType     Name               ModuleName                                                                                                                                        
#         -----------     ----               ----------                                                                                                                                        
#         Application     slmgr.vbs                    

notepad (Get-Command slmgr).Path

###########################################################################################################################################################################################

# Tip 56: Calling Native Commands Safely

findstr /s /i "New-Object" *.ps1 C:\Users\v-sihe\Desktop\Tools\PowershellScripts  # To list all PowerShell scripts in c:\windows or a subfolder that contains the word "New-Object"

# run this command both in a cmd.exe shell and in PowerShell, Compare the results. They are different!

# One common workaround is to explicitly run those commands in a cmd.exe shell like this:

cmd.exe /c findstr /s /i 'New-Object' *.ps1 C:\Users\v-sihe\Desktop\Tools\PowershellScripts
# As it turns out, this won't help either. The result is still different. 


# With PowerShell v3, you can finally make sure that your arguments reach the console tool untouched by the parser. Use the new parameter --%:

# Solution for this problem:

cmd.exe --% /c findstr /s /i "New-Object" *.ps1 C:\Users\v-sihe\Desktop\Tools\PowershellScripts

###########################################################################################################################################################################################

# Tip 57: $PSItem in PowerShell v3

Get-ChildItem $env:windir | Where-Object {$_.Length -gt 1MB}       # In PowerShell, the variable "$_" has special importance. It works like a placeholder 

Get-ChildItem $env:windir | Where-Object {$PSItem.Length -gt 1MB}  # In PowerShell v3, there is an alias for the cryptic "$_": $PSItem. So now code can become more descriptive

Get-ChildItem $env:windir | Where-Object Length -GT 1MB            # In PowerShell v3, "$_" isn't necessary in many scenarios anymore at all

# Output: 
#        Directory: C:\Windows
#    
#    
#    Mode                LastWriteTime     Length Name                                                                                                                                                                    
#    ----                -------------     ------ ----                                                                                                                                                                    
#    -a---        2011/06/03  01:21 AM    2871808 explorer.exe                                                                                                                                                            
#    -a---        2014/03/24  09:09 AM  547770451 MEMORY.DMP                                                                                                                                                              
#    -a---        2014/10/08  10:52 AM    1551125 WindowsUpdate.log 

###########################################################################################################################################################################################

# Tip 58: Discovering Useful Console Commands

# There are plenty of useful console commands such as ipconfig, whoami, and systeminfo. Most of these commands hide inside the Windows folder.


function Get-ConsoleCommand          # lists all available commands with a short description of what they do in a window
{
    $ext = $env:PATHEXT -split ";" -replace "\.","*."
    $desc = @{Name = "Description"; Expression = {$_.FileVersionInfo.FileDescription}}

    Get-Command -Name $ext | Select-Object Name, Extension, $desc 
}

Get-ConsoleCommand

# Output: (Part of output)
#         ...                        ...              ...

#         xcopy.exe                  .exe             Extended Copy Utility                                                 
#         xperf.exe                  .exe             Performance Analyzer Command Line                                     
#         xperfview.exe              .exe             Performance Analyzer                                                  
#         xpsrchvw.exe               .exe             XPS Viewer                                                            
#         xwizard.exe                .exe             Extensible Wizards Host Process                                       
#         ZuneWlanCfgSvc.exe         .exe             Zune Wireless Configuration Service

###########################################################################################################################################################################################

# Tip 59: Using Closures 


# Script blocks are dynamic by default, so variables in a script block will be evaluated each time the script block runs.

# By turning a script block in a "closure", it keeps the state of the variables.

$var = "A"
$scriptBlock = {$var}
& $scriptBlock                              # Output: A
                                            
$var = "B"                                  
& $scriptBlock                              # Output: B

$var = "D"
$closure = $scriptBlock.GetNewClosure()
$var = "E"
& $closure                                  # Output: D

$var = "F"
& $closure                                  # Output: D

# Note that the closure script block will always return "D" because that was the value of $var when the closure was created. 

###########################################################################################################################################################################################

# Tip 60: Resetting Console Colors in Powershell

# If a console application or script has changed the console colors and you want to reset them to the default colors defined in your console properties, try this:

$Host.UI.RawUI.ForegroundColor = "Red"
[Console]::ResetColor()                   

# Note that this will work in a real console only. It does not work in the ISE Console pane.

###########################################################################################################################################################################################

# Tip 61: Using Local Variables Remotely

# If you want to send a script block to a remote computer, it is important to understand that the script block is evaluated on the remote computer. So all local variables are null

$path = "C:\Test"
$scriptBlock = {"Variable path is $path"}
& $scriptBlock                                                       # Output: Variable path is C:\Test
Invoke-Command -ScriptBlock $scriptBlock -ComputerName IIS-CTI5052   # Output: Variable path is 


$path = "C:\Test"
$scriptBlock = {"Variable path is $using:path"}                      # In Powershell V3, prefix a local variable with "using:" so that it will get transferred to the remote machine
Invoke-Command -ScriptBlock $scriptBlock -ComputerName IIS-CTI5052   # Output: Variable path is C:\Test

# Note that the prefix "using" is allowed only in script blocks that are executed by Invoke-Command and sent to a remote machine or with InlineScript inside a workflow.

###########################################################################################################################################################################################

# Tip 62: Executing Code Locally and Remotely Using Local Variables

function Get-Log($logName = "System", $computerName = $env:COMPUTERNAME)    # this function is designed to work both locally and remotely
{
    $code = {
        
        param($logName)
        
        Get-EventLog -LogName $logName -EntryType Error -Newest 5
    }

    $null = $PSBoundParameters.Remove("logName")

    Invoke-Command -ScriptBlock $code.GetNewClosure() @PSBoundParameters -ArgumentList $logName
}


Get-Log
Get-Log -computerName IIS-CTI5052


# Note: $code = { Get-EventLog -LogName $Using:LogName -EntryType Error -Newest 5 }, 
# the code would run fine remotely, but now it would fail locally. The prefix "using" is allowed only in remote calls or in a workflow inlinescript statement.


#[Reference site] http://powershell.com/cs/blogs/tips/archive/2012/10/26/executing-code-locally-and-remotely-using-local-variables.aspx

###########################################################################################################################################################################################

# Tip 63: Line Breaks After "." and "::"

# In PowerShell v3 language syntax, it is finally allowed to have line breaks after "." and "::". These symbols are used to access dynamic and static object properties.

"Hello".ToUpper().Replace("HE","RO")       # Output: ROLLO



"Hello".
toUpper().
Replace("HE","RO")                         # Output: ROLLO



# take some text...
"Hello".
# ...convert it to uppercase...
toUpper().
# ...and replace some text:
Replace('HE', 'RO')                        # Output: ROLLO

###########################################################################################################################################################################################

# Tip 64: Validation Attributes On Variables

[ValidateRange(1,10)][int]$x = 5

$x = 12                                    # Exception here: The variable cannot be validated because the value 12 is not a valid value for the x variable.

###########################################################################################################################################################################################

# Tip 65: New Operators in PowerShell v3

<#
  There are four new operators in PowerShell v3:
    
    -shl: shifts bits to the left
    
    -shr: shifts bits to right and preserves sign for signed values
    
    -in: works like -contains, operand order is reversed
    
    -notin: works like -notcontains, operand order is reversed    
#>

###########################################################################################################################################################################################

# Tip 66: Creating Custom Objects in Powershell

# In Powershell V2:
$newObject = "dummy" | Select-Object -Property Name, ID, Address
$newObject.Name = $env:USERNAME
$newObject.ID = 12
$newObject.Address = "Shanghai"

$newObject
# Output:
#         Name        ID        Address                                                               
#         ----        --        -------                                                               
#         v-sihe      12        Shanghai 

# In PowerShell v2, there has been an alternate way using hash tables and New-Object. 
# This never worked well, though, because a default hash table is not ordered, so your new objects used random positions for the properties.




# In Powershell V3:
$newObject = [PSCustomObject][Ordered]@{ Name = $env:USERNAME; ID = 12; Address = "Shanghai"}

$newObject
# Output:
#         Name        ID        Address                                                               
#         ----        --        -------                                                               
#         v-sihe      12        Shanghai 

# In PowerShell v3, you can create ordered hash tables and easily cast them to objects. This way, you can create new custom objects and fill them with data in just one line of code

###########################################################################################################################################################################################

# Tip 67: Find Open Files

openfiles /Query /S IIS-CTI5052 /FO CSV /V | ConvertFrom-Csv | Out-GridView      # To find open files on a remote system, use openfiles.exe and convert the results to rich objects. 

###########################################################################################################################################################################################

# Tip 68: Change Order of CSV Columns

$path = "C:\somepathtocsv.csv"
(Import-Csv -Path $path) | Select-Object -Property column1, column3, column2 | Export-Csv -Path $Path

# Note the parenthesis: they allow you to write back to the same file because PowerShell will first read in the CSV completely before it changes data and writes back to the same file.

# Try this with some non-productive CSV material first. You may have to add the -Delimiter parameter to the CSV cmdlets if your CSV files are using a delimiter other than a comma.

###########################################################################################################################################################################################

# Tip 69: Get-WmiObject Becomes Obsolete

# In PowerShell 3.0, while you still can use the powerful Get-WmiObject cmdlet, it is slowly becoming replaced by the family of CIM cmdlets.

# If you use Get-WmiObject to query for data, you can easily switch to Get-CimInstance. Both work very similar. 
# The results from Get-CimInstance, however, do not contain any methods anymore.

# If you must ensure backwards compatibility, on the other hand, you may want to avoid CIM cmdlets. 
# They require PowerShell 3.0 and won't run on Windows XP and Vista/Server 2003

Get-Help Get-CimInstance

###########################################################################################################################################################################################

# Tip 70: Preserving Special Characters in Excel-generated CSV files

# When you save Excel spreadsheets to a CSV file, special characters get lost. That's because Excel is saving the CSV file using very simple ANSI encoding.

$path = "C:\somepathtocsv.csv"
(Get-Content $path) | Set-Content $path -Encoding UTF8   # Re-encodes the CSV file and uses UTF8 encoding, making special characters readable for Import-CSV

###########################################################################################################################################################################################

# Tip 71: Finding Object Properties in Powershell


# Sometimes, you know the information you are after is present in some object property, 
# but there are so many properties that it is a hassle to search for the one that holds the information.

# In cases like this, you can convert the object to text lines, and then use Select-String to find the line (and the property) you are after:

#  AdapterType      : Ethernet 802.3
#  Name             : Apple Mobile Device Ethernet
#  Name             : Apple USB Ethernet Adapter
#  AdapterType      : Ethernet 802.3
#  AdapterType      : Ethernet 802.3
#  Name             : VirtualBox Host-Only Ethernet Adapter
#  Name             : Apple Mobile Device Ethernet
#  
#  As it turns out, the information is present in two different properties: AdapterType as well as Name.

###########################################################################################################################################################################################

# Tip 72: Detecting STA-Mode

[Runspace]::DefaultRunspace.ApartmentState -eq "STA"

# This information may be important because only in STA-mode can PowerShell run WPF functionality (like windows based on WPF or system dialogs like Open or Save As). 
# In PowerShell 3.0, both console and ISE run in STA mode by default. In PowerShell 2.0, the console used MTA mode. It needs to be started with the parameter -STA to enable STA-mode.

###########################################################################################################################################################################################

# Tip 73: Providing "Static" IntelliSense for Your Functions

# To get rich IntelliSense in PowerShell ISE 3.0, you should start adding the OutputType attribute to your functions. 
# If you do, then ISE is able to provide IntelliSense inside your code without the need to actually have real values in your variables.

function Test-IntelliSense
{
    [OutputType("System.DateTime")]
    param()

    return Get-Date
}

$a = Test-IntelliSense
# $a.

# The moment you type a dot after the variable, ISE provides IntelliSense for the System.DateTime data type. Isn't that nice? 
# No need to run the script all the time to fill variables for real-time IntelliSense.

###########################################################################################################################################################################################

# Tip 74: Finding Built-In Cmdlets

# In times where cmdlets can originate from all kinds of modules, 
# it sometimes becomes important to find out which cmdlets are truly built into PowerShell and which represent external dependencies.

# One way of getting a list of built-in cmdlets is to temporarily open another runspace and enumerate its internal cmdlet list

$ps = [Powershell]::Create()
$ps.Runspace.RunspaceConfiguration.Cmdlets | Select-Object Name
$ps.Dispose()

###########################################################################################################################################################################################

# Tip 75: Rich IntelliSense for Function Arguments

# To take advantage of the new PowerShell 3.0 argument completion, make sure you're adding ValidateSet attribute to your function parameters (where appropriate). 

function Select-Color
{
    param
    (
        [ValidateSet("Red","Green","Blue")]
        $color
    )

    "You chose $color"
}

# Once you run this code and then in the ISE, enter this, you'll see the new argument completion in action
Select-Color -color # <-ISE opens IntelliSense menu with color values



# Using Enumeration Types for Parameter IntelliSense

function Select-Color2
{
    param
    (
        [System.ConsoleColor]
        $color
    )

    "You chose $color"
}

# Once you run this code and then in the ISE, enter this, you'll see the new argument completion in action:
Select-Color2 -color # <-ISE opens IntelliSense menu with color values

###########################################################################################################################################################################################

# Tip 76: Finding Enumeration Data Types

[AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetExportedTypes() } | Where-Object { $_.isEnum } | Sort-Object FullName | ForEach-Object {

    $values = [System.Enum]::GetNames($_) -join ","
    
    $rv = $_ | Select-Object -Property FullName, Values
    $rv.Values = $values
    
    $rv
}

# Output:
#        System.Windows.WindowStartupLocation                   Manual,CenterScreen,CenterOwner                                                                           
#        System.Windows.WindowState                             Normal,Minimized,Maximized                                                                                
#        System.Windows.WindowStyle                             None,SingleBorderWindow,ThreeDBorderWindow,ToolWindow                                                     
#        System.Windows.WrapDirection                           None,Left,Right,Both                                                                                      
#        System.Xaml.Schema.AllowedMemberLocations              None,Attribute,MemberElement,Any


# Note: When you run this, you may get some error messages that you can safely ignore. You do get a list of enumeration types plus a list of the values they represent.

###########################################################################################################################################################################################

# Tip 77: WhatIf-Support Without Propagation

function Test-WhatIf
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $really = $PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Do something dangerous!")

    if($really)
    {
        "OK, I am doing it."
    }
    else
    {
        "Just kidding..."
    }
}
# Try and run this function as-is, then again with -WhatIf or -Confirm. It works!

Test-WhatIf          
# Output: OK, I am doing it.

Test-WhatIf -WhatIf
# Output:
#         What if: Performing operation "Do something dangerous!" on Target "IIS-V-SIHE-01".
#         Just kidding...

Test-WhatIf -Confirm
# Output:
#        Yes:  OK, I am doing it.
#        No:   Just kidding...


# However, the parameters get propagated to any other cmdlet that is called directly or indirectly from your function. 
# So if you call some other function or module command, then there may be additional simulation messages coming from there, too.

# If you want to disable this, make sure you reset $WhatIfPreference to the parent scope:

$WhatIfPreference = (Get-Variable WhatIfPreference -Scope 1).Value              # Now, any command outside your If-clauses will always execute.

###########################################################################################################################################################################################

# Tip 78: Creating New Objects the JSON way

# There are numerous ways how you can create new objects that you may use to return results from your functions. 
# One way is using JSON, a very simple description language. It is fully supported in PowerShell 3.0.

$content = '{"Name":"weltner","FirstName":"tobias","id":123}'
ConvertFrom-Json $content
# Output:
#         Name            FirstName            id
#         ----            ---------            --
#         weltner         tobias              123

###########################################################################################################################################################################################

# Tip 79: Finding Newest or Oldest Files

# Return the smallest and largest file size in your Windows folder in Powershell V2
Get-ChildItem $env:windir | Measure-Object -Property Length -Minimum -Maximum | Select-Object -Property Minimum,Maximum  
# Output:
#         Minimum                              Maximum
#         -------                              -------
#               0                            547770451



# In PowerShell 3.0, you could also measure properties like LastWriteTime, telling you the oldest and newest dates:
Get-ChildItem $env:windir | Measure-Object -Property LastWriteTime -Minimum -Maximum | Select-Object -Property Minimum,Maximum 
# Output:
#         Minimum                              Maximum                                                                                                   
#         -------                              -------                                                                                                   
#         2006/02/09 17:50:00                  2014/10/09 16:52:19



# you could get the minimum and maximum start times of all the running processes. Make sure you use Where-Object to exclude any process that has no StartTime value:
Get-Process | Where-Object StartTime | Measure-Object -Property StartTime -Minimum -Maximum | Select-Object -Property Minimum,Maximum 
# Output:
#         Minimum                              Maximum                                                                                                   
#         -------                              -------                                                                                                   
#         2014/10/08 08:43:51                  2014/10/09 17:36:45

###########################################################################################################################################################################################

# Tip 80: Protecting Functions

function Test-Function { "Hello World!" }
Set-Item -Path function:Test-Function -Options ReadOnly

function Test-Function { "try to change the readonly function" }      # Exception here: Cannot write to function Test-Function because it is read-only or constant

###########################################################################################################################################################################################

# Tip 81: Creating Objects in PowerShell 3.0 (Fast and Easy) 

$content = @{
   
    Name = "Silence"
    FirstName = "He"
    id = "No.1"
}

[PSCustomObject]$content

# Output:
#         Name                         FirstName                          id                                                                    
#         ----                         ---------                          --                                                                    
#         Silence                      He                                 No.1  

# The primary advantage of this approach is that you control the order of object properties. With the approach used in PowerShell 2.0, this was not guaranteed:

New-Object PSObject -Property $content

# Note that in PowerShell 2.0, conversion to PSCustomObject fails, and you get back a hash table.

###########################################################################################################################################################################################

# Tip 82: Logging Input Commands

# If you'd like to maintain a log file with all the commands you entered interactively - in the PowerShell console as well as in the ISE editor - here is an easy way:
# Simply redefine the built-in prompt function. It is responsible for writing the prompt text. 
# Since it gets called automatically AFTER each command you enter, you can add code to log the last command. The last command is available from Get-History.

# Here is an example of such a function. It also shortens the prompt text and instead displays the current location in the window title bar.
function prompt
{
    "PS> "
    $host.UI.RawUI.WindowTitle = Get-Location

    if($Global:cmdLogFile)
    {
        Get-History -Count 1 | Select-Object -ExpandProperty CommandLine | Out-File $Global:cmdLogFile -Append
    }
}

$Global:cmdLogFile = "$env:temp\logfile.txt"                         # To enable logging, set $global:CmdLogFile to a path

Remove-Variable cmdLogFile -Scope global                             # To disable logging, remove the variable

###########################################################################################################################################################################################

# Tip 83:　Controlling Object Property Display

# When you create functions that return custom objects, there is no way for you to declare which functions should display by default. 
# PowerShell always displays all properties, and when there are more than 4, you get a list display, else a table
function Get-SomeResult
{
    $resultset = New-Object PSObject | Select-Object Name, FirstName, ID, Language, Skill

    $resultset.Name = "Silence"
    $resultset.FirstName = "He"
    $resultset.ID = 123
    $resultset.Language = "Powershell"
    $resultset.Skill = "5"

    $resultset
}

Get-SomeResult
# Output:
#          Name      : Silence
#          FirstName : He
#          ID        : 123
#          Language  : Powershell
#          Skill     : 5

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# In PowerShell 3.0, you can now define the parameters of a type that should be displayed by default. 
# To make that work with your function, first assign a new custom type to your return objects, 
# and then call Update-TypeData to tell PowerShell which properties it should display for this type of data
function Get-SomeResult
{
    $resultset = New-Object PSObject | Select-Object Name, FirstName, ID, Language, Skill
    $resultset.Name = "Silence"
    $resultset.FirstName = "He"
    $resultset.ID = 123
    $resultset.Language = "Powershell"
    $resultset.Skill = "5"

    if($Host.Version.Major -ge 3)
    {
        $objectType = "myObject"
        $resultset.PSTypeNames.Add($objectType)   # Note: it defines a permanent new type here

        if((Get-TypeData $objectType) -eq $null)
        {
            $p = "Name", "ID", "Language"
            Update-TypeData -TypeName $objectType -DefaultKeyPropertySet $p
        }
    }

    $resultset
}

# It works beautifully in PowerShell 3.0 and falls back to default behavior in PowerShell 2.0:
Get-SomeResult
# Output:
#         Name                   ID                 Language                                                              
#         ----                   --                 --------                                                              
#         Silence               123                 Powershell 

Get-SomeResult | Select-Object *
# Output:
#          Name      : Silence
#          FirstName : He
#          ID        : 123
#          Language  : Powershell
#          Skill     : 5


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Above approach works well, but it defines a permanent new type. 
# Here's an example that defines a custom standard object properties "on the fly" without touching the PowerShell type database

function Get-SomeResult
{
    $resultset = New-Object PSObject | Select-Object Name, FirstName, ID, Language, Skill
    $resultset.Name = "Silence"
    $resultset.FirstName = "He"
    $resultset.ID = 123
    $resultset.Language = "Powershell"
    $resultset.Skill = "5"

    [string[]]$properties = "Name", "ID", "Language"
    [System.Management.Automation.PSMemberInfo[]]$psStandardMembers = New-Object System.Management.Automation.PSPropertySet DefaultDisplayPropertySet,$properties

    $resultset | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $psStandardMembers

    $resultset
}

Get-SomeResult
# Output:
#         Name              ID             Language                                                              
#         ----              --             --------                                                              
#         Silence          123             Powershell 

Get-SomeResult | Select-Object *
# Output:
#        Name      : Silence
#        FirstName : He
#        ID        : 123
#        Language  : Powershell
#        Skill     : 5

###########################################################################################################################################################################################

# Tip 84: Removing Empty Object Properties

# Objects hold a lot of information and often, properties can also have null values. To reduce an object to only those properties that actually have a value, 
# you can convert the object into a hash table and remove all empty properties, then turn the hash table back into an object. 
# This also gives you the opportunity of sorting object property names. 

# This example will read BIOS information from the WMI and then return an object stripped of all empty properties. The code requires PowerShell 3.0:

$bios = Get-WmiObject -Class Win32_BIOS

$hashtable = $bios | Get-Member -MemberType *Property | Select-Object -ExpandProperty Name | Sort-Object | ForEach-Object -Begin {
    
    [System.Collections.Specialized.OrderedDictionary]$rv = @{}
    # Note the use of the System.Collections.Specialized.OrderedDictionary type: it creates a special ordered hash table. 
    # Regular hash tables do not keep a specific order in their keys. 

} -Process {

    if($bios.$_ -eq $null)
    {
        Write-Warning "Removing empty property $_"
    }
    else
    {
        $rv.$_ = $bios.$_
    }

} -End {$rv}

$biosNew = New-Object PSObject
$biosNew | Add-Member ($hashtable) -ErrorAction SilentlyContinue

$biosNew


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Removing Empty Object Properties (All Versions)
$bios = Get-WmiObject -Class Win32_BIOS
$biosNew = $bios | Get-Member -MemberType *Property | Select-Object -ExpandProperty Name | Sort-Object | ForEach-Object -Begin { $obj = New-Object PSObject } {

    if($bios.$_ -eq $null)
    {
        Write-Warning "Removing empty property $_"
    }
    else
    {
        $obj | Add-Member -MemberType NoteProperty -Name $_ -Value $bios.$_
    }
} { $obj }

$biosNew

###########################################################################################################################################################################################

# Tip 85: "Count" Available in PowerShell 3.0

(Get-ChildItem $env:windir\*.txt).Count                                            # Output: 0                                                                                
(Get-ChildItem $env:windir\explorer.exe).Count                                     # Output: 1
(Get-ChildItem $env:windir\nothing.there -ErrorAction SilentlyContinue).Count      # Output: 0


# In PowerShell 2.0, this would have returned 2 for the first call, an exception for the second (since there is only one explorer.exe, 
# PowerShell does not wrap it into an array) and nothing for the last call. That's why in PowerShell 2.0, you would have to manually wrap results into an array 
# (and still need to do if you want your code to be compatible to PowerShell 2.0). This code runs in all versions of PowerShell:

@(Get-ChildItem $env:windir\*.txt).Count                                            # Output: 0                                                                                 
@(Get-ChildItem $env:windir\explorer.exe).Count                                     # Output: 1
@(Get-ChildItem $env:windir\nothing.there -ErrorAction SilentlyContinue).Count      # Output: 0

# Another reason why the "old" syntax may be better than the new: once you enable Strict-Mode, the behavior falls back to PowerShell 2.0 standards and breaks the new syntax.

###########################################################################################################################################################################################

# Tip 86: NULL values have a Count property

# in PowerShell 3.0, every object has a Count property now. This even includes null values
$null.Count            # Output: 0, This is so that when a command returns "nothing", you still can find out how many items you received by querying Count.

# In PowerShell 3.0, you can always use Count to find out whether you received some result, 
# and if so, how many results you got. Except if you use Strict-Mode. Then, the magic Count is removed.

# This redesign might also be in the context of another change: folder objects no longer have a truly empty Length property. 
# In PowerShell 2.0, you were able to list folders this way:
Get-ChildItem $env:windir | Where-Object { $_.Length -eq $null }


# This no longer works in PowerShell 3.0. You would have to resort to another suitable property (or use the new PowerShell 3.0 parameters in Get-ChildItem):
Get-ChildItem $env:windir | Where-Object { $_.PSIsContainer }
Get-ChildItem $env:windir -Directory

###########################################################################################################################################################################################

# Tip 87: Using CIM Cmdlets Against PowerShell 2.0

# Get-CimInstance, for example, works very similar to Get-WmiObject. It focuses on object properties, though, because the CIM cmdlets can base on different remoting techniques, 
# some of which are stateless (and very robust). By default, all CIM cmdlets use WinRM, the new remoting technique found in PowerShell 3.0 as well.

# So by default, you cannot remotely access a system that has no WinRM capabilities. 
# With a little trick, however, you can convince the CIM cmdlets to fall back and use the old DCOM protocol instead. Here is an example:

Get-CimInstance -ClassName Win32_BIOS
Get-CimInstance -ClassName Win32_BIOS -ComputerName IIS-CTI5052

# This may fail if that system won't support WinRM, for example, because it might be some old server 2003 with PowerShell 2.0 on it. 
# In this case, you'd get an ugly red exception message, complaining that some DTMF resource URI wasn't found.

# So next are the lines you can use to force Get-CimInstance to use the old DCOM technology:

$option = New-CimSessionOption -Protocol Dcom
$session = New-CimSession -ComputerName IIS-CTI5052 -SessionOption $option

Get-CimInstance -ClassName Win32_BIOS -CimSession $session

# This time, the server could be contacted and returns the data.

###########################################################################################################################################################################################

# Tip 88: Combining Objects

$bios = Get-WmiObject -Class Win32_BIOS
$os = Get-WmiObject -Class Win32_OperatingSystem

$hashtable = $bios | Get-Member -MemberType *Property | Select-Object -ExpandProperty Name | Sort-Object | ForEach-Object { $rv = @{} } {

    Write-Warning $_

    $rv.$_ = $bios.$_

} { $rv }

$os | Add-Member ($hashtable) -ErrorAction SilentlyContinue
$os
# Output:
#         SystemDirectory : C:\Windows\system32
#         Organization    : Microsoft IT
#         BuildNumber     : 7601
#         RegisteredUser  : Silence
#         SerialNumber    : 00392-918-5000002-85646
#         Version         : 6.1.7601

# When you output $os, at first sight it did not seem to change because PowerShell continues to display the default properties. However, the BIOS information is now added to it:

$os | Select-Object *BIOS*
# Output:
#         BiosCharacteristics : {7, 9, 11, 12...}
#         BIOSVersion         : {HPQOEM - 20090825}
#         PrimaryBIOS         : True
#         SMBIOSBIOSVersion   : 786G6 v01.03
#         SMBIOSMajorVersion  : 2
#         SMBIOSMinorVersion  : 6
#         SMBIOSPresent       : True

###########################################################################################################################################################################################

# Tip 89: Show-Command Creates PowerShell-Code for You

Show-Command Get-Process       # In PowerShell 3.0, there is a cool new cmdlet called Show-Command

# It works both in the console and the ISE editor, and when you specify a cmdlet, a dialog window opens and shows a form that helps you discover and fill in the cmdlet parameters. 
# Once done, click Copy, then Cancel, and the complete command line code is available from the clipboard. To insert it into the console, for example, simply right-click.

# The dialog window produced by Show-Command also shows the different parameter sets (groups of parameters) a cmdlet supports, 
# so you never again run into issues where you accidentally mix parameters from different parameter sets.

###########################################################################################################################################################################################

# Tip 90: Examine Parameter Binding

Get-ChildItem -Path $env:windir -Filter *.jpg -Recurse -Name -ErrorAction SilentlyContinue

ls -r $env:windir -n *.jpg -ea 0 

# The first line is "politically correct" and uses named parameters and cmdlet names. The second trusts in positional parameters and aliases, 
# and PowerShell internally "binds" the arguments to the correct parameters. If you'd like to see how this magic is done 
# (and maybe want to "decipher" a cryptic line of PowerShell code yourself), use Trace-Command like this:

Trace-Command -PSHost -Name ParameterBinding { ls -r $env:windir -n *.jpg -ea 0  }

###########################################################################################################################################################################################

# Tip 91: Executing Elevated PowerShell Code

$code = "New-Item HKLM:\SoftWare\somekey2"

Start-Process -FilePath Powershell -Verb runas -WindowStyle Minimized -ArgumentList(' -noprofile -noexit -command ' + $code)

Test-Path hklm:\software\somekey

###########################################################################################################################################################################################

# Tip 92: Controlling Process Priority and Processor Affinity

# When you get yourself a process using Get-Process, what you get back is an object that has useful methods and writeable properties.

# This line will assign a high priority to your current PowerShell host and would run on all 4 CPU cores (provided your machine has 4 cores):
$process = Get-Process -id $pid
$process.BasePriority = "High"
$process.ProcessorAffinity = 15



# ProcessorAffinity really is a bitmask where each bit represents one CPU core. Since Windows won't limit default processes to a specific CPU core, 
# you can use this property also to find out how many CPU cores your machine has:

[System.Convert]::ToString([int]$process.ProcessorAffinity, 2)      # Output: 111

# It is admittedly a creative approach, but to find out the number of CPUs, you'd have to count the bits, and once you convert the bits to a string, that's the length of the string:

$cpus = ([System.Convert]::ToString([int]$process.ProcessorAffinity, 2)).Length
"Your machine has $cpus (cores)."

###########################################################################################################################################################################################

# Tip 93: Sending Results to Excel

function Out-ExcelReport
{
    param($path = "$env:temp\$(Get-Random).csv")

    $input | Export-Csv -Path $path -Encoding UTF8 -NoTypeInformation -UseCulture

    Invoke-Item -Path $path
}

Get-Process | Out-ExcelReport

###########################################################################################################################################################################################

# Tip 94: Ripping All Links from a Website

# PowerShell 3.0 comes with a great new cmdlet: Invoke-WebRequest! You can use it for a zillion things, 
#　but it can also simply retrieve the content of a website. It will even do basic parsing, so opening a window with all links on that website is a piece of cake:

$webSite = Invoke-WebRequest -UseBasicParsing -Uri http://www.baidu.com
$webSite.Links | Out-GridView

###########################################################################################################################################################################################

# Tip 95: Use -f with N0

# Often, it is necessary to output numbers, but you may want to control the number of digits and would like to control the formatting. 
# The -f operator can do this and has a trillion options but there's just one you need to remember: N0 (the "0" is the number zero).

"{0:N0}" -f (8gb / 12kb)       # Output: 699,051
"{0:N1}" -f (8gb / 12kb)       # Output: 699,050.7
"{0:N2}" -f (8gb / 12kb)       # Output: 699,050.67

# It rounds the number so there are no digits after the decimal, and it adds a separator every three numbers if the number is more than 999.
# Replace "N0" with "N1" or any other number to control the digits after the decimal

###########################################################################################################################################################################################

# Tip 96: Use Comparison Operators for Logfile Parsing

# Comparison operators usually return either $true or $false, but when applied to an array, return the array elements that match the comparison. 
# You can use this to easily parse text-based logfile information.

# returns all updates installed on a machine
(Get-Content -Path $env:windir\windowsupdate.log -ReadCount 0 -Encoding UTF8) -like '*successfully installed*' | ForEach-Object { ($_ -split ': ')[-1] }

###########################################################################################################################################################################################

# Tip 97: Cutting Off Text at the End

$text = 'C:\folder\file.txt'

$text.Substring(3)                      # Output: folder\file.txt
$text.Remove($text.Length - 4)          # Outout: C:\folder\file
$text -replace ".{4}$"                  # Output: C:\folder\file         to search for "anything" (".") four times ({4}) at the end of a text ($). 
$text -replace "\..*?$"                 # Output: C:\folder\file         cut off a file extension, no matter how long it is

# It looks for a real dot (\.), then anything (.) as many times as necessary (*?) to reach the text end ($).

###########################################################################################################################################################################################

# Tip 98: Splitting Texts without Losing Anything

$profile                               # Output: C:\Users\v-sihe\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1

# Typically when you split a text using the -split operator or the Split() method, the split character is removed from the text:

$profile -split "\\"
# Output:
#         C:
#         Users
#         v-sihe
#         Documents
#         WindowsPowerShell
#         Microsoft.PowerShellISE_profile.ps1

# If you want to keep it, make it a "look ahead" by adding "?<=". This way, PowerShell looks for the backslash, then looks "ahead" and cuts right after it:

$profile -split "(?<=\\)"
# Output:
#         C:\
#         Users\
#         v-sihe\
#         Documents\
#         WindowsPowerShell\
#         Microsoft.PowerShellISE_profile.ps1

# You can also make it a "look behind", so it searches for the backslash, then turns around and looks "back", 
# cutting the text right before the backslash, again without removing it. Add a "?=" this time:

$profile -split "(?=\\)"
# Output:
#         C:
#         \Users
#         \v-sihe
#         \Documents
#         \WindowsPowerShell
#         \Microsoft.PowerShellISE_profile.ps1

###########################################################################################################################################################################################

# Tip 99: Splitting Hexadecimal Pairs

# If you'd have to process a long list of encoded information, let's say a list of hexadecimal values, how would you split the list into pairs of two? Here is a way:

'this gets splitted in pairs of two' -split "(?<=\G.{2})(?=.)"
# Output:
#         th
#         is
#          g
#         et
#         s 
#         sp
#         li
#         tt
#         ed
#         (...)



# This would be more specific and split only hex values:

'00AA1CFFAB1034' -split '(?<=\G[0-9a-f]{2})(?=.)'
# Output:
#        00
#        AA
#        1C
#        FF
#        AB
#        10
#        34



# Now it's easy to reformat MAC addresses as well:

'00AA1CFFAB1034' -split '(?<=\G[0-9a-f]{2})(?=.)' -join ':'       # Output: 00:AA:1C:FF:AB:10:34

###########################################################################################################################################################################################

# Tip 100: Creating "Mini-Modules"

function Get-BIOS
{
    param($computerName, $credential)

    Get-WmiObject -Class Win32_BIOS @PSBoundParameters
}

# This function called Get-BIOS will get the computer BIOS information and supports -ComputerName and -Credential for remote access, too. Make sure you run the function and test it.

# turn the function into a module
$name = "Get-BIOS"
New-Item -Path $home\Documents\WindowsPowerShell\Modules\$name\$name.psm1 -ItemType File -Force -Value "function $name { $((Get-Item function:\$name).Definition) }"

# Now why is this conversion important? Because PowerShell 3.0 auto-detects functions in modules!
# So if you use PowerShell 3.0 and have created the "mini module", open up a new PowerShell and type:

Get-BIOS

# Bam! Your function is now available automatically, and you can easily add new functionality.

###########################################################################################################################################################################################