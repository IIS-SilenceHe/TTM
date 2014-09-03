# Reference site: http://powershell.com/cs/blogs/tips/
###########################################################################################################################################################################################

# Tip 1: Asynchronous Downloads with BITS

# grabs PowerShellPlus 32Bit asynchronously and saves it on your hard drive
Start-BitsTransfer "http://downloads.idera.com/products/IderaPowerShellPlusInstallationKit.zip" "$home\psplus32bit.zip" -async

Get-BitsTransfer | Format-List *                                                   # list all pending downloads and also let you know whether the download has completed or its progress

# Note: all you need for large downloads is mark the operation as asynchronously. This way, it will run as a background job and be scheduled with Windows. 
 # No longer will the PowerShell console have to wait for the download to complete, and the download will continue even when you close PowerShell or reboot your machine

###########################################################################################################################################################################################

# Tip 2: Exporting Certificate

dir cert:\ -Recurse

@(dir cert:\ -Recurse | Where-Object {$_.Subject -like "*$env:username*"})[0]      # to grab the first certificate that has your username in its subject

# Store the certificate in a variable and then view all of its properties to view certificate details
$cert = @(dir cert:\ -Recurse | Where-Object {$_.Subject -like "*$env:username*"})[0]
$cert | Format-List *

# call the export() method, which gets you a byte array, to export the certificate. Next, use .NET to write the byte array to disk
$bytes = $cert.Export("Cert")
[System.IO.File]::WriteAllBytes("$home\mycert.cer", $bytes)
dir $home\*.cer



# Exporting Certificate With Private Key
dir Cert:\CurrentUser\My | Where-Object {$_.HasPrivateKey}                   # to find all certificates which have private key in your personal store

dir Cert:\LocalMachine\My | Where-Object {$_.HasPrivateKey}                  # to see all machine certificates which have private key (provided you are Admin)

# export all of your personal certificates, including private key to pfx-files in your user profile. Each file uses the certificate thumbprint as its file name
dir Cert:\CurrentUser\my | Where-Object {$_.HasPrivateKey} | ForEach-Object {[System.IO.File]::WriteAllBytes("$home\test\$($_.Thumbprint).pfx", ($_.Export("PFX", "secret")))}
# Note: The certificate can be exported successfully but exception here, the line has set this password to 'secret'

###########################################################################################################################################################################################

# Tip 3: Find Next Available Drive Letter

function Get-NextFreeDrive
{
    68..90 | ForEach-Object {"$([char]$_):"} | Where-Object { "h:", "k:", "z:" -notcontains $_} | Where-Object {(New-Object System.IO.DriveInfo $_).DriveType -eq "noRootdirectory"}
}

# Note: It starts by enumerating ASCII codes for letters D: through Z:. The pipeline will then convert those ASCII codes into drive letters. 
 # Next, check out how the pipeline uses an exclusion list with drives you do not want to use for mapping (optional). 
 # In this example, drive letters h:, k:, and z: are never used. Finally, the pipeline uses a .NET DriveInfo object to check whether the drive is in use or not

(Get-NextFreeDrive)[0]

net use (Get-NextFreeDrive)[0] "\\127.0.0.1\C$"                # to get the first available drive letter for drive mapping

###########################################################################################################################################################################################

# Tip 4: Finding Unused Drives

Get-WmiObject Win32_LogicalDisk | Format-List *
Get-WmiObject Win32_LogicalDisk | Format-List DeviceID, Description, FileSystem, Providername

$drives = Get-WmiObject Win32_LogicalDisk | ForEach-Object{$_.DeviceID}

$drives -contains "C:"                 # True
$drives -contains "L:"                 # False

###########################################################################################################################################################################################

# Tip 5: Filter Out Unavailable Servers

# to filter lists with IP addresses and computer names. Check-Online will only let those pass a pipeline that can be pinged successfully
filter Check-Online
{
    trap {continue}
    .{
        $obj = New-Object System.Net.NetworkInformation.Ping
        $result = $obj.Send($_, 1000)
        
        if($result.Status -eq "Success")
        {
            $_
        }   
     }
}

"127.0.0.1","noexists","powershell.com" | Check-Online


# Create Hardware Inventory
Get-WmiObject -List                        # get WMI class info

Get-WmiObject Win32_BIOS
Get-WmiObject Win32_BIOS | Format-list *

"127.0.0.1", "server12", "pc-01-w3" | Check-Online | ForEach-Object {Get-WmiObject Win32_BIOS -ComputerName $_}               # use Check-Online to filter result


# Network Segment Scan
1..255 | ForEach-Object {"172.16.44.$_"} | Check-Online | ForEach-Object {[System.Net.Dns]::GetHostAddresses($_)}

###########################################################################################################################################################################################

# Tip 6: Parameters

Get-Help dir -Parameter *                                                                 # get all parameters detailed info related to dir

Get-Help dir -Parameter * | Where-Object {$_.Position -as [int]} | Sort-Object Position
# Where-Object checks whether the position property can be converted to an integer value. If so, it is a positional parameter. If not, it is a named parameter and excluded from the list

Get-Help dir -Parameter * | Format-Table Name, {$_.Description[0].Text} -Wrap             # get parameters and its description

# get property alias(or property shortcuts) for get-childitem
Get-Command Get-ChildItem | Select-Object -ExpandProperty ParameterSets | ForEach-Object {$_.Parameters} | 
    Where-Object {$_.Aliases -ne $null} | Select-Object Name, Aliases -Unique | Sort-Object Name

        # Name                                                                        Aliases                                                                                                   
        # ----                                                                        -------                                                                                                   
        # Debug                                                                       {db}                                                                                                      
        # Directory                                                                   {ad, d}                                                                                                   
        # ErrorAction                                                                 {ea}                                                                                                      
        # ErrorVariable                                                               {ev}                                                                                                      
        # File                                                                        {af}                                                                                                      
        # Hidden                                                                      {ah, h}                                                                                                   
        # LiteralPath                                                                 {PSPath}                                                                                                  
        # OutBuffer                                                                   {ob}                                                                                                      
        # OutVariable                                                                 {ov}                                                                                                      
        # ReadOnly                                                                    {ar}                                                                                                      
        # Recurse                                                                     {s}                                                                                                       
        # System                                                                      {as}                                                                                                      
        # UseTransaction                                                              {usetx}                                                                                                   
        # Verbose                                                                     {vb}                                                                                                      
        # WarningAction                                                               {wa}                                                                                                      
        # WarningVariable                                                             {wv}  

###########################################################################################################################################################################################

# Tip 7: Creating Custom Objects

$hash = @{}
$hash.Name = "Silence"
$hash.Age = 24
$hash.hadDog = $false

$object = New-Object PSObject -Property $hash
$object

###########################################################################################################################################################################################

# Tip 8: Using Comment-Based Help

function Test-Me($parameter)
{
    <#
        .SYNOPSIS
            A useless function
        .DESCRIPTION
            This function really does nothing
        .NOTES
            Demonstrates comment based help
        .LINK
            http://www.powershell.com
        .EXAMPLE
            Test-Me "Hello World"   
    #>
}

Get-Help Test-Me
Get-Help Test-Me -Examples

# Note: <# ... #> not provided in PowerShell v.1, so you should use "#" instead of "<# ... #>" in that version

###########################################################################################################################################################################################

# Tip 9: # Analyze Parameter Binding

Trace-Command -Name ParameterBinding {dir $env:windir *.* -Recurse} -PSHost               # This command analyzes the code in brackets

###########################################################################################################################################################################################

# Tip 10: Calculating Time Differences

# Note: New-Timespan automatically compares this date to the current date (and time) and returns the timespan in a number of different formats from which you can choose
(New-TimeSpan 12/24/2009).Days

# Note that dates that are in the future return negative timespans - unless you bind your date as -end parameter instead of -start
(New-TimeSpan -End 12/24/2009).Days
(New-TimeSpan -Start 12/24/2009).Days

(New-TimeSpan 8/28/1990).Days

###########################################################################################################################################################################################

# Tip 11: Add New Properties To Your Objects

dir | Select-Object *, Age | ForEach-Object {$_.Age = (New-TimeSpan $_.LastWriteTime).Days; $_} | Format-Table Name, Age  # use Select-Object to add new properties to an existing object

###########################################################################################################################################################################################

# Tip 12: Use Regular Expression to Validate Input

$pattern = '^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'

do
{
    $ip = Read-Host "Enter IP"

}while($ip -notmatch $pattern)

###########################################################################################################################################################################################

# Tip 13: Retrieving Error Messages From Your Event Logs

Get-EventLog -Newest 5 -LogName System -EntryType Error -ComputerName localhost | Format-Table eventid, message, mach*    # retrieves the last five error events from localhost

###########################################################################################################################################################################################

# Tip 14: Enabling PowerShell V.2 Remotely

Set-WSManQuickConfig                                                                                   # Setting up the brand new PowerShell v.2 remotely in a domain environment is easy

  # WinRM Quick Configuration
  # Running the Set-WSManQuickConfig command has significant security implications, as it enables remote management through
  #  the WinRM service on this computer.
  # This command:
  #  1. Checks whether the WinRM service is running. If the WinRM service is not running, the service is started.
  #  2. Sets the WinRM service startup type to automatic.
  #  3. Creates a listener to accept requests on any IP address. By default, the transport is HTTP.
  #  4. Enables a firewall exception for WS-Management traffic.
  #  5. Enables Kerberos and Negotiate service authentication.

# Note: As long as you have proper privileges to run this command, it does everything automatically: it runs the WinRM service, sets up a PowerShell listener and a firewall exception. 
 # To play with remotely, run this on every computer you want to connect to as it is required on both ends

Invoke-Command {Get-Process } -computername iis-cti5052

Invoke-Command {stop-Process -name notepad } -computername iis-cti5052                  
# Note: note that remote by default requires Kerberos authentication, so it will only work in a domain environment

###########################################################################################################################################################################################

# Tip 15: Remote Configuration in a Peer-to-Peer environment (or across domains)

# Note: By default, PowerShell requires Kerberos authentication to operate remotely, so you cannot use it in a simple peer-to-peer scenario. 
 # You can also not use it in a cross-domain scenario with untrusted domains. You will need to allow WSMan to use different authentication types to work remotely everywhere. 
 # All that is required is to add the IP addresses or computer names of computers you'd like to talk to. Note that this has to be done on both ends. 
 # The easiest (and most unsecure) way is to allow communication between any computer by specifying "*":
Set-Item WSMan:\localhost\Client\TrustedHosts * -Force

# Note: A more selective approach would use an IP address or computer name instead of "*". Once done, you can use all remote cmdlets to work remotely. 
 # Just make sure you use the -credential parameter to enter a User Name and Password for authentication

Invoke-Command {dir $env:windir} -computer iis-cti5052 -Credential (Get-Credential)

###########################################################################################################################################################################################

# Tip 16: Configuring WSMan Remotely for multiple computers

# Note: When working remotely in a peer-to-peer or cross-domain scenario, you will have to add all the computers you'd like to communicate with into the trusted hosts list. 
 # Unfortunately, when you try this, any new entry will overwrite the existing one, so there does not seem to be a way of adding multiple computer names
Set-Item WSMan:\localhost\client\trustedhosts 10.10.10.10 -force
Set-Item WSMan:\localhost\client\trustedhosts 10.10.10.11 -force

Get-Item WSMan:\localhost\client\trustedhosts

# As you see, only the last entry is stored in the list. You should make sure to run PowerShell with full administrative privileges
 # and run Set-WSManQuickConfig beforehand if you cannot run the commands at all because of an access denied-error

# Use the -Concatenate switch to add new entries without overwriting existing entries
Set-Item WSMan:\localhost\client\trustedhosts 10.10.10.10 -force -concatenate

###########################################################################################################################################################################################

# Tip 17: Running Commands On Multiple Computers

# run your commands against multiple computers 
$computer = "IIS-CTI5052"                                                   # $computer = get-content .\server.txt                # all computer names are read from a plain text file
$session = New-PSSession -ComputerName $computer

Invoke-Command -Session $session {md HKLM:\SOFTWARE\hey}
Invoke-Command -Session $session {Remove-Item HKLM:\SOFTWARE\hey}

# Note:  PowerShell limits the number of simultaneous connections to 32 (unless you change this with the -throttleLimit parameter) 
 # so if you should specify more than 32 computers, it will postpone additional calls

###########################################################################################################################################################################################

# Tip 18: About Registry

# Get-RegistryValues lists all values stored in a registry key
function Get-Registryvalues($key)
{
    (Get-Item $key).GetValueNames()
}

Get-Registryvalues "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion"



# Get-RegistryValue reads any value in any key and returns just the value
function Get-RegistryValue($key, $value)
{
    (Get-ItemProperty $key $value).$value
}

Get-RegistryValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' "RegisteredOwner"



# Adding or changing registry values
function Set-RegistryValue($key, $name, $value, $type = "String")
{
    if((Test-Path $key) -eq $false)
    {
        md $key | Out-Null
    }

    Set-ItemProperty $key $name $value -type $type
}

Set-RegistryValue HKCU:\Software\TestABC myValue Hello
Set-RegistryValue HKCU:\Software\TestABC myValue 12 Dword
Set-RegistryValue HKCU:\Software\TestABC myValue ([Byte[]][Char[]]"Hello") Binary

Remove-Item HKCU:\Software\TestABC -Force

###########################################################################################################################################################################################

# Tip 19: Using Real Registry Keys

dir HKLM:\Software
dir Registry::HKEY_LOCAL_MACHINE\Software


dir HKCU:\Software
dir Registry::HKEY_CURRENT_USER\Software

###########################################################################################################################################################################################

# Tip 20: Remove Registry Keys and Values

#  delete a registry key
function Remove-RegistryKey($key)
{
    Remove-Item $key -Force
}

# delete a registry value
function Remove-RegistryValue($key, $value)
{
    Remove-ItemProperty $key $value
}

###########################################################################################################################################################################################

# Tip 21: Dynamically Create Script Blocks

# retrieves the last five error records from that log

# Method 1:
function Get-RemoteLog($name)
{
    $code = "Get-EventLog -LogName $name -Newest 5 -EntryType Error"
    $sc = $ExecutionContext.InvokeCommand.NewScriptBlock($code)

    Invoke-Command -ComputerName myServer -ScriptBlock $sc
}

Get-RemoteLog "System"

# Note: Since the name of the log submitted in $name needs to be sent to the remote system, the function first has to resolve this variable locally 
 # before it creates the script block and sends it to the remote machine. If you had submitted the code directly, 
 # the remote machine would try and resolve the $name variable, which would not exist on the remote system

# Method 2: 
function Get-RemoteLog($name)
{
    Invoke-Command -ComputerName myServer -ScriptBlock {
        param($logname)

        Get-EventLog -LogName $logname -Newest 5 -EntryType Error

    } -ArgumentList $name
}

Get-RemoteLog "System"

###########################################################################################################################################################################################

# Tip 22: Get Remotely, Store Locally

Invoke-Command -ComputerName remotepc { Get-Process | Select-Object Name | Export-Csv $home\result.csv }             # the csv file with your results is stored on the remote system

Invoke-Command -ComputerName remotepc { Get-Process | Select-Object Name } | Export-Csv $home\result.csv             # to store the remote executed results on your machine

###########################################################################################################################################################################################

# Tip 23: Exit a Script Immediately

# Use the Exit statement if you run into a condition where a script should quit. The script breaks whenever you call exit from within your script. 
 # Add a number to Exit to set the error level (return code). The caller of your script can then check the automatic
 # variable $LASTEXITCODE to see if the script ended prematurely and determine its exit code.

 exit 411
 
 $LASTEXITCODE

###########################################################################################################################################################################################

# Tip 24: Exporting and Importing Credentials

# to save your credential to file and re-use this saved version for unattended scripts

function Export-Credential($cred, $path)
{
    $cred = $cred | Select-Object *
    $cred.password = $cred.password | ConvertFrom-SecureString
    $cred | Export-Clixml $path
}

function Import-Credential($path)
{
    $cred = Import-Clixml $path
    $cred.password = $cred.password | ConvertTo-SecureString

    New-Object System.Management.Automation.PSCredential($cred.username, $cred.password)
}

Export-Credential (Get-Credential) $home\cred.xml

Get-WmiObject Win32_BIOS -ComputerName iis-cti5052 -Credential (Import-Credential $home\cred.xml)  
# Note that your Password is encrypted and can only be imported by the same user that exported it

###########################################################################################################################################################################################

# Tip 25: Creating New Objects

# One: Using Add-Member
$object = New-Object PSObject
$object | Add-Member -MemberType NoteProperty -Name FirstName -Value Silence -PassThru | 
    Add-Member -MemberType NoteProperty -Name LastName -Value He -PassThru | 
    Add-Member -MemberType NoteProperty -Name Age -Value 24 -PassThru
$object


# Three: Using Select-Object
$object = "" | Select-Object FirstName, LastName, Age
$object.FirstName = "Silence"
$object.LastName = "He"
$object.Age = 24
$object


# Four: Using HashTable
$hash = @{}
$hash.FirstName = "Silence"
$hash.LastName = "He"
$hash.Age = 24
$object = New-Object PSObject -Property $hash
$object


$object = New-Object PSObject -Property @{ FirstName = "Silence"; LastName = "He"; Age = 24}
$object


$object = New-Object PSCustomObject -Property ([ordered]@{ FirstName = "Silence"; LastName = "He"; Age = 24 })
$object


$object = [PSCustomObject]@{ FirstName = "Silence"; LastName = "He"; Age = 24 }
$object


$object = [PSCustomObject][ordered]@{ FirstName = "Silence"; LastName = "He"; Age = 24 }
$object

# To see the performance for these all method, go to site: http://www.pstips.net/performance-of-custom-psobject.html

###########################################################################################################################################################################################

# Tip 26: Use Hash Tables for Custom Columns

$age = @{label = "Age"; Expression = {(New-TimeSpan $_.LastWriteTime).Days}}

dir $env:windir | Format-Table Name, $age

dir $env:windir | Select-Object Name, $age | Export-Csv $home\test.csv               # Select-Object can add new properties to objects and fill them with calculated content

###########################################################################################################################################################################################

# Tip 27: Adding File Age to file objects

dir $env:windir | Format-List *                                                      # to see all the information, you can pipe the result to Format-List *

# fileAge.ps1xml
'
<Types>
  <Type>
      <Name>System.IO.FileInfo</Name>
      <Members>
          <ScriptProperty>
              <Name>Age</Name>
              <GetScriptBlock>(New-TimeSpan $this.LastWriteTime).Days</GetScriptBlock>
          </ScriptProperty>
      </Members>
  </Type>
</Types>
' | Out-File $home\fileAge.ps1xml

Update-TypeData $home\fileAge.ps1xml                                                  # update PowerShell using this command

dir $env:windir | Format-Table Name, Age                                              # after update powershell, the Age perproty point to (New-TimeSpan $this.LastWriteTime).Days

dir $env:windir -Filter *.log -Recurse | Where-Object {$_.Age -lt 30} | Format-Table FullName, Age             # selects all logs which have been modified within the last 30 days

###########################################################################################################################################################################################

# Tip 28: Change Service Account Password

# to automatically change the Password a service uses to log on to its account
$localAccount = ".\Administrator"
$newPassword = "secret@password"

$service = Get-WmiObject Win32_Service -Filter "Name='Spooler'"
$service.Change($null,$null,$null,$null,$null,$null,$localAccount,$newPassword)

# Note: These lines assign a new User Account and Password for the Spooler service. Note that your account will need special privileges to be able to do that

###########################################################################################################################################################################################

# Tip 29: Get Process Owners

$processes = Get-WmiObject Win32_Process -Filter "name='notepad.exe'"

$appendeProcessess = foreach ($process in $processes){
    Add-Member -MemberType NoteProperty -Name Owner -Value ($process.GetOwner().User) -InputObject $process -PassThru
}

$appendeProcessess | Format-Table Name, Owner
# A loop adds the necessary Owner property. The resulting objects now show the owner. You could then use the information in Owner to selectively stop all processes owned by a specific user


function Get-ProcessOwner($processName)
{   
    $processes = Get-Process $processName -ErrorAction SilentlyContinue

    if($processes -eq $null) {retrun}

    $processes | ForEach-Object {
        Get-WmiObject Win32_Process -Filter ("Handle={0}" -f $_.Id) | 
        ForEach-Object {Add-Member -InputObject $_ -MemberType NoteProperty -Name Owner -Value ($_.GetOwner().User) -PassThru} | Select-Object Name, Handle, Owner
    }
}

Get-ProcessOwner explorer
Get-ProcessOwner exp*

###########################################################################################################################################################################################

# Tip 30: Opening Databases from PowerShell

# Note: The easiest way of accessing databases right from PowerShell is to visit control panel and open the Data Sources (ODBC) module
 # (which resides in Administrative Tools inside control panel). Use the GUI to set up the database type by clicking the "User DN" or "System DN" tab and then click Add. 
 # Remember to write down the "Data Source name" you assign because that's the only thing PowerShell is going to need

# Control Panel\All Control Panel Items\Administrative Tools  => Data Sources (ODBC)
$connection = New-Object -ComObject ADODB.Connection
$connection.Open("myDataSource")
$objRS = $connection.Execute("select * from tablename")

while($objRS.EOF -ne $true)
{
    foreach($field in $objRS.Fields)
    {
        "{0,30} = {1, -30}" -f $field.name, $field.value
    }
    ""
    $objRS.MoveNext()
}


# Converting Database Records into PowerShell objects
$connection = New-Object -ComObject ADODB.Connection
$connection.Open("myDataSource")
$objRS = $connection.Execute("select * from tablename")

while($objRS.EOF -ne $true)
{
    $hash = @{}

    foreach($field in $objRS.Fields)
    {
        $hash.$($field.name) = $field.value
    }
    $hash
    New-Object PSObject -Property $hash

    $objRS.MoveNext()
}

###########################################################################################################################################################################################

# Tip 31: Adding Write Protection to functions

function ImportantFunction
{
    "You can not overwrite me!"
}
 
(dir function:ImportantFunction).Options = "ReadOnly"                            # make a function "read-only"

function ImportantFunction                                                       # Exception here: Cannot write to function ImportantFunction because it is read-only or constant
{
    "test if it can be owerwrite"
}


# Creating "Constant" Functions
New-Item -Path function: -Name constantFunction -Options Constant -Value {
    "You can't get rid of me except by closing powershell ..."
}

# Note: When you make a function read-only, it can no longer be overwritten but you would still be able to delete the function and recreate it from scratch. 
 # You can make them constant if you'd like to create functions that cannot be changed as long as the PowerShell session runs

###########################################################################################################################################################################################

# Tip 32:　Resolve Host Names

function Get-HosttoIP($hostname)
{
    $result = [System.Net.Dns]::GetHostByName($hostname)
    $result.AddressList | ForEach-Object {$_.IPAddressToString}
}

Get-HosttoIP "www.baidu.com"                                                     # output: 180.76.3.151

###########################################################################################################################################################################################

# Tip 33: Speeding Up Remote Inventory

Get-Content $home\serverlist.txt | ForEach-Object {Get-WmiObject Win32_BIOS -ComputerName $_}  
# Note: this approach may take a long time since processing is sequentially, especially when systems are offline

# Solution: run it as separate background jobs
Get-Content $home\serverlist.txt | ForEach-Object {Get-WmiObject Win32_BIOS -ComputerName $_ -AsJob}

# Note: Background jobs require PowerShell v.2, and they also require remote setup. Even though a standalone Get-WMIObject cmdlet does not require remote setup when using -computername, 
 # it *is* required when running as background job. So, all remote machines must also use PowerShell v.2 and be configured to work remotely

###########################################################################################################################################################################################

# Tip 34: Validating Input Type

# The -as parameter is not widely known but is extremely versatile. It tries to convert data into a .NET type, and when it fails, it simply returns $null.

function Test-Numeric($test)                 # Test-Numeric function to validate whether someone has entered a numeric value
{
    ($test -as [Double]) -ne $null
}

Test-Numeric a                               # False
Test-Numeric 28                              # True

###########################################################################################################################################################################################

# Tip 35: Getting Hotfix Information

Get-HotFix

Get-HotFix -ComputerName localhost

Get-WmiObject Win32_QuickFixEngineering

###########################################################################################################################################################################################

# Tip 36: Creating Temporary File Names

$tempfileName = (Get-Date -Format "yyyy-MM-dd hh-mm-ss") + ".tmp"
$tempfileName

###########################################################################################################################################################################################

# Tip 37: Using Relative Dates

# Method 1:
$at14daysago = (Get-Date) - (New-TimeSpan -Days 14)
$at14daysago


# Method 2:
(Get-Date).AddDays(-14)

###########################################################################################################################################################################################

# Tip 38: Listing Program Versions

Get-process -FileVersionInfo -ErrorAction SilentlyContinue                                               # get the files (and versions) the running processes are using

###########################################################################################################################################################################################

# Tip 39: About Get-Help

# Note: Many cmdlets support the -force switch parameter. With it, the cmdlet will do more than usual. What exactly -force does depends on the cmdlet. 
 # For example, Get-Childitem (aka Dir) by default does not list hidden files. However, with -force, it does. Another example: Stop-Computer shuts down a computer only if all data is saved. 
 # By specifying -force, the machine will shut down immediately

Get-Help * -Parameter force                                                                             # To list all cmdlets that support -force


# Note: Find Potentially Harmful Cmdlets
# Any cmdlet that can change and potentially damage your system supports the -whatif parameter, allowing you to just simulate the action without actual change

Get-Help * -Parameter whatif                                                                            # to locate all cmdlets that you should be careful with


# Getting Advanced Help
Get-Help about_*                                                                                        # to find out more about general PowerShell concepts

help wildcard                                                                                           # it is sufficient to specify a unique keyword to access the Help information

Get-Help about_*operator*                                                                               # to get you all topics that deal with operators

Get-Help about_*operator* | Get-Help | Out-File $env:windir\myhelp.txt                                  # to open all of these Help files and create your own Help file

###########################################################################################################################################################################################

# Tip 40:　Childproofing PowerShell

# Note: If you are new to PowerShell, you may be worried about causing unwanted damage or change. One way of childproofing PowerShell is by changing the whatif-default like so

$WhatIfPreference = $true

# From now on, any cmdlet supporting the -whatif parameter will use it without you having to specify it. So any cmdlet that would change things on your system is effectively disabled. 

md C:\Test                                                                                              # a new folder is not allowed anymore if $WhatIfPreference = $true

md C:\Test -WhatIf:$false                                                                               # To override the default, you need to explicitly add this option: -whatif:$false

###########################################################################################################################################################################################

# Tip 41: Visiting Help Topics

# Note: All Help topics inside of PowerShell are stored as plain text files. You can read them by using Get-Help, 
 # but you can also more easily open the appropriate folder in Explorer and do whatever you like

explorer "$pshome\$($host.CurrentCulture.Name)"

###########################################################################################################################################################################################

# Tip 42: Copying Help Information (or other things) to Clipboard

# Note: Most people aren't aware that Windows comes with a small application called clip.exe (introduced in Windows Vista). 

Get-Help about_commonparameters | clip                                                                 # within the PowerShell pipeline to copy information to the clipboard

###########################################################################################################################################################################################

# Tip 43: Secret Parameter Alias Names

'Get-ChildItem' | ForEach-Object {(Get-Command $_).Parameters | ForEach-Object {$_.Values |Where-Object {$_.Aliases.Count -gt 0}| Select-Object Name, Aliases } }

###########################################################################################################################################################################################

# Tip 44: Uncovering Parameter Binding

# You should use Trace-Command: if you are ever in doubt about just how PowerShell binds cmdlet parameters to a cmdlet 
Trace-Command -PSHost -Name ParameterBinding {Get-ChildItem $env:windir *.log}

###########################################################################################################################################################################################

# Tip 45: Finding Cmdlets by Keyword

function ??($keywords)
{
    Get-Help * | Where-Object {$_.Description -like "*$keywords*"} | Select-Object Name, Synopsis
}

?? shutdown
?? print
?? random

###########################################################################################################################################################################################

# Tip 46: Remote Access Without Admin Privileges

# Note: In PowerShell v.2, remote access is available only to users who hold local administrator privileges. So, even if you do have appropriate remote access to a machine,
 # you cannot remotely access the system if you are not an Admin. This is not a technical limitation, though, just a safe default. 

Set-PSSessionConfiguration -Name Microsoft.PowerShell -ShowSecurityDescriptorUI                    # change the defaul value and get remote access without admin privileges

###########################################################################################################################################################################################

# Tip 47: Out-GridView Dirty Tricks

Get-Process | Out-GridView                                                                         # output objects to a "mini" excel sheet
Get-Process | Select-Object * | Out-GridView

# Note: this only works if .NET Framework 3.51 is installed

###########################################################################################################################################################################################

# Tip 48: Listing Installed Software

Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, HelpLink, UninstallString

Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, HelpLink, UninstallString | Out-GridView

# Note: listing installed software can be somewhat difficult as WMI provides the Win32_Product class, which only covers managed installs (installed by MSI). 
 # You should consider reading the registry, which is a better approach. One little known fact is that Get-ItemProperty (used to read registry values) accepts wildcards

###########################################################################################################################################################################################

# Tip 49: Changing Execution Policy without Admin Rights

# Note: In PowerShell v.2, a number of parameters have been added to Set-ExecutionPolicy, 
 # which allows you to change this setting without Admin privileges, unless an administrator has restricted this
Set-ExecutionPolicy -Scope Process RemoteSigned -Force                 # change the execution policy only for the current session

Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force             # change your personal execution policy permanently

###########################################################################################################################################################################################

# Tip 50: Filter is Faster Than Include

# If you have a choice, you should always pick -filter. For starters, it is much faster (4x and more) and secondly, -include only works when combined with -recurse. -filter always works.

dir $env:windir -Filter *.log
dir $env:windir -Include *.log -Recurse

# So why is there -include at all? Because not all providers support -filter. For example, you are limited to -include when you list registry content

###########################################################################################################################################################################################

# Tip 51: Open Current Folder in Your Explorer

explorer .
Invoke-Item .                                     # or: ii .

###########################################################################################################################################################################################

# Tip 52: Launching Programs

Start-Process notepad -WindowStyle Maximized      # to launch notepad maximized 

# Note: Supported arguments are Maximized, Minimized, Normal, and Hidden. Be sure to watch out with Hidden! You should only use it for programs that do not require user interaction

Start-Process notepad -Wait                       # to launch a Windows application and wait for it until it finishes

###########################################################################################################################################################################################

# Tip 53: Use CHOICE to Prompt for Input

choice /N /C:123 /M "Enter a number between 1 and 3!"
"Your choice: $LASTEXITCODE"

choice /N /C:YN /M "Do you agree (Y/N)?"
"Your answer: $LASTEXITCODE"

###########################################################################################################################################################################################

# Tip 54: Stopping a Program Whenever You Feel Like It

$notepad = Start-Process notepad -PassThru
Start-Sleep 2
Stop-Process -InputObject $notepad

# Note: When you launch a program using Start-Process with -passThru, you will get back the process object representing the started program. 
 # You can then use this object later to kill the program whenever necessary

###########################################################################################################################################################################################

# Tip 55: Closing a Program Gracefully

Get-Process notepad | Stop-Process                                          # it will stop instantaneously. The user will get no chance to save unsaved documents

Get-Process notepad | ForEach-Object {$_.CloseMainWindow() | Out-Null}      # If the content has not been saved, a dialog opens and asks for a choice


#The user will then get 15 seconds to save all unsaved data. All processes that still run after 15 seconds will be killed
$proc = Get-Process notepad -ErrorAction SilentlyContinue
$proc | ForEach-Object {$_.CloseMainWindow()} | Out-Null
Start-Sleep -Seconds 15
$proc | Where-Object {$_.HasExited -ne $true} | Stop-Process -ErrorAction SilentlyContinue -WhatIf 

###########################################################################################################################################################################################

# Tip 56: Accessing Profile Scripts

# Method One:
$profile | Get-Member *Host* | ForEach-Object {$_.Name} | ForEach-Object {
    $rv = @{}
    $rv.Name = $_
    $rv.Path = $profile.$_
    $rv.Exists = (Test-Path $profile.$_)

    New-Object PSObject -Property $rv

} | Format-Table -AutoSize


# Method Two:
$profile.psextended | Get-Member | Select-Object Name, @{ n = "Path"; e = {$profile.($_.name)}}, @{n = "Exists"; e = {Test-Path $profile.($_.name)}}


# Method Three:
$profile.psextended | Format-List *

###########################################################################################################################################################################################

# Tip 57: Running PowerShell Scripts as Scheduled Task

# Note: If you have jobs that need to execute regularly, you can manage them with a PowerShell script and make it a scheduled task
schtasks /CREATE /TN CheckHealthScript /TR "powershell.exe -noprofile -executionpolicy unrestricted -file $env:windir\test.ps1" /IT /RL HIGHEST /SC DAILY

schtasks /DELETE /TN CheckHealthScript                                                                                   # To remove the scheduled task, specify the name you assigned

###########################################################################################################################################################################################

# Tip 58: Creating Large Dummy Files

fsutil file createnew $env:windir\dummy.bin 1gb                       # Note:  fsutil requires admin privileges


# Creating Large Dummy Files With .NET
$path = "$env:temp\testfile.txt"
$file = [io.file]::Create($path)
$file.SetLength(1gb)
$file.Close()

Get-Item $path

###########################################################################################################################################################################################

# Tip 59: Stopping the Pipeline

# Usually, once a pipeline runs, you cannot stop it prematurely, even if you already received the information you were seeking
filter Stop-Pipeline([scriptblock]$condition = {$true})
{
    $_

    if(& $condition)
    {
        continue
    }
}

Get-EventLog Application | Stop-Pipeline {$_.InstanceId -gt 10000}        # get all Application events until you find one with an InstanceID of greater than 10,000


# If you'd like to use this technique inside a script, you will need to make sure Stop-Pipeline does not stop your entire script. 
 # You can do that by embedding the pipeline inside a dummy loop
$result = do {   
    Get-EventLog Application | Stop-Pipeline {$_.InstanceId -gt 10000}
}while($false)

###########################################################################################################################################################################################

# Tip 60: Hide Error Messages

dir $env:windir *.log -Recurse                                            # Exception here

dir $env:windir *.log -Recurse -ErrorAction SilentlyContinue              # Hide Error Messages

dir $env:windir *.log -Recurse -ea 0                                      # Hide Error Messages

###########################################################################################################################################################################################

# Tip 61:　Discovering Impact Level

# get you that information so you can see how severe the action is that a given cmdlet will take
Get-Command -CommandType Cmdlet | ForEach-Object {
    $_.ImplementingType.GetCustomAttributes($true) | 
    Where-Object {$_.VerbName -ne $null} | 
    Select-Object @{Name="Name"; Expression = {"{0} - {1}" -f $_.VerbName, $_.NounName}}, ConfirmImpact | Sort-Object ConfireImpact -Descending
}

###########################################################################################################################################################################################

# Tip 62: ExpandProperty to the Rescue

Get-Process | Select-Object Name, Company                      # Select-Object is often used to select object properties and discard unneeded information

Get-Process | Select-Object Name                               # get objects with a name property (which is why PowerShell displays a column header)

Get-Process | Select-Object -ExpandProperty Name               # add -ExpandProperty if you just want the names as strings without column header

###########################################################################################################################################################################################

# Tip 63: Adding New Virtual Drives

dir Registry::HKEY_CLASSES_ROOT

New-PSDrive HKCR Registry Registry::HKEY_CLASSES_ROOT
dir HKCR:

###########################################################################################################################################################################################

# Tip 64: Resolve Paths Gets Lists of Paths

Resolve-Path $env:windir\*.log                                 # use Resolve-Path to support wildcards so you can use it to easily put together a list of file names 

###########################################################################################################################################################################################

# Tip 65: Reading Default Values

# Each registry key has a default value. It is called (default) when you look into regedit. To read this default value, you can just use this name (and put it into quotes)

Get-ItemProperty Registry::HKEY_CLASSES_ROOT\.ps1 "(Default)" | Select-Object -ExpandProperty "(Default)"           # to find out the currently associated program for PowerShell scripts

###########################################################################################################################################################################################

# Tip 66: Adding New Snapins and Modules

# Snapins and modules can add new cmdlets and/or providers to PowerShell
Get-PSSnapin -Registered                                                             # to see the list of available Snapins

dir HKLM:\SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns | Select name            # PowerShell peeks at the following registry location when exec "Get-PSSnapin -Registered"


Get-Module -ListAvailable | Select-Object Name

# Note: This is where all snapins need to register themselves or else PowerShell cannot load them. In contrast, modules do not need to register themselves, 
 # which is why you can simply copy and load them. You can use Get-Module -ListAvailable and Import-Module accordingly
 
###########################################################################################################################################################################################

# Tip 67: List Installed Updates

Get-Content $env:windir\WindowsUpdate.log -encoding utf8 | Where-Object {$_ -like "*successfully installed*"} | 
    ForEach-Object {

        $infos = $_.split("`t")
        $result = @{}
        $result.Date = [DateTime]$infos[6].Remove($infos[6].LastIndexOf(":"))
        $result.Product = $infos[-1].SubString($infos[-1].LastIndexOf(":") + 2)

        New-Object PSObject -Property $result
    } | Format-Table -AutoSize

###########################################################################################################################################################################################

# Tip 68: WMI Server Side Filtering

Get-WmiObject Win32_Service | Where-Object {$_.started -eq $false} | Where-Object {$_.startmode -eq "Auto"} | 
    Where-Object {$_.exitCode -ne 0} | Select-Object Name, DisplayName, StartMode, ExitCode

# A faster and better approach is to filter on the WMI end by using the query parameter:
Get-WmiObject -Query ("select Name, DisplayName, StartMode, ExitCode from win32_service where Started=false and StartMode='Auto' and ExitCode<>0")

###########################################################################################################################################################################################

# Tip 69: Create HTA Files

Get-Process | ConvertTo-Html | Out-File $home\result.hta
Invoke-Item $home\result.hta

# Note: ConvertTo-HTML is a convenient way of converting object results in HTML. However when you open these files, 
 # your browser starts with all the bells and whistles. You should try a better way by storing the HTML in a file with HTA extension

###########################################################################################################################################################################################

# Tip 70: Generate PC Lists

1..40 | ForEach-Object {"PC-W7-A{0:000}" -f $_}                            # creating lists of PC names or IP address ranges etc, creating lists of PC names or IP address ranges etc 

###########################################################################################################################################################################################

# Tip 71: Getting Assigned IP Addresses

Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.IPEnabled -eq $true} | ForEach-Object {$_.IPAddress}

Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled=$true" | ForEach-Object {$_.IPAddress}                            #  let GWMI filter more efficiently

# Note: You can also access remote systems and check their IP addresses since Get-WMIObject supports the -ComputerName parameter



Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.IPEnabled -eq $true} | ForEach-Object {$_.IPAddress} | 
    ForEach-Object {[IPAddress]$_} | Where-Object {$_.AddressFamily -eq "INternetwork"} | ForEach-Object {$_.IPAddressToString}      # to get all IPv4 addresses assigned to your system

###########################################################################################################################################################################################

# Tip 72: Create Groups (and Examine Alias groups)

Get-Alias | Format-Table -GroupBy Definition

Get-Alias | Select-Object Definition | Format-Table -GroupBy Definition

Get-Alias | Group-Object Definition

# Note: Format-Table has a parameter called -GroupBy which creates groups based on the property you supply

###########################################################################################################################################################################################

# Tip 73: Multi-Column Lists

Get-Service                                 # Default: Status | Name | DisplayName

Get-Service | Select-Object Name            # only output service name

Get-Service | Format-Wide -Column 5         # format all output with 5 columns

###########################################################################################################################################################################################

# Tip 74: Examining Object Data

get-process -id $pid
get-process -id $pid | Select-Object *                    # to see all properties a result object provides, you should probably add Select-Object * to it 

get-process -id $pid | Format-Custom * -Depth 5           # you can specify a depth, which will allow you to see nested object properties down to the depth you specified

###########################################################################################################################################################################################

# Tip 75: Outputting Text Data to File

# When you output results to text files, you will find that the same width restrictions apply that are active when you output into the console. 
 # You should use a combination of Format-Table and Out-File with -Width to allow more width

Get-Process | Format-Table Name, Description, Company, StartTime -AutoSize | Out-File $home\test.txt -Width 1000
Invoke-Item $home\test.txt

# Note: Format-Table -AutoSize uses only as much width as is needed to display all information, and Out-File -Width specifies the maximum width that Format-Table can use

###########################################################################################################################################################################################

# Tip 76: Printing Results

Get-WmiObject Win32_Printer | Select-Object -ExpandProperty Name              # find out the names of your installed printers

Get-Process | Out-Printer -Name "Microsoft XPS Document Writer"               # print results to your default printer, You can also print to any other printer when you specify its name

###########################################################################################################################################################################################

# Tip 77: Create HTML System Reports

$style = @"
<style>
  body { background-color:#EEEEEE; }
  body,table,td,th { font-family:Tahoma; color:Black; Font-Size:10pt }
  th { font-weight:bold; background-color:#AAAAAA; }
  td { background-color:white; }
</style>
"@

& {
     "<HTML><HEAD><TITLE>Inventory Report</TITLE>$style</HEAD>"
     "<BODY><h2>Report for '$env:computername'</h2><h3>Services</h3>"
     
     Get-Service | Sort-Object Status, DisplayName | ConvertTo-HTML DisplayName, ServiceName, Status -Fragment 

     '<h3>Operating System Details</h3>'

     Get-WMIObject Win32_OperatingSystem | Select-Object * -exclude __* |   ConvertTo-HTML -as List -Fragment

     '<h3>Installed Software</h3>'

     Get-ItemProperty Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | 
     Select-Object DisplayName, InstallDate, DisplayVersion, Language | Sort-Object DisplayName | ConvertTo-HTML -Fragment 
     
     '</BODY></HTML>'

   } | Out-File $home\report.hta

Invoke-Item $home\report.hta

###########################################################################################################################################################################################

# Tip 78: Built-in Methods

"Hello" | Get-Member -MemberType *method

"   Hello   ".Trim()                                                # Output: Hello



# to split a path into its segments, you should use Split():        
"c:\test\subfolder\file.txt".Split("\")                             # Output: c:   test   subfolder   file.txt
"c:\test\subfolder\file.txt".Split("\")[0]                          # Output: C:
"c:\test\subfolder\file.txt".Split("\")[-1]                         # Output: file.txt
"c:\test\subfolder\file.txt".Split(".")[-1]                         # Output: txt

###########################################################################################################################################################################################

# Tip 79: Finding Object Types

dir $env:windir | Get-Member | Group-Object TypeName                # pipe the result of a command to Get-Member, it examines the returned objects and shows the properties and methods

dir $env:windir | Get-Member | Group-Object TypeName -NoElement     # only to focus on the different object types a command returns

###########################################################################################################################################################################################

# Tip 80: Filtering Day of Week

function is-Weekend
{
    (Get-Date).DayOfWeek -gt 5
}

if(is-Weekend)
{
    "No service at weekend!"
}
else
{
    "Heya, time to work!"
}

###########################################################################################################################################################################################

# Tip 81: Changing Error Background Color

[System.Enum]::GetNames([System.ConsoleColor])                    # to see the colors that you can assign


# If you would like to make error messages more readable, you can change their background color from black to white
$host.PrivateData.ErrorBackgroundColor = "White"

# To use a transparent error background color, you should assign the current console background color
$host.PrivateData.ErrorBackgroundColor = $host.UI.RawUI.BackgroundColor

###########################################################################################################################################################################################

# Tip 82: Changing File/Folder Creation Date

Get-ChildItem $env:windir\test | Select-Object CreationTime                                      # get the original time for the files

Get-ChildItem $env:windir\test | ForEach-Object {$_.CreationTime = "2014/09/02 15:12"}           # update the create time

Get-ChildItem $env:windir\test | ForEach-Object {$_.LastWriteTime = "2014/09/02 15:12"}          # update the last write time

###########################################################################################################################################################################################

# Tip 83: Accessing Object Properties

# Note: Objects store information in various properties. There are two approaches if you would like to get to the content of a given property.

# Method 1: 
(Get-Process -Id $pid).CPU


# Method 2:
Get-Process -Id $pid | Select-Object -ExpandProperty CPU

###########################################################################################################################################################################################

# Tip 84: How Much RAM Do You Have?

# Ever wondered what type of RAM your PC uses and if there is a bank available to add more? Ask WMI! 
 # This example also converts the cryptic type codes into clear-text using hashtables to create new columns

Get-WmiObject Win32_PhysicalMemory | Select-Object FormFactor, MemoryType | Format-Table -AutoSize

# Output:
  # FormFactor MemoryType
  # ---------- ----------
  #          8          0
  #          8          0
  #          0         11

$memorytype = "Unknown", "Other", "DRAM", "Synchronous DRAM", "Cache DRAM", "EDO", "EDRAM", "VRAM", "SRAM", "RAM", "ROM", 
              "Flash", "EEPROM", "FEPROM", "EPROM", "CDRAM", "3DRAM", "SDRAM", "SGRAM", "RDRAM", "DDR", "DDR-2"

$formfactor = "Unknown", "Other", "SIP", "DIP", "ZIP", "SOJ", "Proprietary", "SIMM", "DIMM", "TSOP", "PGA", "RIMM", "SODIMM", 
              "SRIMM", "SMD", "SSMP", "QFP", "TQFP", "SOIC", "LCC", "PLCC", "BGA", "FPBGA", "LGA"

$col1 = @{Name = "Size(GB)"; Expression = {$_.Capacity / 1GB}}
$col2 = @{Name = "Form Factor"; Expression = {$formfactor[$_.FormFactor]}}
$col3 = @{Name = "Memory Type"; Expression = {$memorytype[$_.MemoryType]}}

Get-WmiObject Win32_PhysicalMemory | Select-Object BankLabel, $col1, $col2, $col3

Get-WmiObject Win32_PhysicalMemory -ComputerName remoteServer | Select-Object BankLabel, $col1, $col2, $col3                # -ComputerName make it possible to run on remote server   

###########################################################################################################################################################################################

# Tip 85: Adding Extra Information

# Note: Sometimes you may want to tag results returned by a cmdlet with some extra information, such as a reference to some PC name or a timestamp. 
 # You can use Add-Member to tag a note property to the result.

# gets all services and adds two new columns: the PC name, and the date and time the results were returned
$result = Get-Service | Add-Member NoteProperty Computer $env:COMPUTERNAME -PassThru | Add-Member NoteProperty TimeStamp (Get-Date) -PassThru
$result | Select-Object Status, Name, Computer, Time*

###########################################################################################################################################################################################

# Tip 86: Discover Hidden Object Members

"Hello" | Get-Member                        # Get-Member is a great cmdlet to discover object members, but it will not show everything

"Hello" | Get-Member -Force                 # add -force to really see a complete list of object members

"Hello".PSTypeNames                         # One of the more interesting hidden members is called PSTypeNames and lists the types this object was derived from
# Output:
  # System.String
  # System.Object

# These types become important when PowerShell formats the object. 
 # The first type in that PSTypeNames-list that can be found in PowerShell’s internal type database is used to format the information

$a = "Hello"
$a.PSTypeNames.Clear()
$a.PSTypeNames
$a

$a -eq "Hello"

# Note: Without PSTypeNames, PowerShell no longer knows how to format the string and will display empty space instead.

###########################################################################################################################################################################################

# Tip 87: Replace Text in Files

Get-Content $home\test.txt | ForEach-Object {$_ -replace "Old", "New"} | Set-Content -Path $home\test.txt                  # Exception here: because it is being used by another process

# Solution: use parenthesis so that PowerShell reads the file first and only and then processes the content
(Get-Content $home\test.txt) | ForEach-Object {$_ -replace "Old", "New"} | Set-Content -Path $home\test.txt     
start $home\test.txt

###########################################################################################################################################################################################

# Tip 88: Grouping with Script Blocks

# Group-Object creates groups based on object properties
Dir $env:windir | Group-Object Extension
Get-Process | Group-Object Company


dir $env:windir | Group-Object { if($_.Length -gt 100kb) {"Largs"} else {"Small"}}            # If the object has no property that reflects your grouping needs, you can create one

###########################################################################################################################################################################################

# Tip 89: Accessing WMI Instances in One Line

[WMI]"Win32_LogicalDisk='C:'" | Select-Object *

# Note: While you can use Get-WMIObject to query for WMI objects and then select the ones you are really after, 
 # you can also cast a WMI object path to a WMI object and get to that instance immediately

###########################################################################################################################################################################################

# Tip 90: E-mail-Address-Extractor via RegEx

$regex = [RegEx]'(?i)\s[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\s'

$text = @"
sdfwefswe
abc@23.com
bas@hotmail.com
ihe@microsoft.com
hhbsar@hotmail.com
2342.232@fwe.com
"@

$regex.Matches($text) | Select-Object -ExpandProperty Value

###########################################################################################################################################################################################

# Tip 91: Calculating Server Uptime

function Get-UpTime                        #  to find out the effective uptime of your PC (or some servers)
{
    param($computerName = "localhost")

    Get-WmiObject Win32_NTLogEvent -Filter "Logfile='System' and EventCode>6004 and EventCode<6009" -ComputerName $computerName | 
        ForEach-Object {
            
            $rv = $_ | Select-Object EventCode, TimeGenerated

            switch($_.EventCode)
            {
                6006 {$rv.EventCode = "shutdown"}
                6005 {$rv.EventCode = "start"}
                6008 {$rv.EventCode = "crash"}
            }

            $rv.TimeGenerated = $_.ConvertToDateTime($_.TimeGenerated)
            $rv
        }
}

Get-UpTime

###########################################################################################################################################################################################

# Tip 92: Creating Your own Types

# Note: you can compile any .NET source code on the fly and use this to create your own types.
 # Here is an example illustrating how to create a new type from c# code that has both static and dynamic methods

$source = @"
public class Calculator
{
    public static int Add(int a, int b)
    {
        return (a + b);
    }
    public int Multiply(int a, int b)
    {
        return (a * b);
    }
}
"@

Add-Type -TypeDefinition $source

[Calculator]
[Calculator]::Add(5, 10)                                                   # static method can be called with class name
                                                                           
$myCalculator = New-Object Calculator                                      # call non-static method, you should new a instance first
$myCalculator.Multiply(3, 12)

###########################################################################################################################################################################################

# Tip 93: Accessing Function Parameters by Type

# a function that accepts both numbers and text, and depending on which type you submitted, choose a parameter
function Test-Binding                                                      # to assign values to different parameters based solely on type
{
    [CmdletBinding(DefaultParameterSetName='Name')]
    param(
    [Parameter(ParameterSetName='ID', Position=0, Mandatory=$true)]
    [int]
    $id,

    [Parameter(ParameterSetName='Name', Position=0, Mandatory=$true)]
    [String]
    $name
    )

    $set = $PSCmdlet.ParameterSetName

    if($set -eq "ID")
    {
        "The numeric ID is $id"
    }
    else
    {
        "You entered $name as a name"
    }
}

Test-Binding 12
Test-Binding "hello"
Test-Binding

###########################################################################################################################################################################################

# Tip 94: Store Pictures in User Accounts

$file = C:\Users\v-sihe\Pictures\iis_logo.jpg
$userPath = "LDAP://mydc01/CN=Tobias,CN=Users,DC=powershell,DC=local"
$pic = New-Object System.Drawing.Bitmap($file)
$ms = New-Object IO.MemoryStream
$pic.Save($ms, "jpeg")
$ms.Flush()
$byte = $ms.ToArray()
$user = New-Object System.DirectoryServices.DirectoryEntry($userPath)
$user.Properties["jpegPhoto"].Value = $byte
$user.setInfo()

# Exception here.

###########################################################################################################################################################################################

# Tip 95: Finding Secret Date-Time Methods

Get-Date | Get-Member -MemberType *Method                                   # list of all methods provided by date-time objects

(Get-Date).AddYears(6)                                                      # To add six years to today's date

###########################################################################################################################################################################################

# Tip 96: Secret Timespan Shortcuts

# Note: Usually, to create time spans, you will use New-Timespan. However, you can also use a more developer-centric approach by converting a number to a time span type

[TimeSpan]100                                                               # converting a number to a time span type, This will get you a time span of 100 ticks

[TimeSpan]5d

###########################################################################################################################################################################################

# Tip 97: Finding Available .NET Frameworks

dir $env:windir\Microsoft.NET\Framework\v* -Name                            # return a string array with all installed .NET framework versions

###########################################################################################################################################################################################

# Tip 98: Use PowerShell Cmdlets

[System.DateTime]::Now

Get-Date

# Note: Whenever possible, try to avoid raw .NET access if you would like to create more readable code, Remember: Raw .NET access is the last resort if no appropriate cmdlet is available

###########################################################################################################################################################################################

# Tip 99: Retrieving Clear Text Password

$cred = Get-Credential

# to restore the clear text password entered into the dialog
$pwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $cred.Password ))
"Password: $pwd"

# Get-Credential is a great way of prompting for credentials, but the Password you enter into the dialog will be encrypted. Sometimes, you may need a clear text password

###########################################################################################################################################################################################

# Tip 100: Finding Your Current Domain

[ADSI]"" | Out-Host                        # Try this quick and simple way to find out the domain name that you are currently connected

$env:USERDOMAIN

###########################################################################################################################################################################################