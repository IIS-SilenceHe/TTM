# Reference site: http://powershell.com/cs/blogs/tips/
###########################################################################################################################################################################################

# Tip 1: Finding Maximum Values

[Int32]::MaxValue
[Int32]::MinValue

[UInt32]::MaxValue
[UInt32]::MinValue

[Int64]::MaxValue

[Byte]::MaxValue

###########################################################################################################################################################################################

# Tip 2: Check for Wildcards

$path = "C:\test\*"

[Management.Automation.WildcardPattern]::ContainsWildcardCharacters($path)       # Output: True

###########################################################################################################################################################################################

# Tip 3: Find Latest Processes

# Examples: to find all processes that were started within the past 10 minutes

# Method One:
Get-Process | Where-Object {

    try
    {
        (New-TimeSpan $_.StartTime).TotalMinutes -le 10    
    }
    catch
    {
        $false
    
    }
}



# Method Two: to filter out the $null $_.StartTimes with the where clause instead of using the try/catch to do it
Get-Process | Where-Object {$_.StartTime -and (New-TimeSpan $_.StartTime).TotalMinutes -le 10}



# Method Three: 
Get-Process | Where-Object {$_.StartTime -gt (Get-Date).AddMinutes(-10)}

###########################################################################################################################################################################################

# Tip 4: Enumerate Device Services

[System.ServiceProcess.ServiceController]::GetDevices()     # Get-Service will not  list device services. Here is how you can enumerate those low-level services

# Note: In order to run this command, in PS V2 you need to at least once call Get-Service to load the assemblies. If you do not, you get an exception complaining about an unknown type

###########################################################################################################################################################################################

# Tip 5: Read text files that are in use

# Method One: Using .Net 
$file = [System.IO.File]::Open("$env:windir\windowsupdate.log", "Open", "Read", "ReadWrite")
$reader = New-Object System.IO.StreamReader($file)
$text = $reader.ReadToEnd()
$reader.Close()
$file.Close()

# Open() will tell Windows that you want to open a file for reading and to  allow others to access the file for reading and writing while you are using it. 
 # This will allow the file to be fully accessible for others and you can now read files that are in use by other applications


# Note: [System.IO.File]::ReadAllText("c:\sometextfile.txt") would open the file for exclusive read access. 
 # No other application can access the file while you are using it, and you cannot read files that are in use by someone else.


# Method Two: Using powershell command
$myString = [string](Get-Content $env:windir\windowsupdate.log)                     # just cast Get-Content's return value to string, it's much more like powershell and less like C#

###########################################################################################################################################################################################

# Tip 6: Creating new GUIDs in various formats

$guid = [GUID]::NewGuid()

"N", "D", "B", "P" | ForEach-Object { "[GUID]::NewGuid().ToString(''{0}'') = {1}" -f $_, $guid.ToString($_)} # Format String accessable: "D", "d", "N", "n", "P", "p", "B", "b", "X" or "x"
# Output:
#        [GUID]::NewGuid().ToString(''N'') = a8a4825ce7144d1dbf0b7d8d04329622
#        [GUID]::NewGuid().ToString(''D'') = a8a4825c-e714-4d1d-bf0b-7d8d04329622
#        [GUID]::NewGuid().ToString(''B'') = {a8a4825c-e714-4d1d-bf0b-7d8d04329622}
#        [GUID]::NewGuid().ToString(''P'') = (a8a4825c-e714-4d1d-bf0b-7d8d04329622)

###########################################################################################################################################################################################

# Tip 7: Formatting multiple text lines

# to format the first dynamic information into a six-digit number and to place the second dynamic information in a column nine characters wide and format it as currency
$format = "{0:D6}  {1, 9:C}"

(4, 33.20), (12, 8.34), (2, 44.30) | ForEach-Object { "Items      Price" } { $format -f $_[0], $_[1] }
# Output:
#        Items      Price
#        000004     $33.20
#        000012      $8.34
#        000002     $44.30

# Use the awesome formatting operator -f to insert dynamic information into text! 
# You can store the formatting information in a variable and use it in a loop to  format multiple lines of text.

###########################################################################################################################################################################################

# Tip 8: Escaping text in regular expressions (RegEx) patterns

[regex]::Escape("C:\Windows")                                       # Output: C:\\Windows

# Some PowerShell operators, such as –match, expect regular expressions. If you just want to match plain text, 
 # you will need to escape any special regular expressions character in your text. Let RegEx handle it for you rather than doing that manually

# 通过将最少量的一组字符（\、*、+、?、|、{、[、(、)、^、$、.、# 和空白）替换为其转义码，将这些字符转义。 此操作指示正则表达式引擎以字面意义而非按元字符解释这些字符

###########################################################################################################################################################################################

# Tip 9: Creating random passwords

$list = [char[]]'abcdefgABCDEFG0123456&%$'

-join (1..10 | ForEach-Object { Get-Random $list -Count 1 })     # Output: 62D216G&65
-join (1..10 | ForEach-Object { Get-Random $list -Count 1 })     # Output: 3g3f%A$Gd4
-join (1..10 | ForEach-Object { Get-Random $list -Count 1 })     # Output: fa05gB%gD6

# Simply create a list of allowable characters. Next, a loop will randomly pick characters out of that list, and -join will bind them together to a string. 
 # You can adjust the list of allowable characters to draw from  and then adjust the number of iterations to control the password length.


-join (Get-Random $list -Count 10)                               # Output: 2fcd6GFa0g


-join (1..30)                                                    # Output: 123456789101112131415161718192021222324252627282930
[string](1..30)                                                  # Output: 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30

# Note: -join can keep array element together

###########################################################################################################################################################################################

# Tip 10: Appending text files without new line

# use Out-File with -Append to append lines to a text file, but there is no way to add text information without a line break. To do that, you can use .NET Framework directly.

"New Line1" | Out-File $HOME\test.txt
"New Line2" | Out-File $HOME\test.txt -Append
"New Line3" | Out-File $HOME\test.txt -Append

start $HOME\test.txt

# Output:
#        New Line1
#        New Line2
#        New Line3

[System.IO.File]::AppendAllText("$Home\test.txt", "New Word1", [System.Text.Encoding]::Unicode)
[System.IO.File]::AppendAllText("$Home\test.txt", "New Word2", [System.Text.Encoding]::Unicode)
[System.IO.File]::AppendAllText("$Home\test.txt", "New Word3", [System.Text.Encoding]::Unicode)

Get-Content $HOME\test.txt
# Output:
#        New Line1
#        New Line2
#        New Line3
#        New Word 1New Word 2New Word 3

###########################################################################################################################################################################################

# Tip 11: Use Foreach-Object instead of Select-Object -expandProperty

# Select-Object can return the content of an object property when you use the parameter -expandProperty

Get-Process | Select-Object -ExpandProperty Company            
# Select-Object -expandProperty has a number of limitations. For example, if there is no company information in a process object, you will get an error and the list is incomplete. 
# As a non-Administrator, you will experience this regularly because you cannot access processes that are owned by someone else.


Get-Process | ForEach-Object { $_.Company }                            # This approach is fault-tolerant and it also avoids some of the other bugs in Select-Object -expandProperty.

###########################################################################################################################################################################################

# Tip 12: Dump enumerations

function Get-Enum($name)      # to list all the values in an enumeration
{
    [enum]::GetValues($name) | Select-Object @{n = "Name"; e = {$_}}, @{n = "Value"; e = {$_.value__}} | format-table -autosize
}

Get-Enum IO.FileAttributes    # to easily list all the values in any enumeration

###########################################################################################################################################################################################

# Tip 13: Splitting hex dumps

# Imagine you have a text string with a hex dump so that each hex number consists of two characters, here you can split this into individual hex numbers

'00AB1CFF' -split '(?<=\G[0-9a-f]{2})(?=[0-9a-f]{2})'
# Output:
#        00
#        AB
#        1C
#        FF

###########################################################################################################################################################################################

# Tip 14: Running PowerShell on 64-bit systems

# On 64-bit systems, PowerShell will by default run in a 64-bit process. This can cause problems with snap-ins, some COM-objects (like ScriptControl)  
 # and database drivers that are designed to run in 32-bit processes. In this case, you can run them in the 32-bit PowerShell console.

$is64BitSystem = ($env:PROCESSOR_ARCHITECTURE -ne "x86")                       # $is64BitSystem returns $true if you are running on a 64-bit machine
                                                                               
$is64BitPowerShell = [IntPtr]::Size -eq 8                                      # $is64BitPowerShell returns $true if you are running in a 64-bit PowerShell session



# Making sure PowerShell scripts run in 32-bit
if ($env:Processor_Architecture -ne "x86")   
{ 
    write-warning 'Launching x86 PowerShell'

    &"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noninteractive -noprofile -file $myinvocation.Mycommand.path -executionpolicy bypass

    exit
}

"Always running in 32bit PowerShell at this point."

$env:Processor_Architecture
[IntPtr]::Size

###########################################################################################################################################################################################

# Tip 15: Getting significant bytes

# If you need to split a decimal into bytes, you can use  a function called ConvertTo-HighLow, which uses a clever combination of type casts to get you the high and low bytes
function ConvertTo-HighLow($number)
{
    $result = [System.Version][String]([System.Net.IPAddress]$number)

    $obj = 1 | Select-Object Low, High, Low64, High64

    $obj.Low = $result.Major
    $obj.High = $result.Minor
    $obj.Low64 = $result.Build
    $obj.High64 = $result.Revision

    $obj
}

ConvertTo-HighLow -number 127.0.0.1
# Output:
#        Low            High               Low64               High64
#        ---            ----               -----               ------
#        127               0                   0                    1

###########################################################################################################################################################################################

# Tip 16: Creating IP segment lists

# If you need a list of consecutive IP addresses, you can check out this function. You can see that it takes a start and an end address and then returns all IP addresses in between:
function New-IPSegment($start, $end)
{
    $ip1 = ([System.Net.IPAddress]$start).GetAddressBytes()
    [Array]::Reverse($ip1)
    $ip1 = ([System.Net.IPAddress]($ip1 -join '.')).Address

    $ip2 = ([System.Net.IPAddress]$end).GetAddressBytes()
    [Array]::Reverse($ip2)
    $ip2 = ([System.Net.IPAddress]($ip2 -join '.')).Address
    
    for($x = $ip1; $x -le $ip2; $x++)
    {
        $ip = ([System.Net.IPAddress]$x).GetAddressBytes()
        [Array]::Reverse($ip)

        $ip -join '.'
    }
}

New-IPSegment -start 1.1.1.1 -end 1.1.1.5
# Output:
#        1.1.1.1
#        1.1.1.2
#        1.1.1.3
#        1.1.1.4
#        1.1.1.5

###########################################################################################################################################################################################

# Tip 17: Checking all event logs

# What if you would like to get a quick overview of all error events in any event log. Get-EventLog can only query one event log at a time.
 #  So, you can use -list to get the names of all event logs and then loop through them.

$existingLogs = Get-EventLog -List

foreach($logType in $existingLogs)
{
    Get-EventLog $logType.Log -Newest 5 -EntryType Error -ErrorAction SilentlyContinue     # get you the latest five error events from all event logs on your system
}

###########################################################################################################################################################################################

# Tip 18: Calculate time zones

(Get-Date).ToUniversalTime()                         # Output: Thursday, September 11, 2014 07:19:12

((Get-Date).ToUniversalTime()).AddHours(8)           # Output: Thursday, September 11, 2014 15:19:12

# If you need to find out the time in another time zone, you can convert your local time to Universal Time and then add the number of offset hours to the time zone you want to display.

###########################################################################################################################################################################################

# Tip 19: Error handling for native commands

# When you need to handle errors created by native commands, you can use a wrapper function like Call. 
 # It will automatically discover when a native command writes to the error stream and return a warning

function Call
{
    $command = $args -join " "
    $command += " 2>&1"

    $result = Invoke-Expression ($command)

    $result | ForEach-Object { $e = "" } { 
    
        if($_.WriteErrorStream) 
        {
            $e += $_          
        } 
        else
        {
            $_
        }     
    }
    Write-Warning $e
}

Call net use willi
# Output:
#        WARNING: System error 67 has occurred.
#        The network name cannot be found.

###########################################################################################################################################################################################

# Tip 20: Find system restore points

# Windows Update and Software installations will frequently create system restore points. Run this to get a list of such events:
Get-EventLog -LogName Application -InstanceId 8194 | ForEach-Object {

    $i = 1 | Select-Object Event, Application

    $i.Event,$i.Application = $_.ReplacementStrings[1,0]

    $i
}


# list the command lines that actually ran after creating the restore points:
Get-EventLog -LogName Application -InstanceId 8194 | ForEach-Object {

    $i = 1 | Select-Object Event, Application

    $i.Event,$i.Application = $_.ReplacementStrings[1,0]

    $i
} | Group-Object Application | Select-Object -ExpandProperty Name 

###########################################################################################################################################################################################

# Tip 21: Analyze automatic defragmentation

# Check out this line to visualize when your system defragmented your hard drives:
Get-EventLog -LogName Application -InstanceId 258 | ForEach-Object {

    $i = 1 | Select-Object Date, Type, Drive

    $i.Date = $_.TimeWritten
    $i.Type, $i.Drive = $_.ReplacementStrings

    $i
}

# Specify InstanceID -eq 258 is not enough , U must specify also correct Source
Get-EventLog -LogName Application -InstanceId 258 -Source Microsoft-Windows-Defrag | ForEach-Object {

    $i = 1 | Select-Object Date, Type, Drive

    $i.Date = $_.TimeWritten
    $i.Type, $i.Drive = $_.ReplacementStrings

    $i
}

###########################################################################################################################################################################################

# Tip 22: Windows license validation

# Search for the appropriate events in your event log to discover when Windows has validated your license

Get-EventLog -LogName Application -InstanceId 1073745925 | Select-Object TimeWritten, Message 
# Output:
#        TimeWritten                                               Message                                                                                                   
#        -----------                                               -------                                                                                                   
#        2014/09/09 08:54:25                                       Windows license validated.                                                                                
#        2014/09/03 15:58:57                                       Windows license validated.                                                                                
#        2014/09/03 11:06:17                                       Windows license validated.                                                                                
#        2014/09/01 08:55:26                                       Windows license validated.                                                                                
#        2014/08/26 09:13:33                                       Windows license validated.                                                                                
#        2014/08/25 09:11:56                                       Windows license validated.                                                                                
#        2014/08/18 09:04:50                                       Windows license validated.                                                                                
#        2014/08/14 09:17:04                                       Windows license validated.

###########################################################################################################################################################################################

# Tip 23: Translate EventID to InstanceID

function ConvertTo-InstanceID($eventID)
{
    try
    {
        Get-WmiObject Win32_NTLogEvent -Filter "EventCode=$eventID" | ForEach-Object { $_.EventIdentifier; throw "Done" }
    }
    catch {}
}

ConvertTo-InstanceID 258

###########################################################################################################################################################################################

# Tip 24: Find multiple matches

# When you want to find matches based on regular expressions, PowerShell will only support the -match operator which finds the first match. 
 # There does not seem to be a -matches operator that returns all matches. 

"1 2  3  4  5  6" | Select-String -AllMatches -Pattern "\d" | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value


"My email is tobias.weltner@email.de and also tobias@powershell.de" | Select-String -AllMatches -Pattern '\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b' | 
    Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value

# Output:
#        tobias.weltner@email.de
#        tobias@powershell.de




# Finding multiple RegEx matches
$pattern = '\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b'
$text = "My email is tobias.weltner@email.de and also tobias@powershell.de"

filter Matches($pattern)
{
    $_ | Select-String -AllMatches $pattern | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
}

$text | Matches $pattern

###########################################################################################################################################################################################

# Tip 25: Find Dependent Services

# If you would like to check the implications of stopping  a service, you should have a look at its dependent services
(Get-Service Winmgmt).DependentServices

Get-Service Winmgmt | Select-Object -ExpandProperty DependentServices

(Get-Service Winmgmt).DependentServices | Stop-Service -WhatIf                 # To stop all dependent services, you can pipe them to Stop-Service

###########################################################################################################################################################################################

# Tip 26: Restart required

# Check out this line to determine when an installation required a system restart:
Get-EventLog -LogName Application -InstanceId 1038 | ForEach-Object {

    $i = 1 | Select-Object Date, Product, Version, Company

    $i.Date = $_.TimeWritten
    $i.Product, $i.Version, $i.Company = $_.ReplacementStrings[0,1,5]

    $i
}

# It will return an exception if no such events are found. Otherwise, you will get a listing with the dates and the products that required a system restart

###########################################################################################################################################################################################

# Tip 27: RegEx Magic

# The [RegEx] type has a method called Replace(), which can be used to replace text by using regular expressions. 

 # This line would replace the last octet of an IP address with a new fixed value, 
 # In this example, the RegEx pattern represents a 1-3 digit number at the end of a text ("$"). It is replaced with a new value of "201."

[regex]::Replace("192.168.1.200", "\d{1,3}$", "205")                           # Output: 192.168.1.205


# If you want to replace the pattern with a calculated value, you can also submit a script block, which accepts the original value in $args[0].
 # The next line would actually increment the last octet of an IP address by 1

[regex]::Replace("192.168.1.1", "\d{1,3}$", {[int]$args[0].value + 1})         # Output: 192.168.1.2

###########################################################################################################################################################################################

# Tip 28: Copy Registry Hives

# You can use Copy-Item to quickly copy entire structures of registry keys and sub-keys in a matter of milliseconds. 
 # Take a look at this example – it creates a new registry key in HKCU and then copies a key with all of its values and sub-keys from HKLM to the new key:

md Registry::HKCU\Software\Testkey
Copy-Item -Path Registry::HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall -Destination Registry::HKCU\Software\Testkey -Recurse

# Note the use of -Recurse to get a complete copy of all sub-keys and values.

###########################################################################################################################################################################################

# Tip 29: Use WMI to Create Hardware Inventory

Get-WmiObject -Class CIM_LogicalDevice | Select-Object -Property Caption, Description, __Class            # use a generic device class that applies to all hardware devices


# The __Class column lists the specific device class. If you'd like to find out more details about  a specific device class, you can then pick the name and query again:
Get-WmiObject -Class Win32_PnPEntity | Select-Object Name, Service, ClassGUID | Sort-Object ClassGUID

###########################################################################################################################################################################################

# Tip 30: C:\Users

function Get-FolderSize($path)     # Here is a useful function that you can use to calculate the total size of a folder
{
    Get-ChildItem $home -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum | Select-Object -ExpandProperty Sum 
}

Get-FolderSize $HOME

# Note: It may take some time for the function to complete because it recursively visits all folders and sums up all file sizes.

###########################################################################################################################################################################################

# Tip 31: Parsing Text-Based Log Files


# Extracting useful information from huge text-based log files 
Get-Content $env:windir\windowsUpdate.log -encoding utf8 | Where-Object {$_ -like "*successfully installed*"} | ForEach-Object {($_ -split ": ")[-1]} # product name of installed updates


# the pipeline has considerable overhead. Without the pipeline approach, you can get the same results about 10-times faster:
$lines = Get-Content $env:windir\windowsUpdate.log -ReadCount 0 -encoding utf8
foreach($line in $lines)
{
    if($line -like "*successfully installed*")
    {
        ($line -split ": ")[-1]
    }
}

###########################################################################################################################################################################################

# Tip 32: Grouping Using Custom Criteria

Get-Process | Group-Object -Property Company                         # Try using Group-Object to group objects by any property

$criteria = {

    if($_.Length -lt 1kb)
    {
        "Tiny"
    }
    elseif($_.Length -lt 1mb)
    {
        "Average"
    }
    else
    {
        "Huge"
    }
}

dir $env:windir | Group-Object -Property $criteria                    # Grouping Using Custom Criteria: to group by file size into three categories



# Group-Object can also auto-create hash tables so that you can easily create groups of objects of a kind
$myFiles = dir $env:windir | Group-Object -Property $criteria -AsHashTable -AsString

$myFiles.Huge
$myFiles.Average
$myFiles.Tiny

###########################################################################################################################################################################################

# Tip 33: Creating Byte Arrays

$byte = New-Object byte[] 100                      # to create a new empty byte array with 100 bytes
                                                   
$byte = [byte[]](, 0xFF * 100)                     # to create a byte array with a default value other than 0
                                                   
                                                   
                                                   
$array = New-Object object[,] n,m                  # create array with n * m elements

###########################################################################################################################################################################################

# Tip 34: Case-Sensitive Hash Tables

$hash = @{}
$hash.key = 1
$hash.kEy = 2
$hash.KEY                                          # Output: 2, PowerShell hash tables are, by default, not case sensitive


# If you need case-sensitive keys, you can create the hash table this way:
$hash = New-Object System.Collections.Hashtable
$hash.Key = 1
$hash.kEy = 2

$hash.KEY                                          # return nothing here because its case-sensitive now

$hash
# Output:
#        Name                           Value                                                                                                                                                                                 
#        ----                           -----                                                                                                                                                                                 
#        kEy                            2                                                                                                                                                                                     
#        Key                            1                                               

###########################################################################################################################################################################################

# Tip 35: Tile Windows

$shell = New-Object -ComObject Shell.Application

$shell.TileHorizontally()                          # to tile all open windows horizontally
$shell.TileVertically()                            # to tile all open windows vertically
$shell.CascadeWindows()                            # to cascade all open windows

###########################################################################################################################################################################################

# Tip 36: Generate Random Passwords

-join ([char[]]'abcdefgABCDEFG0123456&%$' | Get-Random -Count 20)

-join ([char[]]@(33..126) | Get-Random -Count 20)

-join ([char[]]@(33..126) | Get-Random -Count $(Get-Random -Minimum 7 -Maximum 15))

###########################################################################################################################################################################################

# Tip 37: Re-Assigning Types to Variables

[int]$a = 1
$a = "hello"                                       # exception here because you can no longer assign other types when you strongly type a variable

[String]$a = "hello"                               # However, you can always re-assign a new type to a variable

###########################################################################################################################################################################################

# Tip 38: Check Array Content With Wildcards

$name = dir $env:windir | Select-Object -ExpandProperty Name

$name -contains "explorer.exe"                                    # Output: True
$name -contains "explorer*"                                       # Output: False, -contains does not support wildcards

$name -like "explorer*"                                           # Output: explorer.exe, use -like instead -contains, then it works well

###########################################################################################################################################################################################

# Tip 39: Check For Numeric Characters

# to check a single character and find out whether or not it is numeric
[char]::IsNumber("1")
[char]::IsNumber("A")


[char] | Get-Member -Static                                       # The type Char has a bunch of other useful methods

###########################################################################################################################################################################################

# Tip 40: Switch Accepts Arrays

switch(1,5,2,4,3,1)
{
    1 {"One"}
    2 {"Two"}
    3 {"Three"}
    4 {"Four"}
    5 {"Five"}
}

# Output: 
#        One
#        Five
#        Two
#        Four
#        Three
#        One





# Multiple Text Replace

$text = 'Österreich überholt außen Ängland'

# Method One:
-join $(switch -CaseSensitive ([char[]]$text){

    ä { 'ae' }
    ö { 'oe' }
    ü { 'ue' }
    Ä { 'Ae' }
    Ö { 'Oe' }
    Ü { 'Ue' }
    ß { 'ss' }
    default {$_}
})

# Output: Oesterreich ueberholt aussen Aengland


# Method Two:
$text = $text.replace("Ä", "Ae")
$text = $text.replace("Ö", "Oe")
$text = $text.replace("Ü", "Ue")
$text = $text.replace("ä", "ae")
$text = $text.replace("ö", "oe")
$text = $text.replace("ü", "ue")
$text = $text.replace("ß", "ss")
$text

# Or:
$text.replace("Ä", "Ae").replace("Ö", "Oe").replace("Ü", "Ue").replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")

###########################################################################################################################################################################################

# Tip 41: Test for STA mode

function Test-STA
{
    $Host.Runspace.ApartmentState -eq "STA"
}

# By default, the PowerShell console does not use the STA mode whereas the ISE editor does. 
 # STA is needed to run Windows Presentation Foundation scripts and to use WPF-based dialog windows, such as Open.


# A more reliable way is this:
[Threading.Thread]::CurrentThread.ApartmentState -eq "STA"

###########################################################################################################################################################################################

# Tip 42: About Clipboard

# Get-Clipboard
function Get-Clipboard
{
    if($Host.Runspace.ApartmentState -eq "STA")          # If your PowerShell host uses the STA mode, you can easily read clipboard content 
    {
        Add-Type -AssemblyName PresentationCore
        [Windows.Clipboard]::GetText()
    }
    else
    {
        Write-Warning ('Run {0} with the -STA parameter to use this function' -f $Host.Name)
    }
}


# Set-Clipboard

dir $env:windir | clip                                   # you can pipe text to clip.exe  to copy it to your clipboard

function Set-Clipboard($text)
{
    if($Host.Runspace.ApartmentState -eq "STA")
    {
        Add-Type -AssemblyName PresentationCore
        [Windows.Clipboard]::SetText($text)
    }
    else
    {
        Write-Warning ('Run {0} with the -STA parameter to use this function' -f $Host.Name)
    }
}

Set-Clipboard "Silence, you're good!"

###########################################################################################################################################################################################

# Tip 43: Speed Up Loops

$array = 1..10000

Measure-Command {

    for($i = 0; $i -lt $array.count; $i++)
    {
        $array[$i]
    }
}

# Ouput:
#       Days              : 0
#       Hours             : 0
#       Minutes           : 0
#       Seconds           : 0
#       Milliseconds      : 205
#       Ticks             : 2054654
#       TotalDays         : 2.37807175925926E-06
#       TotalHours        : 5.70737222222222E-05
#       TotalMinutes      : 0.00342442333333333
#       TotalSeconds      : 0.2054654
#       TotalMilliseconds : 205.4654


# this one much faster than above
Measure-Command {

    $length = $array.count

    for($i = 0; $i -lt $length; $i++)
    {
        $array[$i]
    }
}

# Output:
#        Days              : 0
#        Hours             : 0
#        Minutes           : 0
#        Seconds           : 0
#        Milliseconds      : 62
#        Ticks             : 629810
#        TotalDays         : 7.28946759259259E-07
#        TotalHours        : 1.74947222222222E-05
#        TotalMinutes      : 0.00104968333333333
#        TotalSeconds      : 0.062981
#        TotalMilliseconds : 62.981

###########################################################################################################################################################################################

# Tip 44: Duplicate Output

$result = Get-Process                               # Once you assign results to a variable, however, the console will no longer show the results:

($result = Get-Process)                             # You should place the line in parenthesis to store the results in a variable and also output the results to the console

###########################################################################################################################################################################################

# Tip 45: List NTFS Permissions

# To view NTFS permissions for folders or files, use Get-Acl. It won't show you the actual permissions at first, but you can make them visible like this:
Get-Acl -Path $env:windir | Select-Object -ExpandProperty Access

###########################################################################################################################################################################################

# Tip 46: Echoing The Error Channel

"
@echo off
echo Starting ...
netstat
echo done!
"

# When you run it from PowerShell and save its results, all information is stored in $result, including your messages to the user:
$result = .\test.bat


# This version makes sure the user messages remain visible and do not get redirected to the variable:
"
@echo off
echo Starting... 1>&2
netstat
echo Done! 1>&2
"

###########################################################################################################################################################################################

# Tip 47： Detecting Remote Visitors

# Note: Whenever someone connects to your computer using PowerShell remoting, there is a host process called wsmprovhost.exe. 
Get-Process wsmprovhost -ErrorAction SilentlyContinue


@(Get-Process wsmprovhost -ErrorAction SilentlyContinue).Count -gt 0          # To check whether there is (at least one) active remote PowerShell session on your computer


# List remote connection owner
function Get-PSRemotingVisitor($computerName = "localhost")
{
    Get-WmiObject Win32_Process -Filter 'Name="wsmprovhost.exe"' -ComputerName $computerName | ForEach-Object {
    
        $o = $_.GetOwner()
        $o = $o.Domain + "\" + $o.User

        $obj = $_ | Select-Object Name, CreationDate, Owner
        $obj.Owner = $o
        $obj.CreationDate = $_.ConvertToDateTime($_.CreationDate)

        $obj
    }
}

###########################################################################################################################################################################################

# Tip 48: Analyze Cmdlet Results

# There are two great ways to analyze the results a cmdlet returns: you can send the results to Get-Member to get a formal analysis, telling you the properties, 
 # methods and data types, and you can send them to Select-Object to view the actual property content to get a feeling where to look for specific information. 

Get-Process | Get-Member

Get-Process | Select-Object -Property * -First 1                  # Note the use of -First 1: this gets you only the first result which most of the time is sufficient for analysis.


###########################################################################################################################################################################################

# Tip 49: Finding Type Definitions

Select-XML $env:windir\System32\WindowsPowerShell\v1.0\types.ps1xml -Xpath /Types/Type/Name | Select-Object -expand Node | Sort-Object # Sort does not work here, see below right one:
 
Select-XML $env:windir\System32\WindowsPowerShell\v1.0\types.ps1xml -Xpath /Types/Type/Name | Select-Object -expand Node | Sort-Object "#text"

# Note: PowerShell enhances many .NET types and adds additional information. These changes are defined in xml files. 



# Finding Useful WMI Classes

# To find the most useful WMI classes you can use Get-WmiObject, and let PowerShell provide you with a hand-picked list:
Select-XML $env:windir\System32\WindowsPowerShell\v1.0\types.ps1xml -Xpath /Types/Type/Name |
ForEach-Object { $_.Node.innerXML } | Where-Object { $_ -like '*#root*' } |
ForEach-Object { $_.Split('\')[-1] } | Sort-Object


# These are the WMI classes found especially helpful by the PowerShell developers. Just pick one and submit it to Get-WmiObject. Here is an example:
Get-WmiObject Win32_BIOS

###########################################################################################################################################################################################

# Tip 50: Removing File Extensions 

[System.IO.Path]::GetFileNameWithoutExtension("C:\test\report.txt")                # Output: report

###########################################################################################################################################################################################

# Tip 51: Finding Open Files

openfiles                                         # to see which files are opened by network users on your machine

openfiles /id 1234 /disconnect                    # To force a file to close, use /Disconnect along with the connection id


# Openfiles can also track open files on your local machine. You pay for it with a lower overall system performance because the system now tracks all open files. 

openfiles /local ON                               # To enable tracking (or lists) local files, machine restart is need

openfiles /local OFF                              # To return to system default and stop monitoring local files, machine restart is need

###########################################################################################################################################################################################

# Tip 52: Escape Regular Expressions

'c:\test\subfolder\file' -split '\'               # Exception here: Split expects a regular expression and fails when you use special characters like "\"

[regex]::Escape("\")                              # Output: \\, To translate a plain text into an escaped regular expression text

'c:\test\subfolder\file' -split '\\'              # It's OK here after "\\" instead of "\"

###########################################################################################################################################################################################

# Tip 53: Find Local Users

function Get-LocalUser
{
    $users = net user
    $users[4..($users.count - 3)] -split "\s+" | Where-Object { $_ }
}

Get-LocalUser
# Output: 
#        Administrator
#        Guest
#        Silence

(Get-LocalUser) -contains "Silence"                                      # Output: True

(Get-LocalUser) -like "*silence*"                                        # Output: Silence

###########################################################################################################################################################################################

# Tip 54: Find Local Groups

function Get-LocalGroup
{
    net localgroup | Where-Object { $_.StartsWith("*") } | ForEach-Object { $_.SubString(1) }
}


# Find Local Group Members
function Get-LocalGroupMember
{
    param([Parameter(Mandatory = $true)]$name)

    try
    {
        $ErrorActionPreference = "Stop"
        $users = net localgroup $name 2>&1                                #  the function receives native error messages by redirecting the error to the input channel (2>&1)
        $users[6..($users.count - 3)] -split "\s+" | Where-Object { $_ }
    }
    catch
    {
        $errmsg = $_

        if($errmsg -match '\b(\d{1,8})\b')                                # uses a regular expression to check for an error number
        {
            $errmsg = net helpmsg ($Matches[1])
        }

        Write-Warning "Get-LocalGroupMember : $errmsg"
    }
}

Get-LocalGroupMember admin                                                 # Output: WARNING: Get-LocalGroupMember :  The specified local group does not exist.
Get-LocalGroupMember Guests                                                # Output: Guest

###########################################################################################################################################################################################

# Tip 55: Shortcut to Network Cards

explorer.exe '::{7007ACC7-3202-11D1-AAD2-00805FC1270E}'                    # To quickly access the settings for your network cards

ncpa.cpl                                                                   # It's OK to use cmd command



# Enumerating Network Cards programmatically 
Get-WmiObject Win32_NetworkAdapter -Filter "NetConnectionID!=NULL" | Select-Object -ExpandProperty NetConnectionID

###########################################################################################################################################################################################

# Tip 56: Getting Network Adapter Settings

$adapterid = "Local Area Connection"
$nic = Get-WmiObject Win32_NetworkAdapter -Filter "NetConnectionID='$adapterid'"

$nicconfig = $nic.getRelated('Win32_NetworkAdapterConfiguration')
$nic
$nicconfig

# About getRelated method details: http://richardspowershellblog.wordpress.com/2011/05/01/powershell-deep-dive-v-wmi-associations/

###########################################################################################################################################################################################

# Tip 57: Adding Personal Drives

# WMI network adapter information is separated into two classes. Win32_NetworkAdapter represents the hardware, and Win32_NetworkAdapterConfiguration contains the configuration details.


# To mix information from both classes Win32_NetworkAdapter and Win32_NetworkAdapterConfiguration
Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled=true" | ForEach-Object {

    $info = $_ | Select-Object *

    $nic = $_.GetRelated("Win32_NetworkAdapter")

    $nic | Get-Member -MemberType *property | ForEach-Object{ 
        
        if($_.Name.StartsWith("__") -eq $false)
        {
            Add-Member -InputObject $info NoteProperty $_.Name $nic.$($_.Name) -ErrorAction SilentlyContinue
        }
    }

    $info
}

# The code gets all network adapters with an IP address assignment, then use GetRelated() to find the related hardware information from Win32_NetworkAdapter. 
 # At the end, the hardware information is merged into the configuration information so you get a new rich object which combines all network adapter information.

###########################################################################################################################################################################################

# Tip 58: When to use Select-Object's -ExpandProperty

Get-Process | Select-Object -Property *                               # Use -Property * when you want to see maximum information
Get-Process | Select-Object -Property Name, Description, Company      # Use -Property a,b,c to select more than one column

Get-Process | Select-Object -ExpandProperty Name                      # Use -ExpandProperty Column to select exactly one column
Get-Process | Select-Object -Property Name
# Note: It does not make sense to preserve the column when you select only one column anyway, so use -ExpandProperty to remove the column and show only the actual column content.

###########################################################################################################################################################################################

# Tip 59: Adding Personal Drives

function Add-PersonalDrive
{
    [System.Enum]::GetNames([System.Environment+SpecialFolder]) | ForEach-Object {
    
        $name = $_
        $target = [System.Environment]::GetFolderPath($_)

        New-PSDrive -Name $name -PSProvider FileSystem -Root $target
    }
}

Add-PersonalDrive
Get-PSDrive                                 # when you call that code from within a function, your new drives are gone right after you created them

# Soultion: Add global scope
function Add-PersonalDrive
{
    [System.Enum]::GetNames([System.Environment+SpecialFolder]) | ForEach-Object {
    
        $name = $_
        $target = [System.Environment]::GetFolderPath($_)

        New-PSDrive -Name $name -PSProvider FileSystem -Root $target -Scope Global   # Add Scope here then you can get the new added PSDrive with Get-PSDrive
    }
}

Add-PersonalDrive
Get-PSDrive

###########################################################################################################################################################################################

# Tip 60: Eliminating Empty Text

get-process | Where-Object {$_.Company -ne $null} | Select-Object Name, Company, Description                                         # filter based on $null values

Get-Process | Where-Object {$_.Company -ne $null} | Where-Object {$_.Company -ne ""} | Select-Object Name, Company, Description      # filter based on $null and empty values

Get-Process | Where-Object { -not [string]::IsNullOrEmpty($_.Company)} | Select-Object Name, Company, Description                    # # filter based on $null and empty values

###########################################################################################################################################################################################

# Tip 61: Assigning Two Unique Random Numbers

# If you need to get two random numbers from a given numeric range, and you want to make sure they cannot be the same, 
 # simply tell Get-Random to pick two numbers, and assign them to two different variables at the same time:
$foreground, $background = Get-Random -InputObject (0..15) -Count 2

# In our example, you would draw a random foreground and background color while ensuring that foreground and background color can never be the same:
Write-Host -ForegroundColor $foreground -BackgroundColor $background "This is random coloring."

###########################################################################################################################################################################################

# Tip 62: Create CSV without Header

$filepath = "$env:windir\rawcsv.txt"

$process = Get-Process | ConvertTo-Csv 

$process[2..$($process.count - 1)] | Out-File $filepath

notepad $filepath


# A slightly simpler option:
Get-Process | ConvertTo-Csv -UseCulture | Select-Object -Skip 2 | Out-File $filepath

###########################################################################################################################################################################################

# Tip 63: Appending CSV Data

# To append a CSV file with new data, first of all make sure the type of data you append is the same type of data already in a file (or else column names will not match).
$filePath = "$home\processes.csv"

Get-Process | Select-Object Name, Company, Description -Unique | Export-Csv $filepath -UseCulture -NoTypeInformation -Encoding UTF8       # list of unique running processes

# Let's assume you'd like to append new processes to that list, so whenever you run the following code, 
 # you want it to check whether there are new processes, and if so, add them to the CSVfile:
$oldproc = Import-Csv $filepath -UseCulture | Select-Object -ExpandProperty Name

$newproc = Get-Process | Where-Object {$oldproc -notcontains $_.Name} | Select-Object Name, Company, Description -Unique | ForEach-Object {

    Write-Host "Found new process: $($_.Name)"
    $_

} | ConvertTo-Csv -UseCulture                                                               # save new processes as CSV. Make sure to use the same delimiter

$newproc[2..$($newproc.count - 1)] | Out-File -Append $filepath -Encoding utf8              # add new CSV to the old file. Make sure to use same encoding

###########################################################################################################################################################################################

# Tip 64: Managing File System Tasks

Get-Command -Noun item*, path                                                               # to list all cmdlets that deal with file system-related tasks

Get-Alias -Definition *-item*, *-path* | Select-Object Name, Definition                     # Many of these cmdlets have historic aliases that will help you guess what they are doing

###########################################################################################################################################################################################

# Tip 65: Launching Applications

Start-Process -FilePath notepad -ArgumentList "$env:windir\system32\drivers\etc\hosts"

# Note: When you launch *.exe-applications with arguments, you may get exceptions because PowerShell may misinterpret the arguments. 
 # A better way to do this is using Start-Process and then separate file path and arguments with the parameters -FilePath and -ArgumentList. 
 # This way, you can safely enclose the arguments in quotes, too

###########################################################################################################################################################################################

# Tip 66: Use Select-String with Context

ipconfig | Select-String DNS -Context 0

###########################################################################################################################################################################################

# Tip 67: Use RegEx to Extract Text

$text = 'The problem was discussed in KB552356. Mail feedback to tobias silence@powershell.com'
$pattern = "KB\d{4,6}"

if($text -match $pattern)
{
    $Matches[0]                                                      # Output: KB552356
}


# Extract Text without Regular Expressions
$words = $text.Split(" ")

$words -like "KB*"                                                   # Output: KB552356.

$words -like "*@*.*"                                                 # Output: silence@powershell.com

# Note: By splitting the text into words, you can then use -like and simple patterns that use the "*" wildcard

###########################################################################################################################################################################################

# Tip 68: Locking the Computer

rundll32.exe user32.dll, LockWorkStation

# To lock your computer from PowerShell, remember that you can launch applications, including rundll32.exe, which can be used to call methods from inside DLL files

###########################################################################################################################################################################################

# Tip 69: Simple Breakpoint

$Host.EnterNestedPrompt()                # If you want PowerShell to halt your script at some point, you can simply add this line

# This will suspend execution and you will get back to the prompt. You can now examine your script variables or even change them. To resume the script, you can then type exit.

###########################################################################################################################################################################################

# Tip 70: Open File Exclusively

$path = "$home\test.txt"

$file = [System.IO.File]::Open($path, "Open", "Read", "None")    # To open a file in a locked state so no one else can open, access, read, or write the file
$reader = New-Object System.IO.StreamReader($file)               
$text = $reader.ReadToEnd()                                      
$reader.Close()                                                  
                                                                 
Read-Host "Press Enter to release file!"                         # This will lock a file and read its content. To illustrate the lock, the file will remain locked until you press ENTER
$file.Close()

###########################################################################################################################################################################################

# Tip 71: Creating Self-Updatable Variables

# If you want a variable to update its content every time you retrieve it, you can assign a breakpoint and an action to it. 
 # The action script block will then be executed each time the variable is read

$Global:Now = Set-PSBreakpoint -Variable Now -Mode Read -Action {Set-Variable -Name Now -Value (Get-Date) -Option ReadOnly, AllScope -Scope global -Force}

# So now, each time you output $now, it will return the current date and time.

###########################################################################################################################################################################################

# Tip 72: Save result with Excel file

# To send PowerShell results to Excel, you can use CSV-files and automate the entire process by using this function:

function Out-Excel($path="$home\silencetest\$(Get-Date -format 'yyyyMMddHHmmss').csv") 
{
    $input | Export-Csv $path -UseCulture -Encoding UTF8 -NoTypeInformation
    
    Invoke-Item $path 
}

Get-Process | Out-Excel
Get-WmiObject Win32_BIOS | Out-Excel

###########################################################################################################################################################################################

# Tip 73: Finding Files

dir $env:windir\System32\ -Filter hosts -Recurse -ErrorAction SilentlyContinue -Name                                   # Output: drivers\etc\hosts

# Note: Dir (Get-Childitem) is a simple, but effective way to find files. The following line will find any file or folder called "hosts" anywhere inside your Windows folder. 
 # It may take some time for the results to display because this command will initiate a full recursive search

###########################################################################################################################################################################################

# Tip 74: Check Whether a Program is Running

(Get-Process winword -ErrorAction SilentlyContinue) -ne $null

# This line will returns $true if at least one instance is running. Note that this will also report running instances in other sessions. 
 # So, if multiple users are logged on or you are using a terminal server, you may want to check for process ownership.

# Disscuss:
(get-process notepad -ea 0) -ne $null                              # Result is a list but not a boolean value

[bool](get-process notepad -ea 0) -ne $null                        # The fact that PowerShell evalutes (non empty) collections to a boolean true should work

###########################################################################################################################################################################################

# Tip 75: Solving Problems with Parenthesis 

try { 1 } catch { 2 }

(try { 1 } catch { 2 })                                            # Exception here: some language keywords are not legal inside parenthesis, use $() instead can work around this

$(try { 1 } catch { 2 })                                           # With $(), you can group commands without the language restrictions that () imposes

###########################################################################################################################################################################################

# Tip 76: Re-Encoding ISE-Scripts in UTF8

dir *.ps1 | ForEach-Object { (Get-Content -Path $_.FullName) | Set-Content -Path $_.FullName -Encoding UTF8 }  


###########################################################################################################################################################################################

# Tip 77：Validating Function Parameters

function Get-ZIPCode
{
    param(
          [ValidatePattern('^\d{5}$')]                         # use Regular Expression patterns to validate function parameters
          [String]
          $ZIP
    )

    "Here is the ZIP code you entered: $ZIP"
}

Get-ZIPCode "abcde"                                            # Exception here: The argument "abcde" does not match the "^\d{5}$" pattern
Get-ZIPCode "12345"                                            # Output: Here is the ZIP code you entered: 12345

# Note: You can add a [ValidatePattern()] attribute to the parameter that you want to validate, and specify the RegEx pattern that describes valid arguments. 

###########################################################################################################################################################################################

# Tip 78: Create Shortcut on Your DeskTop for PowerShell

$shell = New-Object -ComObject WScript.Shell

$lnk = $shell.CreateShortcut("$([System.Environment]::GetFolderPath('Desktop'))\MyPS.lnk")
$lnk.TargetPath = (Get-Process -id $pid).Path
$lnk.Save()

# Note that the shortcut created by this code will launch the very PowerShell host you ran to create the shortcut. 
 # So when you run it in PowerShell ISE, you will get a link to ISE, and when you run it in powershell.exe, the link will point to the PowerShell console.

###########################################################################################################################################################################################

# Tip 79: Return Arrays

function Test                                          # Normally, PowerShell will not preserve the type of an array returned by a function. It is always reduced to Object[]
{
    $al = [System.Collections.ArrayList](1..10)
    $al
}

(Test).GetType().FullName                              # Output: System.Object[]



function Test
{
    $al = [System.Collections.ArrayList](1..10)
    ,$al                                               # By adding a comma telling PowerShell that it should return the entire array
}

(Test).GetType().FullName                              # Output: System.Collections.ArrayList


# Returning Array in One Chunk

function test { (1..5) } 
Test | ForEach-Object { "Receiving $_" }
# Output: 
#        Receiving 1
#        Receiving 2
#        Receiving 3
#        Receiving 4
#        Receiving 5

# Arrays returned by functions are by default unwrapped and processed as single values. With a simple comma, this will change.

function test { ,(1..5) } 
Test | ForEach-Object { "Receiving $_" }
# Output:
#        Receiving 1 2 3 4 5

###########################################################################################################################################################################################

# Tip 80: Renaming Computers

# PowerShell can also rename computers. 
 # The next example will read the serial number from the system enclosure class and rename the computer accordingly (provided you have local admin privileges):

$serial = Get-WmiObject Win32_SystemEnclosure | Select-Object -ExpandProperty serialnumber

if($serial -ne "None")
{
    (Get-WmiObject Win32_ComputerSystem).rename("Desktop_$serial")
}
else
{
    Write-Warning "Computer has no serial number"
}

###########################################################################################################################################################################################

# Tip 81: Monitoring Folder Content

$log = "$home\silencetest\newfiles.txt"

$folder = "$HOME\silencetest"
$timeout = 1000

$fileSystemWatcher = New-Object System.IO.FileSystemWatcher $folder

Write-Host "Press any key to abort monitoring $folder."
do
{
    $created = $fileSystemWatcher.WaitForChanged("Created", $timeout)

    if($created.TimedOut -eq $false)
    {
        "{0} : {1}" -f (Get-Date), ($created.Name + " created") | Out-File $log -Append

        Write-Warning ("Detected new file: " + $created.Name)
    }

}until([System.Console]::KeyAvailable)

Write-Host "Monitoring aborted."
Invoke-Item $log

###########################################################################################################################################################################################

# Tip 82: Bulk Renaming Files

$global:i = 1
dir "$home\silencetest" -Filter cover*.jpg | Rename-Item -NewName { "picture_$i.jpg"; $global:i++} -WhatIf

# This line will take all *.jpg files in c:\test1 and rename them to "picture_x.jpg" where "x" is an incrementing number. 
 # The secret is to submit a script block to Rename-Item, which will then dynamically calculate the name that will need to be assigned to a particular file.

###########################################################################################################################################################################################

# Tip 83: Adding Members to Local Group

# To manage local groups, you can think about using net.exe. It may be much easier than using COM interfaces.

net localgroup Administrators Silence /add                                    # Adding Members to Local Group

# Adding Members to Local Group on non-en OS
$lag=((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')).Translate([System.Security.Principal.NTAccount]).Value.Split('\')[1])
net localgroup $lag redmond\fwtlaba /add

###########################################################################################################################################################################################

# Tip 84: Creating Local Admins

# Method One:

$computername = $env:computername   # place computername here for remote access
$username = 'AdminAccount1'
$password = 'topSecret@99'
$desc = 'Automatically created local admin account'

# create a local user account and put it into the local Administrators group
$computer = [ADSI]"WinNT://$computername,computer"
$user = $computer.Create("user", $username)
$user.SetPassword($password)
$user.Setinfo()
$user.description = $desc
$user.setinfo()
$user.UserFlags = 65536
$user.SetInfo()
$group = [ADSI]("WinNT://$computername/administrators,group")
$group.add("WinNT://$username,user")

# You do need local Admin privileges to execute this script. 
 # Adjust the name of the "Administrators" group to match your locale. For example, on German systems, it is called "Administratoren".




# Method Two:
# Create user
net user $username $userpassword /add
net localgroup Administrators Silence /add

###########################################################################################################################################################################################

# Tip 85: How Long Has Shell Been Running?

((((Get-Date)-(Get-Process -id $pid).starttime) -as [string]) -split '\.')[0]      # To find out how long your PowerShell session has been running

###########################################################################################################################################################################################

# Tip 86: Forwarding Parameters

# Get-BIOS works both locally and remotely. Only the parameters submitted by the user will actually be forwarded to Get-WmiObject.
function Get-BIOS($computername, $credential)
{
    Get-WmiObject Win32_BIOS @psboundparameters
}

# To forward function parameters to a cmdlet, use $psboundparameters automatic variable and splatting

###########################################################################################################################################################################################

# Tip 87: Enabling Remote WMI and DCOM

# Many cmdlets have a built-in -ComputerName parameter that will allow for remote access without using the new PowerShell remoting. 
 # For this to work, your firewall will need to be adjusted on the target machine:

# In addition, some cmdlets (like Get-Service) will need the Remote Registry service to be running on the target side:
netsh firewall set service type = remoteadmin mode = enable

Start-Service RemoteRegistry
Set-Service RemoteRegistry -StartupType Automatic

###########################################################################################################################################################################################

# Tip 88: Filtering Files or Folders

# To filter folder content by file or folder, check whether the Length property is present. It is present for files and missing in folders

dir $env:windir | Where-Object {$_.PSIsContainer}              # get folders
dir $env:windir | Where-Object {-not $_.PSIsContainer}         # get files

###########################################################################################################################################################################################

# Tip 89: Forwarding Selected Parameters

function Get-BIOS($computerName, $credential, [switch]$Verbose)
{
    $a = $Global:psboundparameters
    $psboundparameters.Remove("Verbose") | Out-Null
    $bios = Get-WmiObject Win32_BIOS @psboundparameters

    if($Verbose)
    {
        $bios | Select-Object *
    }
    else
    {
        $bios
    }
}

# The function supports three parameters, but only two should be forwarded to Get-WmiObject. The remaining parameter -Verbose is used internally by the function. 
 # To prevent -Verbose from being forwarded to Get-WmiObject, you can remove that key from $psboundparameters before you splat it to Get-WmiObject.

Get-BIOS -Verbose
Get-BIOS -ComputerName IIS-CTI5052 -Verbose

###########################################################################################################################################################################################

# Tip 90: Extracting Icons

# To extract an icon from a file, use .NET Framework methods. 
 # Here is a sample that extracts all icons from all exe files in your Windows folder (or one of its subfolders) and puts them into a separate folder.

[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null

$folder = "$home\silencetest\icons"
md $folder -ErrorAction SilentlyContinue | Out-Null

dir $env:windir -Filter *.exe -ErrorAction SilentlyContinue | ForEach-Object {

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)

    Write-Progress "Extracting Icon" $baseName

    [System.Drawing.Icon]::ExtractAssociatedIcon($_.FullName).ToBitmap().Save("$folder\$baseName.ico")
}

###########################################################################################################################################################################################

# Tip 91: Displaying Balloon Tip

# to share status information via a balloon message in the system tray area

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$balloon = New-Object System.Windows.Forms.NotifyIcon

$path = Get-Process -id $pid | Select-Object -ExpandProperty Path
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)

$balloon.Icon = $icon
$balloon.BalloonTipIcon = "Info"                                           # Your options for the icon include None, Info, Warning, and Error
$balloon.BalloonTipText = "Completed Operation"
$balloon.BalloonTipTitle = "Done"
$balloon.Visible = $true
$balloon.ShowBalloonTip(10000)

# Note that the code uses the icon of your PowerShell application inside the tray area so the user can associate the message with the application that produced it.

###########################################################################################################################################################################################

# Tip 92: Determine Functions Pipeline Position

# Assume your function wanted to know whether it is the last element in a pipeline or operating in the middle of it. 
function Test([Parameter(ValueFromPipeline=$true)]$data)
{
    process
    {
        if($myinvocation.PipelinePosition -ne $myinvocation.PipelineLength)
        {
            $data
        }
        else
        {
            Write-Host $data -ForegroundColor Red -BackgroundColor White
        }
    }
}

1..3 | test
1..3 | test | Get-Random -Count 3

# "test" will output the data in red color when it is the last element (thus controlling presentation itself) whereas it forwards the data if it is not the last pipeline command.

###########################################################################################################################################################################################

# Tip 93: Listing Windows Updates

# There is a not widely known COM object that you can use to list all the installed Windows Updates on a local machine
$session = New-Object -ComObject Microsoft.Update.Session
$searcher = $session.CreateUpdateSearcher()
$historyCount = $searcher.GetTotalHistoryCount()
$searcher.QueryHistory(1, $historyCount) | Select-Object Date, Title, Description


# Checking Windows Updates Remotely
function Get-SoftwareUpdate($computername, $credential)
{
    $code = {
    
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $historyCount = $searcher.GetTotalHistoryCount()
        $searcher.QueryHistory(1, $historyCount) | Select-Object Date, Title, Description
    }

    $pcname = @{Name = "Machine"; Expression = {$_.PSComputerName}}

    Invoke-Command $code @psboundparameters | Select-Object $pcname, Date, Title, Description
}

Get-SoftwareUpdate -computername IIS-CTI5052

###########################################################################################################################################################################################

# Tip 94: Controlling PSComputerName in Remoting Data

# Whenever you use Invoke-Command to remotely execute code, you will notice that PowerShell automatically adds the column PSComputerName to your results. 
 # That's great because when you run Invoke-Command against more than one computer, you want to still know which computer returned particular information.

Invoke-Command { Get-Service } -ComputerName IIS-CTI5052

# The problem with PSComputerName is: it disappears whenever you use Select-Object to tailor the resulting columns:
Invoke-Command { Get-Service } -ComputerName IIS-CTI5052 | Select-Object PSComputerName, Name, Status

# To work around this problem, define a hash table to add the sender’s name. As a side effect, you now can also give the column a better name:
$pcName = @{ Name = "Machine"; Expression = {$_.PSComputerName} }
Invoke-Command { Get-Service } -ComputerName IIS-CTI5052 | Select-Object $pcName, Name, Status 

###########################################################################################################################################################################################

# Tip 95:　Getting NIC IP addresses and MAC addresses

function Get-IPandMAC($computername)
{
    $nickName = @{Name = "NICname"; Expression = {($_.Caption -split '] ')[-1]}}

    $ipv4 = @{Name = "IPv4"; Expression = {($_.IPAddress -like "*.*.*.*") -join ","}}

    $ipv6 = @{Name = "IPv6"; Expression = {($_.IPAddress -like "*::*") -join ","}}

    Get-WmiObject Win32_NetworkAdapterConfiguration @psboundparameters | Select-Object -Property $nickName, $ipv4, $ipv6, MacAddress | Where-Object {$_.MacAddress -ne $null}
}

Get-IPandMAC
Get-IPandMAC -computername IIS-CTI5052

###########################################################################################################################################################################################

# Tip 96: Get-IPandMAC

# returns the last process that was recently started, including the time in minutes and days it has been running and its description
function Get-LastProcessActivity
{
    $proc = Get-Process | Sort-Object StartTime  -ErrorAction SilentlyContinue -Descending | Select-Object -First 1
    
    $return = 1 | Select-Object Minutes, Days, Name, Description
    $timespan = New-TimeSpan $proc.StartTime
    $return.Days = $timespan.Days
    $return.Minutes = [Math]::Floor($timespan.TotalMinutes)
    $return.Name = $proc.Name
    $return.Description = $proc.Description
    
    $return
}

Get-LastProcessActivity

###########################################################################################################################################################################################

# Tip 97: Text-to-Speech

function Speak-Text([Parameter(Mandatory=$true)]$text, [switch]$drunk)
{
    $object = New-Object -ComObject SAPI.SpVoice

    if($drunk)
    {
        $object.Rate = -10
    }

    $object.Speak($text) | Out-Null
}

Speak-Text -text "Silence, I love you!"
Speak-Text -text "Silence, I love you!" -drunk





function Speak-Text([Parameter(Mandatory=$true)][string]$text, [switch]$async=$false)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.Speech") | Out-Null

    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    
    if($async)                                                                         # if you specify -Async the function will not block while the text is spoken
    {
        $synth.SpeakAsync($text)
    } 
    else
    {
        $synth.Speak($text)
    }
}

Speak-Text -text "Silence, love you so much."
Speak-Text -text "Silence, love you so much." -async $true

###########################################################################################################################################################################################

# Tip 98: Getting Information about Speed Traps and Traffic Jams

# PowerShell can read RSS feeds in just a couple of lines of code. Many radio broadcasters maintain RSS feeds with information about speed traps and traffic conditions.
 # Here is a (German) sample of how you can access and display such information via PowerShell:

$xml = New-Object xml

$xml.Load("http://www.radio7.de/index.php?id=76&type=100&no_cache=1")                                                    # the link is not accessable this current time

$xml.rss.channel.item | Where-Object {$_.category -eq 'Radar'} | Select-Object -ExpandProperty title
$xml.rss.channel.item | Where-Object {$_.category -eq 'Staus und Behinderungen'} | Select-Object -ExpandProperty title

###########################################################################################################################################################################################

# Tip 99: Use "E" in Numbers

# Note: To quickly define large integers, use the keyword "E" inside your number:

6E2                              # Output: 600
64E2                             # Output: 6400
644E6                            # Output: 644000000

###########################################################################################################################################################################################

# Tip 100: Forget the "Finally"-Block 

try
{
    dir nonexisting:\ -ErrorAction Stop
}
catch 
{
    "Error Occured: $_"
}
finally 
{
    'Doing Cleanup'
}

# Whenever a non-terminating error is raised in the try-block, the error can be handled in the catch-block. 
 # The finally-block always executes, whether the try-block encountered an error or not.
try
{
    dir nonexisting:\ -ErrorAction Stop
}
catch 
{
    "Error Occured: $_"
}

'Doing Cleanup'                       # Actually, you could have placed that code outside the finally-block as well and omit it

###########################################################################################################################################################################################