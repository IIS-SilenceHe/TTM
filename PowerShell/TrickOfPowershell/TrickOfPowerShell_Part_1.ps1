# Reference site: http://powershell.com/cs/blogs/tips/
###########################################################################################################################################################################################

# Tip 1: Closing all explorer windows

# Note: Shell.Application has many usefull method here, you can search it on internet for more details
(New-Object -ComObject Shell.Application).Windows() | Where-Object {$_.FullName -ne $null} | Where-Object {$_.FullName.toLower().Endswith("\explorer.exe")} | ForEach-Object {$_.Quit()}

###########################################################################################################################################################################################

# Tip 2: Assigning variables

$a = $b = $c = 0
$a,$b,$c = 1,2,3
$a,$b = $b,$a     # exchange $a and $b value

###########################################################################################################################################################################################

# Tip 3: Add descriptions to variables

$ip = "127.0.0.1"
Set-Variable ip -Description 'Local IP or Name'       # I made mistake here, should be "ip" but not "$ip"
dir variable:ip | Format-Table Name,Value,Description

###########################################################################################################################################################################################

# Tip 4: Making variables constant

$local = "127.0.0.1"
Set-Variable local3 -Option ReadOnly
$local = "10.10.10.10"                  # Exception here: value can't be updated for a readonly variable

Remove-Variable local -Force            # Note: readonly variable can be removed, but constant can neither be modified nor deleted

Set-Variable local -Option None -Force  # Note: option can be updated for a readonly variable

Set-Variable local -Option Constant -Value "127.0.0.1"  # Note: Constant can't be changed once created, you can't remove it

# Above statement will be failed if the variable exist before, so you can remove the previous one and then set it to constant like below:
if(Test-Path variable:local)
{
    Remove-Variable local -Force
}
Set-Variable local -Option Constant -Value "127.0.0.1"

$local = "2"                            # Exception
Set-Variable local -Option None -Force  # Exception here: Cannot overwrite variable local because it is read-only or constant
Remove-Variable local -Force            # Exception here: Cannot remove variable local because it is constant or read-only

###########################################################################################################################################################################################

# Tip 5: Permanent changes to environment vairables

$env:launched = $true  # Add a new environment variable, changes are not permanent and will be discarded once you close PowerShell

# To make permanent changes, you will need to access the static .NET methods in the Environment class

# Read environment:
[Environment]::GetEnvironmentVariable("Temp","User")     # Output: C:\Users\v-sihe\AppData\Local\Temp
[Environment]::GetEnvironmentVariable("Temp","Machine")  # Output: C:\Windows\TEMP

[Environment]::SetEnvironmentVariable("Launched",$true,"User")  # Permanent changes to environment vairables, the updated will worked if the powershell restart

###########################################################################################################################################################################################

# Tip 6: Strongly typed vairables

$date = "November 12, 2008"
$date.GetType().Name                   # Output: String

$date = [DateTime]"November 12, 2008"
$date.GetType().Name                   # Output: DateTime
$date.AddDays(10)                      # Output: Saturday, November 22, 2008 00:00:00
$date = "silence"
$date.GetType().Name                   # Output: String

[DateTime]$date = "November 12, 2008"
$date.AddDays(-6)                      # Output: Thursday, November 06, 2008 00:00:00

# Note: When you strongly type a variable, it no longer is versatile. Instead, it now only accepts data that can be converted into the type assigned to it
$date = "silence"                      # Exception here: Cannot convert value "silence" to type "System.DateTime"

###########################################################################################################################################################################################

# Tip 7: Using hash table

#Method 1:
$person = @{}
$person.Age = 24
$person.Name = "silence"
$person.Status = "online"

$person           # output all info
$person.Name      # Output: silence

$person["Age"]
$info = "Age"
$person.$info

#Method 2:
$dog = @{Name = "lele";Age = 3;}
$dog
$dog.Name
###########################################################################################################################################################################################

# Tip 8: Sorting hash table

$hash = @{Name = "Silence"; Age = 24; Status = "Online"}

$hash | Sort-Object Name  # doesn't work here

# To sort a Hash Table, you need to transform its content into individual objects that can be sorted by using GetEnumerator():
$hash.GetEnumerator() | Sort-Object Name

###########################################################################################################################################################################################

# Tip 9: Converting hash tables to objects

# Method 1:
function ConvertTo-Object($hashtable)
{
    $object = New-Object PSObject
    $hashtable.GetEnumerator() | ForEach-Object {
        Add-Member -InputObject $object -MemberType NoteProperty -Name $_.Name -Value $_.Value    
    }

    $object
}

$hash = @{Name = "Silence"; Age = 24; Status = "Online"}
$hash

ConvertTo-Object $hash

# Method 2:
# You can also turn this into a pipeline filter. It may be easier to pipe hashtables into this function, and also this way you can combine a number of hashtables into one object:

function ConvertTo-Object
{
    begin
    {
        $object = New-Object Object
    }

    process
    {
        $_.GetEnumerator() | ForEach-Object {
        Add-Member -InputObject $object -MemberType NoteProperty -Name $_.Name -Value $_.Value
        }
    }

    end
    {
        $object
    }
}

$hash = @{Name = "Silence"; Age = 24; Status = "Online"}
$hash2 = @{ID = 100;Remark = "Second hash table"}

$hash, $hash2 | ConvertTo-Object
$hash, $hash2 | ConvertTo-Object | Format-Table

# Method 3:
$hash = @{Name = "Silence"; Age = 24; Status = "Online"}
$obj = New-Object PSObject -Property $hash

# Note: [Ordered] keep the hast table sort in your way
$hash = @{Name = "Silence"; Age = 24; Status = "Online"}            # non-ordered: Status,Name,Age
$hash2 = [Ordered]@{Name = "Silence"; Age = 24; Status = "Online"}  # ordered: Name,Age,Status 

$hash = [Ordered]@{Name = "Silence"; Age = 24; Status = "Online"}
$hash.Insert(1,"Gender","Male")                                     # Insert new element for a hast table
$obj = New-Object PSObject -Property $hash

###########################################################################################################################################################################################

# Tip 10: Finding duplicate files

# Note: Hash Tables are a great way to find duplicates

$lookup = @{}
function Find-Duplicates
{
    $input | ForEach-Object {                          # as an example to learn how to use $input
        if($lookup.ContainsKey($_.Name.toLower()))
        {
            "$($_.FullName) is a duplicate in $($lookup.$($_.Name.toLower()))"
        }
        else
        {
            $lookup.Add($_.Name.toLower(),$_.FullName)
        }
    }
}

dir | Find-Duplicates
dir $env:windir | Find-Duplicates
dir C:\Windows\System32 | Find-Duplicates

###########################################################################################################################################################################################

# Tip 11: Converting results into arrays

 # Note:Whenever you call a function or cmdlet, PowerShell uses a built-in mechanism to handle results:
   #1.If no results are returned, PowerShell returns nothing 
   #2.If exactly one result is returned, the result is handed to the caller 
   #3.If more than one result is returned, all results are wrapped into an array, and the array is returned. 

(dir c:\nonexistent).GetType()    # Exception: Cannot find path 'C:\nonexistent' because it does not exist
(dir C:\logdata.dat).GetType()
(dir $env:windir).GetType()

# Note: use @() you can force PowerShell to always return an array
@(dir c:\nonexistent).GetType()    # Exception: Cannot find path 'C:\nonexistent' because it does not exist
@(dir C:\logdata.dat).GetType()
@(dir $env:windir).GetType()

@(dir c:\nonexistent).count        # Exception: Cannot find path 'C:\nonexistent' because it does not exist
@(dir C:\logdata.dat).count
@(dir $env:windir).count

@(dir c:\notexistent -ErrorAction SilentlyContinue).Count    # To get rid of these error messages: set -ErrorAction SilentlyContinue

###########################################################################################################################################################################################

# Tip 12: Working with arrays

$myArray = "hello",12,(Get-Date),$null,$true     # different type can be stored in a same variable
$myArray.Count

$myArray[0]
$myArray[-1]
$myArray[0,1,-1]
$myArray[0..2]

$myArray += "new element"
$myArray.Count

[int[]]$myArray = 1,2,3
$myArray += 12
$myArray += 14.789

$myArray[-1]
$myArray += "this won't work because it is not convertible to integer"    # Exception here: the string can't be translated to System.Int32

###########################################################################################################################################################################################

# Tip 13: Finding old files

filter FileAge($days)
{
    if(($_.CreationTime -le (Get-Date).AddDays($days * -1)))
    {
        $_
    }
}

dir $HOME\*.png | FileAge 10
dir $HOME\*.png | FileAge 10 | del -WhatIf   # -whatif try to show the results what the step will do but not real to execute it, this one used to keep safe

###########################################################################################################################################################################################

# Tip 14: Outptting calculated properties

dir | Format-Table Name,Length    # You can pick the object properties you want to output

dir | Format-Table Name, {[int]($_.Length/1kb)} -AutoSize   # output name and [int]($_.Length/1kb)

# Note: hash table can be used to format output
$mycolumn = @{expression = {[int]($_.Length/1kb)}; label = "KB"; width = 10; alignment = "left"}
dir | Format-Table Name, $mycolumn -AutoSize

###########################################################################################################################################################################################

# Tip 15: Finding the current user

[Environment]::UserDomainName + "\" + [Environment]::UserName
[Environment]::MachineName

# The dollar sign ($) indicates a variable and the drive name, "env:" indicates an environment variable.
$env:USERNAME
$env:USERDNSDOMAIN
$env:COMPUTERNAME

[Security.Principal.WindowsIdentity]::GetCurrent().Name   # Get "domain\user" via .Net method

###########################################################################################################################################################################################

# Tip 16: Finding out a script parent folder

$MyInvocation.MyCommand.Definition
$parent = Split-Path -Parent $MyInvocation.MyCommand.Definition   # Get the parent folder where the current script locate


###########################################################################################################################################################################################

# Tip 17: Finding system folders

[Environment]::GetFolderPath("Desktop")
[Environment]::GetFolderPath("Temp")

# Get all values which GetFolderPath() provided
[System.Environment+SpecialFolder] | Get-Member -Static -MemberType Property | ForEach-Object {"{0,-25} =  {1}" -f $_.Name, [Environment]::GetFolderPath($_.Name)}

###########################################################################################################################################################################################

# Tip 18: Downloading files from the internet

$object = New-Object Net.WebClient
$url = "http://download.microsoft.com/download/4/7/1/47104ec6-410d-4492-890b-2a34900c9df2/Workshops-EN.zip"
$localpath = "$home\powershellworkshop.zip"
$object.DownloadFile($url,$localpath)       # Note: no progress indicator here, it may take a couple of minutes until the command is processed

###########################################################################################################################################################################################

# Tip 19: Downloads with progress bar

[void][Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
$url = "http://download.microsoft.com/download/4/7/1/47104ec6-410d-4492-890b-2a34900c9df2/Workshops-EN.zip"
$localpath = "$home\powershellworkshop.zip"
$object = New-Object Microsoft.VisualBasic.Devices.Network
$object.DownloadFile($url,$localpath,"","",$true,500,$true,"DoNothing")

# In case you are wondering what kind of arguments DownloadFile takes, look up the method. It's easy:
[System.Diagnostics.Process]::Start("http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.devices.network.downloadfile.aspx")

[System.Diagnostics.Process]::Start("http://www.baidu.com")    # navigate to baidu site

###########################################################################################################################################################################################

# Tip 20: Using COM objects to say "Hi!"

$sam = New-Object -ComObject SAPI.SpVoice
$sam.Speak("How are you doing, dude?")
$sam.Speak((Get-Date))

# Look up all COM objects live on your computer
dir registry::HKEY_CLASSES_ROOT\CLSID -include ProgID -recurse | ForEach-Object {$_.GetValue("")}

###########################################################################################################################################################################################

# Tip 21: Manipulating arrays effectively

$array = New-Object System.Collections.ArrayList
$array.Count
$array.Add("a new element")
$array.Add((Get-Date))
$array.Add(100)
$array.Count

[void]$array.Add(200)  # Note: each time $array.Add(value) will retrun a index, use [void] can discard it

$array.Insert(0,"new element at the beginning!")   # Inserts a new element at the beginning

$array.Remove(200)                                 # Removes the first element with a value of 200
$array.RemoveAt(0)                                 # Removes the first element

$array[0]                                          # Get the first value

###########################################################################################################################################################################################

# Tip 22: Multidimensional arrays

$array1 = 1,2,(1,2,3),3
$array1[0]
$array1[1]
$array1[2]
$array1[2][0]
$array1[2][1]

$array2 = New-Object "Object[,]" 10,20
$array2[4,8] = "Hello"
$array2[9,16] = "Test"
$array2

###########################################################################################################################################################################################

# Tip 23: Validate user input

do
{
    $input = Read-Host "HomePage"

}while($input -notlike "www.*.*")

$warning = "Enter email address! >"
do
{
    Write-Host $warning -ForegroundColor Yellow -BackgroundColor Black -NoNewline
    $email = Read-Host
    $warning = "You did not enter a valid email address, Please try again!"

}while($email -notlike "*@*.*")

# Note: If you need to match more complex patterns, you should use -match and Regular Expressions.

###########################################################################################################################################################################################

# Tip 24: Show battery status as prompt for notebook

function prompt
{
    $charge = Get-WmiObject Win32_Battery -Property EstimatedChargeRemaining

    switch ($charge.EstimatedChargeRemaining)
    {
        {$_ -lt 25} {$color = "red"; break}
        {$_ -lt 50} {$color = "red"; break}
        default {$color = "white"}
    }

    $text = "PS {0} ({1}%)> " -f (Get-Location),$charge.EstimatedChargeRemaining

    Write-Host $text -NoNewline -ForegroundColor $color -BackgroundColor Black
}

###########################################################################################################################################################################################

# Tip 25: Using traps and error handling

Trap {"something awful happened!"}            # Without "continue" here, the script will throw its own error message and ignore your message
1 / $null

Trap {"something awful happened!"; continue}
1 / $null

Trap {"Something awful happened, to be more precise: $($_.Exception.Message)"; continue}  # this one provide a much more friendly message to show the error details
1 / $null

###########################################################################################################################################################################################

# Tip 26: Understanding trap scope

# Note: If you only have one scope, PowerShell continues execution with the next statement following the error
Trap { 'Something terrible happened.'; Continue}
'Hello'
1/$null
'World'


# Note: To create logical blocks, add scopes to your script, which can be functions or a basic script block. 
#  The next example will omit all remaining commands in the script block containing the error and PowerShell will continue with the next statement outside the script block
Trap { 'Something terrible happened.'; Continue}

&{
'Hello'
1/$null
'World'
}

'Outside'

###########################################################################################################################################################################################

# Tip 27: Understanding Exceptions (and why you can't catch some errors)

# Note: Traps are a great way of catching exceptions and handling errors manually but this does not seem to work all of the time

# this one works
Trap {"something awful happened!"; continue}
1 / $null

# below 2 not works
Trap { 'Something terrible happened.'; Continue}
1/0

Trap { 'Something terrible happened.'; Continue}
Dir c:\nonexistentfolder

# The reason: the last two examples did not raise an exception, and thus your Trap never noticed the error. 
    # Some very obvious errors such as 1/0 are discovered by the PowerShell parser and raises an error without even bothering to start the script (and your Trap). 
    # You cannot trap those and should fix these syntactical errors right away.
    # Other errors are handled by the commands internally. When you try and list a folder content of a folder that does not exist, 
        #  the Get-ChildItem cmdlet (alias Dir) will raise the error and handle it internally.

# Solution: So, for your traps to work, you will need to set the $ErrorActionPreference to 'Stop'. 
    # Only then will errors raise external exceptions that your traps can handle. You can set the error preference either individually with the -EA parameter

# The right way to use Trap to catch exceptions:

# Method 1:
Trap { 'Something terrible happened.'; Continue}
Dir c:\nonexistentfolder -ErrorAction Stop

# Method 2:
$ErrorActionPreference = 'Stop'
Trap { 'Something terrible happened.'; Continue}
Dir c:\nonexistentfolder

# Note: Reset $ErrorActionPreference if it needs when your scipt finished.

###########################################################################################################################################################################################

# Tip 28: Quick Loops

# Method 1:
for($x =1 ; $x -le 10; $x++)
{
    $x
}

# Method 2:
foreach($x in 1..10)
{
    $x
}

# Create your personal ASCII character reference table simply by casting the number to a character
foreach($x in 32..255)
{
    "$x = $([char]$x)"
}
# Output:
  #32 =  
  #33 = !
  #34 = "
  #35 = #
  #36 = $
  #37 = %
  #38 = &
  #39 = '
  #40 = (
  #41 = )
  #42 = *
  #...

# Create enumerated lists
foreach($x in 1..20)
{
    "Server{0:00}" -f $x   # Output: Server01 to Server20
}

###########################################################################################################################################################################################

# Tip 29: Arrays of strings
"-" * 50
"localhost " * 10
@("localhost") * 10

###########################################################################################################################################################################################

# Tip 30: Accessing static .Net

[System.Net.Dns]::GetHostByName("microsoft.com")

# Recommend tool: Using PowerShell Plus, you can find out what .NET classes and methods are available

###########################################################################################################################################################################################

# Tip 31: Finding comlets with a givenparameter

Get-Command *service* -CommandType Cmdlet   # Finding cmdlets by name

# Method 1:
Get-Help * -Parameter list # Try to find all cmdlets with a -List parameter

# Method 2:
filter Contains-Parameter
{
    param($name)

    $number = @($_ | ForEach-Object {$_.ParameterSets | ForEach-Object {$_.Parameters} | Where-Object {$_.Name -eq $name}}).Count

    if($number -gt 0)
    {
        $_
    }
}

Get-Command | Contains-Parameter "list"

###########################################################################################################################################################################################

# Tip 32: Stopping and disabling services

# Method 1:
# Temporarily disabling and then stopping the search service
Set-Service wsearch -StartupType Disabled
Stop-Service wsearch

# To trun the service back on:
Set-Service wsearch -StartupType Automatic
Start-Service wsearch

# Method 2: 
Stop-Service wsearch
Start-Service wsearch
Get-Service wsearch    # confirm the service is start

###########################################################################################################################################################################################

# Tip 33: Outputting nicely formatted dates

# Note: Get-Date provides you with the current date and time. With the -format parameter
Get-Date -Format d                                 # use -format with a lowercase d to just output a short date

Get-Date -Format "yyyy-MM-dd hh-mm-ss"             # Note: i had mistake here, "M" is for mouth and "m" for minute

(Get-Date -Format "yyyy-MM-dd hh-mm-ss") + ".log"  # you get time-stamped filenames for temporary or log files

# Format details: http://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo(VS.80).aspx

###########################################################################################################################################################################################

# Tip 34: Using cultures

# Note: PowerShell is culture-independent, you can pick any culture you want and use the culture-specific formats
$culture = New-Object System.Globalization.CultureInfo("zh-cn")
$number = 100
$number.ToString("c")             # Output: $100.00  [Current culture]
$number.ToString("c",$culture)    # Output: ￥100.00 [Chinese culture]

# CultureInfo Class Details: http://msdn.microsoft.com/en-us/library/system.globalization.cultureinfo.aspx
# Formatting Types Details:  http://msdn.microsoft.com/en-us/library/fbxft59x(VS.80).aspx

# Formate date with culture
(Get-Date).ToString()
(Get-Date).ToString("d",$culture)

###########################################################################################################################################################################################

# Tip 35: Accessing date methods

$date = Get-Date

$date.IsDaylightSavingTime()      # to find out if Daylight Savings Time is in effect(夏令时)
$date.AddDays(10)                 # try adding 10 days from today
$date.AddDays(-5)                 # go back 5 days

###########################################################################################################################################################################################

# Tip 36: Filtering based on file age

# One: filter file by lastwritetime property
Filter Select-FileAge
{
    param($days)

    if($_.PSisContainer)   # if folder, do nothing, only find file here
    {
        # do not return folders effectively filtering them out
    }
    elseif($_.LastWriteTime -lt (Get-Date).AddDays($days * -1))
    {
        $_
    }
}

dir $env:windir *.log | Select-FileAge 20    # Try to find .log files that are older than 20 days old

# Two: allowing a user to specify interactively which property to use for filtering
Filter Select-FileAge
{
    param($days, $property = "LastWriteTime")

    if($_.$property -lt (Get-Date).AddDays($days * -1))
    {
        $_
    }
}

dir $HOME | Select-FileAge 20 "CreationTime"   # user can decide with property to filter here

# Three: make the comparison dynamic: -gt -eq -ne -lt ...
Filter Select-FileAge
{
    param($days,$property = "LastWriteTime",$operator = "-lt")

    $condition = Invoke-Expression ("`$_.$property $operator (Get-Date).AddDays($days * -1)")   # Note: i made mistake here, should be "()" but not "{}" for invoke-expression
    if($condition)
    {
        $_
    }
}

dir $HOME | Select-FileAge 10 "CreationTime" "-ge"

###########################################################################################################################################################################################

# Tip 37: Order matters

$number = Read-Host "Enter amount in US dollars"
$rate = 0.7
$result = $number * $rate
"$number USD dquals $result EUR"    # it doesn't convert correctly, you always get back the result you entered

$number.GetType().FullName          # Output: System.String, Read-Host doesn't care what you type in and always returns it as text.

# Solution:

# Method 1: convert string to number
$number = [Double]("Enter amount in US dollars")  

# Method 2: change the order when you multiply
$result = $rate * $number

# Note: Whenever you calculate with different object types, PowerShell looks at the type of the first object. 
 # If you make sure this object is a number, and $rate is a number, then it will automatically convert the second object to the same type

###########################################################################################################################################################################################

# Tip 38: Casting a type without exception

# Note: you might cause your script to crash with an exception if you enter something that cannot be converted to a date
$date = [DateTime](Read-Host "Enter your birthday")
New-TimeSpan $date (Get-Date)
$days = (New-TimeSpan $date (Get-Date)).TotalDays
"You are {0:0} days old!" -f $days

# Solution: it might be safer to use the -as operator to convert a type as it does not throw an exception when conversion fails
$date = Read-Host "Enter your birthday"
$date = $date -as [DateTime]                 # return $null if conversion fails
if($date -eq $null)
{
    "You didn't enter a date!"
}
else
{
    New-TimeSpan $date (Get-Date)
    $days = (New-TimeSpan $date (Get-Date)).TotalDays
    "You are {0:0} days old!" -f $days
}

###########################################################################################################################################################################################

# Tip 39: Converting user input to date

# Note: PowerShell uses the US/English date format when converting user input to DateTime, which can cause unexpected results if using a different culture
 # To convert DateTime values based on your current culture, use the DateTime's parse() method

# if the user enters nonsense that cannot be converted to a date? You get an exception! here are 2 solution:

# Method 1: use -as to cast a type with exception
$date = Read-Host 'Enter your birthday'
if (($date -as [DateTime]) -ne $null) 
{  
    $date = [DateTime]::Parse($date) 
    $date
} 
else 
{  
    'You did not enter a valid date!'
}

# Method 2: use an error handler to catch the exception
$date = Read-Host "Enter your birthday"
trap
{
    "You did not enter a valid date!"
    continue
}

. {
    $date = [DateTime]::Parse($date)
    $date
}

###########################################################################################################################################################################################

# Tip 40: Free space on disks

Get-WmiObject Win32_LogicalDisk | ForEach-Object {"Disk {0} has {1:0.0} GB space abailable" -f $_.Caption, ($_.FreeSpace / 1GB)}

###########################################################################################################################################################################################

# Tip 41: Accessing individual WMI instances

Get-WmiObject Win32_LogicalDisk | Format-Table Name,FreeSpace             # all drives free space

[wmi]'Win32_LogicalDisk="C:"' | Format-Table Name, FreeSpace -AutoSize    # get C: free space, i had mistake here, there should no any spaces between '="C:"'
([wmi]"Win32_LogicalDisk='C:'").FreeSpace

[wmi].FullName     # Output: System.Management.ManagementObject

Get-WmiObject Win32_Service | Format-Table Name, __Path -Wrap             # find out the object path for a Service

[wmi]"Win32_Service='WSearch'" | Format-List *                            # to access a given Service directly

# Note: Why not use Get-Service? You could, but WMI is returning a lot more information about a service:
Get-Service WSearch | Format-List *

###########################################################################################################################################################################################

# Tip 42: Add custom properties

Get-WmiObject Win32_LogicalDisk | Format-Table Name,Size,FreeSpace -AutoSize

Get-WmiObject Win32_LogicalDisk | Format-Table Name, {$_.Size/1GB},{$_.FreeSpace/1GB} -AutoSize   # simply format the output info

$column1 = @{label = "Total Size (GB)"; expression = {[int]($_.Size / 1GB)}}
$column2 = @{label = "Free Space (GB)"; expression = {[int]($_.FreeSpace / 1GB)}}
Get-WmiObject Win32_LogicalDisk | Format-Table Name, $column1, $column2 -AutoSize                 # format the output info in a friendly way

###########################################################################################################################################################################################

# Tip 43: Outputting HTML reports

# Note: it is wise to use Select-Object to first limit the object properties to only those you want to see in your report, otherwise your HTML table gets huge

# format the html content
$head = '<style>
BODY{font-family:Verdana; background-color:lightblue;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{font-size:1.3em; border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:#FFCCCC}
TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:yellow}
</style>'

$header = "<H1>Last 24h Error Events</H1>"
$title = "Error Events Within 24 Hrs"

"127.0.0.1" | ForEach-Object {
    $time = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime((Get-Date).AddHours(-24))

    Get-WmiObject win32_NTLogEvent -computerName $_ -filter "EventType=1 and TimeGenerated>='$time'" | 
    ForEach-Object { $_ | Add-Member NoteProperty TimeStamp ([System.Management.ManagementDateTimeConverter]::ToDateTime($_.TimeWritten)) ; $_ }} | 
    Select-Object __SERVER, LogFile, Message, EventCode, TimeStamp | 
    ConvertTo-Html -head $head -body $header -title $title | Out-File $home\report.htm

    & "$home\report.htm"

###########################################################################################################################################################################################

# Tip 44: Converting numbers

$number = 123456416

[Convert]::ToString($number, 2)      # decimal to binary
[Convert]::ToString($number, 16)     # decimal to hexadecimal

$binary = '1110111000010001'
[Convert]::ToInt64($binary, 2)       # binary to decimal

###########################################################################################################################################################################################

# Tip 45: Counting items in a folder

@(dir $env:windir).Count             # counts the files in your Windows folder
@(dir $env:windir\*.log).Count       # To count just the number of log files

###########################################################################################################################################################################################

# Tip 46: Frouping folder items by extension (and more)

dir $env:windir | Group-Object Extension | Sort-Object Count

Get-Command | Group-Object Verb | Sort-Object Count
Get-Command | Group-Object Noun | Sort-Object Count

dir alias: | Group-Object definition | Sort-Object count

###########################################################################################################################################################################################

# Tip 47: Listing folders only (and finding special folders)

dir $env:windir -Recurse | Where-Object {$_.PSIsContainer}                                                 # output all folders (not file) under $env:windir

dir $env:windir -Recurse | Where-Object {$_.PSIsContainer} | Where-Object {$_.GetFiles().Count -eq 0}      # to show folders with no files (but possibly sub-folders) in them

dir $env:windir -Recurse | Where-Object {$_.PSIsContainer} | Where-Object {(dir $_.FullName).Count -eq 0}  # To find completely empty folders

dir $env:windir -Recurse | Where-Object {$_.PSIsContainer} | Where-Object {@(dir "$($_.FullName)\*.ps1").Count -ne 0} # find only folders taht contain .ps1 file

# Note: you may got error during filter process due to permission issue, set $ErrorAction = SilentlyContinue to geit rid of these issue

dir $env:windir -Recurse -ErrorAction SilentlyContinue | 
Where-Object {$_.PSIsContainer} | 
Where-Object {
    Write-Host "Examining $($_.FullName) ..." -ForegroundColor Yellow
    @(dir "$($_.FullName)\*.ps1" -ErrorAction SilentlyContinue).Count -ne 0
}

# Try to report the progress in a separate panel:
$c = 0
dir $env:windir -Recurse -ErrorAction SilentlyContinue | 
Where-Object {$_.PSIsContainer} | 
Where-Object {
    Write-Progress "Examining $($_.FullName) ..." "Found $c Folders ..." 1
    @(dir "$($_.FullName)\*.ps1" -ErrorAction SilentlyContinue).Count -ne 0
} | ForEach-Object {$c++; $_}

###########################################################################################################################################################################################

# Tip 48: Enumerating drive letters

$letters = 65..89 | ForEach-Object {([char]$_) + ":"}                                          # An easy way to create an array with drive letters
$letters | Where-Object {(!(Test-Path $_))}                                                    # Get the drive letters which unassigned

# Note: above approach is error prone because test-path will report the drive letters of CD/DVD-Drives without a medium as available

# Solution:
$letters | Where-Object {(New-Object System.IO.DriveInfo($_)).DriveType -eq "NoRootDirectory"} # get back all drive letters that are currently unused

@($letters | Where-Object {   (New-Object System.IO.DriveInfo($_)).DriveType -eq 'NoRootDirectory'})[0]  # To get only the first available letter

###########################################################################################################################################################################################

# Tip 49: Exploring privileges

function isAdmin
{
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    $admin = [System.Security.Principal.WindowsBuiltInRole]::Administrator
    $principal.IsInRole($admin)                                                      # will retrun true if the current process run as admin
}

if(isAdmin)
{
    $Host.UI.RawUI.BackgroundColor = "DarkRed"
    Clear-Host
}

# Note: It might also be a good idea to store the privilege status in a variable since your privilege will not change throughout a PowerShell session. 
 # Then, you can refer to it later without having to call the .NET code all the time
 $isAdmin = isAdmin

###########################################################################################################################################################################################

# Tip 50: Automatic aliases

# Note: All Get-Cmdlets (cmdlets that start with "Get") have an automatic type accelerator. You can use those cmdlets without the verb. 
 # So Childitem is the same as Get-Childitem, and service lists services just like Get-Service.

###########################################################################################################################################################################################

# Tip 51: Discover about-topics

Get-Help about_Operators   # to get a list of all available operators

Get-Help about_*           # get other Help topics which are available

# To actually open all Help topics in your Explorer
explorer "$pshome\$($host.CurrentCulture.Name)"
start "$pshome\$($host.CurrentCulture.Name)"

###########################################################################################################################################################################################

# Tip 52: Generate a new GUID

# Note: GUIDs are "Globally Unique Identifiers," which are so random that you can safely assume they are unique worldwide.
[System.Guid]::NewGuid().ToString()

###########################################################################################################################################################################################

# Tip 53： Ping and range ping

$object = New-Object System.Net.NetworkInformation.Ping
$object.Send("127.0.0.1")

# You can wrap it as function to get even more out of this:
function Ping-IP
{
    param($ip)

    Trap{$false; continue}

    $timeout = 1000
    $object = New-Object System.Net.NetworkInformation.Ping
    (($object.Send($ip,$timeout)).Status -eq "Success")
}

Ping-IP 127.0.0.1
Ping-IP "microsoft.com"
Ping-IP "zumsel.soft"

# You can add a loop to create your own network segment scan:
0..255 | ForEach-Object {$ip = "192.168.2.$_"; "$ip = $(Ping-IP $ip)"}

###########################################################################################################################################################################################

# Tip 54: Quick drive info

New-Object System.IO.DriveInfo "C:" | Format-List *

$drive = New-Object System.IO.DriveInfo "C:"
$drive.DriveFormat
$drive.VolumeLabel

###########################################################################################################################################################################################

# Tip 55: Exiting a function (using return)

function Get-NamedProcess($name = $null)
{
    if($name -eq $null)
    {
        Write-Host "Specify a name!" -ForegroundColor Red
        return                                              # To exit a function immediately, use the return statement
    }
    
    Get-Process
}

###########################################################################################################################################################################################

# Tip 56: Validating email address or ip address

# validating email address
function isEmailAddress($object)
{
    ($object -as [System.Net.Mail.MailAddress]).Address -eq $object -and $object -ne $null
}

isEmailAddress "silence"
isEmailAddress "silence@hotmail.com"
isEmailAddress $null

# validating ip address
function isIPAddress($object)
{
    ($object -as [System.Net.IPAddress]).IPAddressToString -eq $object -and $object -ne $null
}

isIPAddress "10"
isIPAddress "127.0.0.1"
isIPAddress "hello"
isIPAddress $null

###########################################################################################################################################################################################

# Tip 57: Limiting variables to a certain length

[String]$a = "hello"

# limits the variable a to text which is between 2 and 8 characters long
$lengLimit = New-Object System.Management.Automation.ValidateLengthAttribute -ArgumentList 2,8 
(Get-Variable a).Attributes.Add($lengLimit)

$a = "longer than 8 characters"                         # assigning text which is shorter or longer than the limit will raise an exception

###########################################################################################################################################################################################

# Tip 58: Limiting Variables to a set of values

# One: ValidateSetAttribute：限制变量的取值集合
$option = "yes"

$limitValue = New-Object System.Management.Automation.ValidateSetAttribute -ArgumentList "yes","no","perhaps"   # only "yes","no" or "perhaps" can be set to this variable
(Get-Variable option).Attributes.Add($limitValue)

$option = "no"
$option = "perhaps"
$option = "don't know"         # exception here

# Two: ValidateRangeAttribute：限制变量的取值范围
$month = 2

$con = New-Object System.Management.Automation.ValidateRangeAttribute -ArgumentList 1,12
(Get-Variable month).Attributes.Add($con)                                                   # only 1 to 12 is valid here

$month = 5

# exception here for below statement
$month = 0
$month = 13

###########################################################################################################################################################################################

# Tip 59: Keep variable not null

# One: ValidateNotNullAttribute：限制变量不能为空
$a = 123

$con = New-Object System.Management.Automation.ValidateNotNullAttribute
(Get-Variable a).Attributes.Clear()                                               # Clear other rules
(Get-Variable a).Attributes.Add($con)

$a = 9556
$a = ""        # null string is allowed here
$a = @()       # null array is allowed here

$a = $null     # Exception here


# Two: ValidateNotNullOrEmptyAttribute：限制变量不等为空，不能为空字符串，不能为空集合
$con2 = New-Object System.Management.Automation.ValidateNotNullOrEmptyAttribute
(Get-Variable a).Attributes.Clear()                                               # Clear other rules
(Get-Variable a).Attributes.Add($con2)

$a = 9556

# exception here for below values
$a = $null
$a = @()
$a = ""

# Note: During Add rules to attributes, if the rule has conflict with the current value, then the add operation will fail
 # For example: if $a = $null, then (Get-Variable a).Attributes.Add($con2) will throw exception

###########################################################################################################################################################################################

# Tip 60: Keep value only accept valid email address (use regex)

$email = "hhb@hotmail.com"

$reg = "/^[a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$/"
$con = New-Object System.Management.Automation.ValidatePatternAttribute -ArgumentList $reg     # ValidatePatternAttribute:限制变量要满足制定的正则表达式
(Get-Variable email).Attributes.Add($con)                                                      # looks i got exception here, didn't know why this current time

###########################################################################################################################################################################################

# Tip 61: Reversing array order

$a = ipconfig
$a

[array]::Reverse($a)
$a

# Note: To reverse the order of elements in an array, the most efficient way is to use the [Array] type and its static method Reverse()

###########################################################################################################################################################################################

# Tip 62: Validating a URL

function isURI($address)
{
    ($address -as [System.Uri]).AbsoluteURI -ne $null                       # retrun true or false
}

function isWebURI($address)
{
    $uri = $address -as [System.Uri]
    $uri.AbsoluteURI -ne $null -and $uri.Scheme -match "[http|https]"       # validate email is valid and starts with http or https
}

isURI('http://www.powershell.com')
isURI('test')
isURI($null)
isURI('zzz://zumsel.zum')

"-" * 50                                  # Dividing line

isWebURI('http://www.powershell.com')
isWebURI('test')
isWebURI($null)
isWebURI('zzz://zumsel.zum')

###########################################################################################################################################################################################

# Tip 63: Checking host name type

function Check-Hostname($name)
{
    [System.Uri]::CheckHostName($name)
}

Check-Hostname "127.0.0.1"                             # Output: IPv4
Check-Hostname "2001:0:d5c7:a2ca:89e:7bd:a865:5eb6"    # Output: IPv6
Check-Hostname "www.powershell.com"                    # Output: Dns
Check-Hostname "///"                                   # Output: Unknown

# NOte: CheckHostName() will return "Unknown" for any invalid host name. If you specify a valid host name, 
 # the method tells you whether it was a valid DNS name, an IPv4 or IPv6 IP address. It does not check whether the host exists.

###########################################################################################################################################################################################

# Tip 64: Escaping Text Strings

# Note: HTML on web pages uses tags and other special characters to define the page. To make sure text is not misinterpreted as HTML tags, 
 # you may want to escape text and automatically convert any ambiguous text character in an encoded format.
function Escape-String($string)
{
    [System.Uri]::EscapeDataString($string)
}

Escape-String "Hello World"                    # Output: Hello%20World
Escape-String "<h1>Hello World</h1>"           # Output: %3Ch1%3EHello%20World%3C%2Fh1%3E

# Note: To do the opposite and convert an encoded string back into the original representation, use UnescapeDataString():
function Unescape-String($string)
{
    [System.Uri]::UnescapeDataString($string)
}

Unescape-String "Hello%20World"                      # Output: Hello World
Unescape-String "%3Ch1%3EHello%20World%3C%2Fh1%3E"   # Output: <h1>Hello World</h1>

# Note: When working with web applications, you may want to escape characters so they are displayed correctly on the web

###########################################################################################################################################################################################

# Tip 65: Analyzing URLs

$result = [System.Uri]"http://www.baidu.com"
$result

$result.port
$result.Authority
$result.Scheme

###########################################################################################################################################################################################

# Tip 66: Sorting arrays

$array = 1,51,54,2,5,66,62
$array | Sort-Object                       # default sort: small to large
$array                                     # $array not change and show like before
 
$array = "sdfe", "s", "a", "F", "hello"
$array | Sort-Object
$array

# Note: If you'd like to sort the array permanently, you could assign the sort result back to the variable
$array = $array | Sort-Object

# A better way of sorting an array permanently is to use the Sort() method provided by System.Array:
$array = "sdfe", "s", "a", "F", "hello"
[Array]::Sort([array]$array)               # no output here, the array is permanently changed once called Sort()
$array

# Use Measure-Command to analyze performance of different approaches:
Measure-Command{
    1..1000 | % {
        $array = "sdfe", "s", "a", "F", "hello"
        $array = $array | Sort-Object
    }
}

Measure-Command{
    1..1000 | % {
        $array = "sdfe", "s", "a", "F", "hello"
        [Array]::Sort([array]$array)
    }
}

###########################################################################################################################################################################################

# Tip 67: Creating lists of letters

$letters = [char[]](97..122)
$letters                              # Output: a to z

###########################################################################################################################################################################################

# Tip 68: Creating random numbers

$random = New-Object Random
$random.Next(1,6)                                                                            # This gives you a new random number between 1 and 6 each time you call Next()

1..10 | ForEach-Object { 1..8 | ForEach-Object {$pwd = ""} {$pwd += [char]$random.Next(33,95)} {$pwd.toLower()}}    # get 10 random password

$char = [char[]](33..95)
$str = ""
foreach($c in $char)
{
    $str += $c   
}
$str                                                                                         # Output： !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_                                                                                    

# Note: Note that PowerShell V2 includes a new Cmdlet called Get-Random
1..8 | ForEach-Object {$pwd = ""} {$pwd += [char](33..95 | Get-Random)} {$pwd.ToLower()}     # Note: there can be many {...} statement blocks after foreach-object 

###########################################################################################################################################################################################

# Tip 69: Creating numeric ranges

# Trick 1:
1..10

# Trick 2:
$start = 100
$stop  = 200
$start..$stop

# Trick 3:
$array = 1..10
$array[1..$($array.Length - 1)]                                                              # output all values for $array except the first one

###########################################################################################################################################################################################

# Tip 70: Why upgrade array to arraylist


# One: Remove element from a array

# Note: Simple arrays have no built-in mechanism to insert new elements or extract elements at given positions
$array = 1..10

$array = $array[0..3] + $array[5..9]             # remove 5 from array
$array

# Solution: change array to arraylist, then you can get the method from arraylist to remove a value more quickly
$betterArray = [System.Collections.ArrayList]$array
$betterArray.RemoveAt(4)
$betterArray
$array = $betterArray.ToArray()
$array


# Two: Keep array write-protected
$array2 = 1..10
$array2 += 11
$array2                                                               # Output: 1 to 11
$array2.GetType().FullName                                            # Output: System.Object[]

$arraylist = [System.Collections.ArrayList]$array2
$readOnly = [System.Collections.ArrayList]::ReadOnly($arraylist)      # make the arraylist read only

$readOnly.GetType().FullName                                          # Output: System.Collections.ArrayList+ReadOnlyArrayList
$readOnly
$readOnly.IsReadOnly                                                  # Output: true
$readOnly.Add("hello")                                                # exception here: Collection is read-only

# Note: write-protection is weak. Once you add another element to the array using the += operator, the write-protection breaks
$readOnly += 12
$readOnly.IsReadOnly                                                  # Output: false

# Reason: PowerShell has silently killed the write-protected ArrayList and replaced it with a new simple array in order to add the new element
$readOnly.GetType().FullName                                          # Output: System.Object[], type has changed here

###########################################################################################################################################################################################

# Tip 71: Calculating space consumption

Get-ChildItem $HOME | Where-Object {$_.PSIsContainer} | 
    ForEach-Object {Write-Progress "Examining folder" ($_.FullName); $_} | 
    ForEach-Object {$result = "" | 
        Select-Object Path,Count,Size

        $result.path = $_.FullName
        
        $temp = Get-ChildItem $_.FullName -Recurse -ErrorAction SilentlyContinue | Measure-Object length -Sum -ErrorAction SilentlyContinue
        
        $result.count = $temp.Count
        $result.size = $temp.Sum

        $result
}

###########################################################################################################################################################################################

# Tip 72: Setting Properties on AD Users

$seacher = New-Object DirectoryServices.DirectorySearcher([ADSI]"","(&(objectCategory=person)(objectClass=user)(msNPAllowDialin=FALSE))")   # should no any space in the query string
$seacher.PageSize = 1000          # Note: here one man said, changing the PageSize property let's you search through more users/groups/computers then the default value (1000).

$seacher.FindAll() | ForEach-Object {$_.GetDirectoryEntry()} | ForEach-Object {$_.PutEx(1,"msNPAllowDialin",0); $_.SetInfo()}               # get exception here, it said access is denied

###########################################################################################################################################################################################

# Tip 73: Comparing results

$base = Get-Process
notepad
$compare = Get-Process

Compare-Object $base $compare                      # SideIndicator: "=>" Add      "<=" Delete

$shot1 = dir $HOME
Set-Content $HOME\testfile1.txt "A new file"
$short2 = dir $HOME

Compare-Object $shot1 $short2

# Note: Compare-Object will return wrong results when there are more than 10 consecutive differences in both snapshots because it uses an internal "SyncWindow" of +- 5 elements.
 # In this case, increase the sync window using the -syncWindow parameter
Compare-Object $shot1 $short2 -SyncWindow 20

###########################################################################################################################################################################################

# Tip 74: Finding folder changes

$shot1 = dir $HOME
Set-Content $HOME\testfile1.txt "some content"
$shot2 = dir $HOME

Compare-Object $shot1 $short2 -SyncWindow 30

# Note: This works fine for new and deleted files, but it will not show you files that changed content
$shot1 = dir $HOME
Add-Content C:\Users\v-sihe\testfile1.txt "A new line"
$shot2 = dir $HOME

Compare-Object $shot1 $shot2 -SyncWindow 30                           # output nothing, it will not monitor the file changes

# By default, Compare-Object uses only the file name for comparison. To include more properties, use the -property parameter
Compare-Object $shot1 $shot2 -syncWindow 30 -property Name, Length

###########################################################################################################################################################################################

# Tip 75: Persisting comparison snapshots

# Note: Persisted result sets do not live in memory but are instead written to xml as file
dir $HOME | Select-Object Name, Length | Export-Clixml $home\baseline.xml
$shot3 = Import-Clixml $HOME\baseline.xml

Set-Content $HOME\testfile1.txt "hello world"
dir $HOME | Select-Object Name, Length | Export-Clixml $HOME\temp.xml
$shot4 = Import-Clixml $HOME\temp.xml

Compare-Object $shot3 $shot4 -Property Name, Length

###########################################################################################################################################################################################

# Tip 76: Advanced Compare-Object: Working with Results

$base = Get-Process
notepad
$new = Get-Process

# Note: To get the newly started processes, use -passthru and filter for SideIndicator equals "=>" to limit the output to only newly started processes 
 # (otherwise you would also get processes that no longer run):
Compare-Object $base $new -PassThru | Where-Object {$_.SideIndicator -eq "=>"}     # Because you used -passThru, you get back the actual process objects with all of their rich information

###########################################################################################################################################################################################

# Tip 77: Cleaning Document Folders

# Clean doc to different folder according to its extension
$folder = "C:\Users\v-sihe\Desktop\Temp"
$types = ".bmp", ".jpg", ".doc", ".ps1", ".ocx", ".pdf",   ".docx", ".txt", ".vbs", ".xls", ".zip", ".htm", ".bat"

$files = dir $folder | Where-Object {-not $_.PSIsContainer} | Group-Object Extension                                # group files accroding to its extension
$files = $files | Where-Object {$types -contains $_.Name}
$files | ForEach-Object {New-Item -ItemType Directory -Path "$folder\$($_.Name)" -ErrorAction SilentlyContinue}
$files | ForEach-Object {$_.Group | Copy-Item -Destination "$folder$($_.Extension)\$($_.Name)"}

###########################################################################################################################################################################################

# Tip 78: Listing Your Media Collection

$player = New-Object -ComObject WMPLAYER.OCX                            # Windows Media Player uses WMPLAYER.OCX as its ProgID

$all = $player.mediaCollection.getAll()
0..$($all.count - 1) | ForEach-Object {$all.Item($_)}                   # This will get you all media organized by WMP

$player.mediaCollection.getByAttribute("MediaType", "audio")            # To retrieve all audio playlists

$all = $player.mediaCollection.getByAttribute("MediaType", "playlist")  # Other attribute values are "playlist", "radio", "video", "photo" and "other"
0..$($all.count - 1) | ForEach-Object {$all.Item($_)}

###########################################################################################################################################################################################

# Tip 79: Playing a Song with Media Player

# Note: Windows Media Player can be accessed using COM, and WMP in turn gives you access to your entire media
$player = New-Object -ComObject WMPLAYER.OCX

$all = $player.mediaCollection.getAll()
0..$($all.count - 1) | ForEach-Object {$all.Item($_)}

$name = Read-Host "Enter media name/song name/playlist name"
$item = $player.mediaCollection.getByName($name)

if($item.Count -gt 0)
{
    $filename = $item.item(0).sourceurl
    "Playing $filename ..."
    $player.openPlayer($filename)
}
else
{
    "'$name' not found."
}

###########################################################################################################################################################################################

# Tip 80: Sorting test files

$file = "$HOME\serverlist.txt"
Set-Content $file "server1`nserver9`nserver5`nserver2`n"

Get-Content $file | Sort-Object | Set-Content $file           # Output: server1,server2,server5,server9

###########################################################################################################################################################################################

# Tip 81: Checking File and Folder Permissions

# Note: Get-Acl is a convenient Cmdlet to expose NTFS file and folder settings
dir $HOME | Get-Acl               # to get a list of ownerships for a folder content

# To find out which "Identities" (Users or groups) have specific permissions granted, 
 # access the Access property of each Acl to see the actual permissions, then group by IdentityReference
dir $HOME | Get-Acl | ForEach-Object {$_.Access} | Group-Object IndentifyReference

###########################################################################################################################################################################################

# Tip 82: Converting ASCII and Characters

[char]65                                    # To convert the ASCII value to a character: 65 -> char

[int][char]"A"                              # convert a character to its ascii value: A -> char -> int

[int[]][char[]]"Hello"                      # To process more than one character at a time, use arrays instead: hello -> char[] -> int[]

###########################################################################################################################################################################################

# Tip 83: Reading Text Files

$text = Get-Content $env:windir\WindowsUpdate.log

# Note: Get-Content reads the file line by line so you will most likely get back an array, This can cause problems when you really need the entire file content as text. 
$text.GetType().FullName
$text -is [array]                # Output: true

# Solution: If you need the complete text in one piece, try combining Get-Content with Out-String, which converts the array into one single string
$text = Get-Content $env:windir\WindowsUpdate.log | Out-String
$text -is [array]

# Powershell3.0以后GC有了新的参数-Raw。它不仅加快了读取大文件也能返回指定的一段原文件内容，没有分割它里面的行
$text = Get-Content $env:windir\WindowsUpdate.log -raw
$text -is [array]                               

###########################################################################################################################################################################################

# Tip 84: Analyzing Event Logs

# to get only events from the System log with an Event Type equals 1 (only errors)
Get-WmiObject Win32_NTLogEvent -Filter "LogFile='System' and EventType=1" | Format-Table ComputerName, EventCode, Message, TimeWritten   # only returned the properties which you want

Get-WmiObject Win32_NTLogEvent -Filter "LogFile='System' and EventType=1" | Format-Table *                                               # use * will return all properties

# filter events with a specific EventCode value
Get-WmiObject Win32_NTLogEvent -Filter "LogFile='System' and EventType=1 and EventCode=7022" | Format-Table ComputerName, EventCode, Message, TimeWritten

# filter events with all EventCodes between 7000 and 7999
Get-WmiObject Win32_NTLogEvent -Filter "LogFile='System' and EventType=1 and EventCode >= 7022 and EventCode < 7999"  | Format-Table ComputerName, EventCode, Message, TimeWritten


# convert a regular time expression into WMI format, below generates the WMI time of now minus 24 hours
$time = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime((Get-Date).AddHours(-24))  
$time

# to see all error events from all eventlogs that have occurred within the past 24 hours
Get-WmiObject Win32_NTLogEvent -Filter "EventType=1 and TimeGenerated >= '$time'" | Format-Table LogFile, Message, EventCode, TimeGenerated -Wrap

###########################################################################################################################################################################################

# Tip 85: Accessing Servers Remotely via WMI

Get-WmiObject Win32_OperatingSystem                                                 # get details for local operating system


# Note: Simply specify the computer name or an IP address and make sure you have appropriate privileges and that no firewall is blocking WMI or RPC
Get-WmiObject Win32_OperatingSystem -ComputerName "IIS-CTI5052"                     # Use the parameter -computer like this to get the very same information from a remote system

$cred = Get-Credential
Get-WmiObject Win32_OperatingSystem -ComputerName "IIS-CTI5052" -Credential $cred   # If you need to log-on to the target machine as a different user, add the -credential parameter    

###########################################################################################################################################################################################

# Tip 86: Finding Out Interesting WMI Classes

Get-WmiObject -List                                 # lists all available WMI classes in the default namespace

Get-WmiObject -List | Select-String Print           # use select-string to find all WMI classes related to printing

###########################################################################################################################################################################################

# Tip 87: Shutting Down Computers Remotely

# Note: you can forcefully shut down a remote system as long as you have appropriate privileges and there is no firewall blocking your way
$os = Get-WmiObject Win32_OperatingSystem -ComputerName 10.10.10.10
$os.Win32Shutdown(6,0)

###########################################################################################################################################################################################

# Tip 88: Creating A Computer Profile

# Note: Often, information needed to comprehensively profile a computer comes from a number of sources. 
 # A great way to meld these different bits of information together is by creating a new result object with just the properties you need. 
 # Simply,, use any simple object like a number or an empty string, and append the needed properties with Select-Object
$info = 0 | Select-Object Name, OS, SP, Hotfixes, Software, Lastboot, Services, Model
$info.SP = "hello"
$info

# Details: use WMI and other technologies to automatically fill your new object
function Get-ComputerInfo
{
    param($server = "127.0.0.1")

    [System.Reflection.Assembly]::LoadWithPartialName("System.ServiceProcess") > $null

    $info = 0 | Select-Object Name, OS, SP, Hotfixes, Software, Lastboot, Services, Model

    $os = Get-WmiObject Win32_OperatingSystem -ComputerName $server

    $info.Name = $server
    $info.OS = $os.Version
    $info.sp = $os.ServicePackMajorVersion
    $info.Lastboot = $os.ConvertToDateTime($os.LastBootUpTime)

    # Note: get hotfixes and software will take a couple of minutes to executed success, so keep them commented as needed
    $info.Hotfixes = @(Get-WmiObject Win32_QuickFixEngineering -ComputerName $server | Select-Object HotfixID, InstalledBy, InstalledOn | Where-Object {$_.InstalledOn})
    $info.Software = @(Get-WmiObject Win32_Product -ComputerName $server) | Select-Object Name, Version

    $info.Services = [ServiceProcess.ServiceController]::GetServices($server)
    $info.Model = (Get-WmiObject Win32_ComputerSystem -ComputerName $server).Model

    $info
}

###########################################################################################################################################################################################

# Tip 89: Using PowerShell To Create Batch Command Calls

$description = "Internal_VLAN"
$segment = 44
$pcName = "$description-Reserved"
$ip1 = "15.77.$segment."
$scopeName = "$($ip1)0"

10..126 | ForEach-Object {
    $ip = "$ip1$_"
    $mac = "001025{0:000}{1:000}" -f $segment, $_

    "netsh DHCP server 10.25.64.22 scope $scopeName and reservedip $ip $mac $pcName"            # only show message but not execute, below commented command will exec the command
    # &"netsh DHCP server 10.25.64.22 scope $scopeName andd reservedip $ip $mac $pcName"        # PowerShell executes the string when you append a "&" to a string
}

###########################################################################################################################################################################################

# Tip 90: Parsing Logfiles With Regular Expressions

# Note: Remember that Get-Content reads text line by line, returning an array. 
 # This is why the code can use Where-Object and the simple -like operator to kick out any line not containing a specific key word
Get-Content $env:windir\WindowsUpdate.log | Where-Object {$_ -like "*WARNING*"}  # The -like operator does not require complex regular expressions so you can use all the simple wildcards 


# use Regular expressions to get the specified result
Get-Content $env:windir\WindowsUpdate.log | Where-Object {$_ -like "*WARNING*"} | 
Where-Object {$_ -match '(.*?)\t(.*?)\t(.*?)\t(.*?)\t(.*?)\t(.*)'} | 
ForEach-Object {
    $result = 1 | Select-Object Date, Time, Origin, Message
    
    $result.Date = $Matches[1]
    $result.Time = $Matches[2]
    $result.Message = $Matches[6]
    $result.Origin = $Matches[5]

    $result
}

###########################################################################################################################################################################################

# Tip 91: Finding and Deleting Orphaned Shares

Get-WmiObject Win32_Share | Where-Object {$_.Path -ne ""} | Where-Object { -not (Test-Path $_.Path)}                                  # get the shares that have no target folder anymore

Get-WmiObject Win32_Share | Where-Object {$_.Path -ne ""} | Where-Object { -not (Test-Path $_.Path)} | ForEach-Object {$_.Delete()}   # remove shares that have no target folder anymore

# Note: Note that you may need administrator privileges, depending on who created the share.

###########################################################################################################################################################################################

# Tip 92: Clone NTFS Permissions

# Note: NTFS access permissions can be complex and tricky. 
 # To quickly assign NTFS permissions to a new folder, you can simply clone permissions from another folder that you know has the correct permissions applied

# Method 1:
md $HOME\sample

# manually assign correct permissions to folder "sample"

md $HOME\newFolder
Get-Acl $HOME\sample | Set-Acl -Path $HOME\newFolder



# Method 2:
md $home\sample
# manually assign correct permissions to folder "sample"
$sddl = (Get-Acl $home\sample).Sddl
md $home\newfolder

$sd = Get-Acl $home\newfolder
$sd.SetSecurityDescriptorSddlForm($sddl)
$sd.Sddl
Set-Acl $home\newfolder $sd

###########################################################################################################################################################################################

# Tip 93: Finding Current Script Path

Split-Path -Parent $MyInvocation.MyCommand.Definition     # get the folder your current PowerShell script is located 

###########################################################################################################################################################################################

# Tip 94: An Easy InputBox

# Input Box:
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
$name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your name", "Name", "$env:username")

"Your name is $name"


# Choice Box:
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
$result = [Microsoft.VisualBasic.Interaction]::MsgBox("Do you agree?", "YesNoCancel,Question", "Respond please")

switch($result)
{
    "Yes"    {"Ah good"}
    "No"     {"Sorry to hear that"}
    "Cancel" {"Bye..."}
}

###########################################################################################################################################################################################

# Tip 95: Check Online Status

filter Check-Online
{
    trap {continue}
    . {
        $timeout = 1000
        $obj = New-Object System.Net.NetworkInformation.Ping
        $result = $obj.Send($_, $timeout)

        if($result.Status -eq "Success")
        {
            $_
        }
    }
}

"127.0.0.1", "noexists", "powershell.com" | Check-Online

# Note that systems may still be online yet ignore ping requests due to security restrictions

###########################################################################################################################################################################################

# Tip 96: Sending Simple SMTP Mail

# Note: 在低版本的PowerShell上发送邮件可以借助.NET的system.net.mail.smtpclient类
Send-MailMessage -From hhbstar@hotmail.com -Subject test -To hhbstar@hotmail.com -Body "send email by poweshell" `
-Credential hhbstar@hotmail.com -Port 587 -Priority Low -SmtpServer smtp.live.com -UseSsl

# About Hotmail:
  # 接收：pop3.live.com，端口：995，安全协议：SSL（用户名要补全）
  # 外发：smtp.live.com，端口：587，安全协议：TLS

###########################################################################################################################################################################################

# Tip 97: Select Folder-Dialog

function Select-Folder
{
    param($message = "select a folder", $path = 0)

    $object = New-Object -ComObject Shell.Application
    $folder = $object.BrowseForFolder(0, $message, 0, $path)
    
    if($folder -ne $null)
    {
        $folder.self.Path
    }
}

Select-Folder
Select-Folder "Select the folder you want!"
Select-Folder -message "Select some folder!" -path $env:windir             # the path parameter initialize the default path 

###########################################################################################################################################################################################

# Tip 98: Detect DHCP State

Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled=true and DHCPEnabled=true"   # get the network adapters which using DHCP

###########################################################################################################################################################################################

# Tip 99: Accessing Internet Explorer

$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$ie.navigate("http://www.baidu.com")
$ie

$ie.document

# Note: Unfortunately, this approach fails when you use IE with enhanced security (Vista UAC for example). 


# Solution: To work around this issue, you can use yet another COM object called Shell.Application
& "$env:ProgramFiles\Internet Explorer\iexplore.exe" "http://powershell.com"

$win = New-Object -ComObject Shell.Application
$try = 0
$ie2 = $null

do
{
    Start-Sleep -Milliseconds 500

    $ie2 = @($win.windows() | Where-Object {$_.locationName -like "*PowerShell*"})[0]
    $try ++

    if($try -gt 10)
    {
        Throw "Web page cannot be opened."
    }

}while($ie2 -eq $null)

$ie2.Document
$ie2.Document.body.innerHTML

###########################################################################################################################################################################################

# Tip 100: Finding Out Whether A Web Page Is Open

Function Check-BrowserURL($url)
{
    $win = New-Object -ComObject Shell.Application
    $count = @($win.Windows() | Where-Object {$_.locationName -like "*$url*"}).Count

    if($count -ne 0)
    {
        "The site $url was opened"
    }
}

Check-BrowserURL powershell.com
Check-BrowserURL google.com

###########################################################################################################################################################################################