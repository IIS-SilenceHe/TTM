# Reference site: http://powershell.com/cs/blogs/tips/
###########################################################################################################################################################################################

# Tip 1: Accessing individual Files and Folders Remotely via WMI

$systemDrive = $env:SystemDrive

Get-WmiObject Win32_Directory -Filter "Drive='$systemDrive' and Path = '\\'" | Format-Table Name                                           # get all folders under path $env:systemDrive
Get-WmiObject CIM_DataFile -Filter "Drive='$systemDrive' and Path='\\Windows\\' and Extension='log'" | Format-Table Name                   # retrieves all log files in your Windows folder
Get-WmiObject CIM_DataFile -Filter "Drive='$systemDrive' and Path='\\Windows\\' and Extension='log'" | Format-Table *                      # get all properties but not just name

# Note: Avoid using the like operator. The next line would get you all log files in the Windows folder and recursively in all sub-folders. 
 # However, WMI would now query every single file which takes a very long time

# Get-WmiObject CIM_DataFile -Filter "Drive='$systemDrive' and Path like '\\Windows\\%' and Extension='log'" | Format-Table Name           # this statement may takes amount of time

###########################################################################################################################################################################################

# Tip 2: Encrypting PowerShell Scripts

function Encrypt-Script($path, $destination)
{
    $script = Get-Content $path | Out-String                           # out-string keep the content is a string but not a array here
    $secure = ConvertTo-SecureString $script -AsPlainText -force
    $export = $secure | ConvertFrom-SecureString
    
    Set-Content $destination $export
    "Script '$path' has been encrypted as '$destination'"
}


function Execute-EncryptedScript($path)
{
    trap { "Decryption Failed."; break }

    $raw = Get-Content $path
    $secure = ConvertTo-SecureString $raw

    $helper = New-Object System.Management.Automation.PSCredential("test", $secure)
    $plain = $helper.GetNetworkCredential().Password

    Invoke-Expression $plain
}

# original.ps1
    # $systemDrive = $env:SystemDrive
    # Get-WmiObject Win32_Directory -Filter "Drive='$systemDrive' and Path = '\\'" | Format-Table Name


Encrypt-Script .\original.ps1 .\secure.bin         
Execute-EncryptedScript .\secure.bin

# send parameter(s) to encrypted script:                                                                 [i tried but not success yet]

   # Change the Encrypt-Script statment to encrypt your file as another ps1.
   # 
   # Take all the parameter declarations out of your original script.
   # 
   # Add them to the arguments on the execute-encryptedscript function e.g.
   # 
   # function Execute-EncryptedScript($path, $yourvariable, $yourvariable2)
   # 
   # set your variables like this:
   # 
   # Execute-EncryptedScript $home\secure.ps1 -yourvariable "value1" - yourvariable2 "value2"

###########################################################################################################################################################################################

# Tip 3: Encrypting Scripts With A Password

# Encrypt the script and anyone who knows the password can execute this script                           
function Encrypt-Script($path, $destination) 
{  
    $script = Get-Content $path | Out-String  

    $key = Read-Host "Enter secret key (at least 16 characters)" -asSecureString  
    if ($key.Length -lt 16) { Throw "key needs to be at least 16 characters" }  

    $helper = New-Object system.Management.Automation.PSCredential("test", $key)  
    $key = $helper.GetNetworkCredential().Password  

    $secure = ConvertTo-SecureString $script -asPlainText -force  
    $export = $secure | ConvertFrom-SecureString -Key (([int[]][char[]]$key)[0..15])  

    Set-Content $destination $export  
    "Script has been encrypted as '$destination'"
}

function Execute-EncryptedScript($path) 
{  
    trap { Write-Host -fore Red "You are not authorized to decrypt and execute this script"; continue }  

    & {    
        $raw = Get-Content $path    

        $key = Read-Host "Enter Passphrase" -asSecureString    
        $helper = New-Object system.Management.Automation.PSCredential("test", $key)    
        $key = $helper.GetNetworkCredential().Password    
        $secure = ConvertTo-SecureString $raw -key (([int[]][char[]]$key)[0..15]) -ea Stop  
          
        $helper = New-Object system.Management.Automation.PSCredential("test", $secure)    
        $plain = $helper.GetNetworkCredential().Password   
         
        Invoke-Expression $plain  
      }
}

Encrypt-Script .\original.ps1 .\secure.bin         
Execute-EncryptedScript .\secure.bin

###########################################################################################################################################################################################

# Tip 4: Auto-Documenting Script Variables

# get all variables from a script and list them with sorted order
function Get-ScriptVariables($path)
{
    $result = Get-Content $path | ForEach-Object {
        if($_ -match "(\$.*?)\s=")
        {
            $Matches[1] | Where-Object {$_ -notlike "*.*"}
        }
    }
    $result | Sort-Object | Get-Unique
}

Get-ScriptVariables '.\HandleExcel - V2.ps1'

$path = '.\HandleExcel - V2.ps1'

###########################################################################################################################################################################################

# Tip 5: Creating Text Files

# Method 1: use classic redirection
"Hello" > $HOME\testfile.txt
"Append this" >> $HOME\testfile.txt
Get-Content $HOME\testfile.txt

# Method 2: To control encoding, use Set-Content
"Hello" | Set-Content $HOME\testfile.txt -Encoding Unicode
"Add this" | Add-Content $HOME\testfile.txt -Encoding Unicode
Get-Content $HOME\testfile.txt

# Method 3: use Out-File
"Hello" | Out-File $HOME\testfile.txt -Encoding unicode
"Add this" | Out-File $HOME\testfile.txt -Encoding unicode -Append

# Method 4: use native applications and tools like fsutil [Note: that tools like fsutil require administrator privileges]
fsutil createNew $HOME\testfile.txt 1000                          #  creates a blank text file with a size of 1,000 bytes

###########################################################################################################################################################################################

# Tip 6: Create a large file quickly

$tempFile="largeFile"
$fs=New-Object System.IO.FileStream($tempFile,[System.IO.FileMode]::OpenOrCreate)
$fs.Seek(2GB,[System.IO.SeekOrigin]::Begin)
$fs.WriteByte(0)
$fs.Close()
 
#生成完毕后，还可以检验下
(Get-Item $tempFile).Length/1gb
<#
 # 输出为：
 # 2.00000000093132
#>

###########################################################################################################################################################################################

# Tip 7: Save (and Load) Current PowerShell Configuration

Export-Console $HOME\console.psc1
Get-Content $HOME\console.psc1

# Simply launch PowerShell with the parameter -psconsolefile and specify the path to your .psc1 file
powershell.exe -PSConsoleFile $HOME\console.psc1 -file $HOME\test.ps1

###########################################################################################################################################################################################

# Tip 8: Finding PowerShell Background Information

Invoke-Item "$pshome\Documents\$($host.CurrentUICulture.Name)"      # on Vista (Multi-Language):
Invoke-Item "$pshome"                                               # on XP or above

Get-Help about_*                                                    # to see a list of all topics

###########################################################################################################################################################################################

# Tip 9: Setting and Deleting Aliases

Set-Alias edit notepad.exe                     # or: Set-Alias -Name edit -Value notepad.exe
edit                                           # will open notepad
$Alias:edit                                    # get the value for edit alias
                                               
Remove-Item alias:edit                         # or: del alias:edit

# export and import alias
# Method 1:
Export-Alias $HOME\myaliases.txt
Import-Alias $HOME\myaliases.txt -Force        # there maybe exception here if you are missing -force parameter, or: Import-Alias $HOME\myaliases.txt -ea SilentlyContinue

# Method 2:
Export-Alias $home\myaliases.ps1 -as Script
. $home\myaliases.ps1                          # run the .ps1 script to import the alias

###########################################################################################################################################################################################

# Tip 10: Create PowerShell Shortcuts

function Create-Shortcuts()
{
    $wsshell = New-Object -ComObject WScript.Shell
    $path1 = $wsshell.SpecialFolders.Item("Desktop")
    $path2 = $wsshell.SpecialFolders.Item("Programs")

    $path1, $path2 | ForEach-Object {
        $link = $wsshell.CreateShortcut("$_\Powershell.lnk")
        $link.TargetPath = "$pshome\powershell.exe"
        $link.Description = "launches Windows PowerShell console"
        $link.WorkingDirectory = $HOME
        $link.IconLocation = "$pshome\powrshell.exe"
        $link.Save()
    }
}
# http://powershell.com/cs/blogs/tips/default.aspx?PageIndex=93
###########################################################################################################################################################################################

# Tip 11: PowerShell Essentials: Get-Command

Get-Command                                                                                         # get all commands that powershell provides

Get-Command -Verb Get
Get-Command -Noun Set

Get-Command -Noun *print*                                                                           # To find all cmdlets that are related to printing
Get-Command *print*                                                                                 # get the list includes any executable that PowerShell can find
Get-Command *print* -CommandType Application | Where-Object {$_.definition -like "*.exe"}           # get .exe related command for print
Get-Command *print* -CommandType Application | Group-Object {$_.definition.Split(".")[-1]}          # group the results according to the extension

# get command details according to alias by get-command
Get-Command dir
Get-Command md

Get-Command dir | Format-List *                                                                     # get  all properties 

# Note: Basically, Get-Command is your primary discovery tool, helping you find the command you are looking for

###########################################################################################################################################################################################

# Tip 12: PowerShell Essentials: Get-Help

Get-Help dir                                                     # get a quick overview for the given command
Get-Help dir -Detailed                                           # get details info for the given command
Get-Help dir -Parameter *                                        # get all parameters for the given command
Get-Help dir -Examples                                           # get examples for the given command
Get-Help print                                                   # Get-Help can even guess so if you'd like to do something with printers
Get-Help service                                                 # If you enter a search phrase that applies to more than one cmdlet, you get back a list of all found cmdlets

# Note: Since sometimes Get-Help provides a lot more information than fits on one screen page, you can either pipe the result to more.com
Get-Help dir -Full | more.com
help dir -Full                                                   # use the pre-defined function help, which does pagination automatically

###########################################################################################################################################################################################

# Tip 13: PowerShell Essentials: Get-Member

dir | Get-Member                                                 # It lists all of the object and property members. By default, you get everything
                                                                 
$result = dir | Get-Member                                       
Compare-Object $result[0] $result[1] -Property Name              #  Get-Member analyzes the returned objects and returns two sets of data, one for the FileInfo and one for DirectoryInfo
                                                                 
dir | Get-Member *time* -MemberType *Property                    # to find out what properties are available related to "time"

# Note: Often, Get-Member provides a lot of information so it is a smart idea to limit it. 

###########################################################################################################################################################################################

# Tip 14: Ejecting CDs

$cdDrive = "D:"

$shell = New-Object -ComObject Shell.Application
$shell.NameSpace(17).parseName($cdDrive)
$shell.NameSpace(17).parseName($cdDrive).InvokeVerb("Eject")              # eject the CD drive

$shell.NameSpace(17).parseName($cdDrive).InvokeVerb("Properties")         # To open the properties of your CD drive

# Note: Note: Property names can be hard to guess because they may be localized. If you want to see all available context menu verbs, try this
$shell.NameSpace(17).parseName($cdDrive).Verbs()

###########################################################################################################################################################################################

# Tip 15: Passing ByRef vs. ByVal

# Usually, when you assign a variable to another variable, its content is copied.
$a = "Hello"
$b = $a
$a = "Hello World"
$b                                          # Output: Hello

# pass a pointer[ref] to a variable, effectively having two variables use the very same memory to store its values, change one the another one will be changed too
$a = "Hello"
$b = [ref]$a
$a = "Hello World"
$b                                          # Output: Hello World
$b.gettype()                                # PSReference`1
$b.Value

# Likewise, to change the $a variable through $b, you should assign a new value to the value property found in $b
$b.Value = "hello silence"
$a

###########################################################################################################################################################################################

# Tip 16: Splitting Text Into Words

$text = Get-Content C:\Users\v-sihe\Desktop\Temp\1.txt | Out-String

# Note: Out-String has one major disadvantage as it uses a fixed maximum line width so words may be truncated. A better approach is Join, found in the .NET String class
$text = [String]::Join(" ", (Get-Content C:\Users\v-sihe\Desktop\Temp\1.txt))

$word = $text.Split(" `t=", [stringsplitoptions]::RemoveEmptyEntries)

# Note: Above statement would use a space, a tab or an equal character to identify word boundaries and remove empty entries. However, this approach is not very dependable 
 # because there are a lot more non-word characters to handle. You should try a better approach of using regular expressions for splitting 
[regex]::Split($text, "[\s,\.]") | Where-Object {$_ -like "a*"} | Group-Object | Sort-Object {$_.name.length} -Descending

# Note: Here, any white space character, comma or dot is used to separate words. Still, this approach is not perfect. Therefore, a much better approach leaves it to 
 # regular expressions to identify word boundaries. Use Matches() instead of Split() to match explicit instances of words (\w+) separated by word boundaries (\b):
[regex]::Matches($text, "\b\w+\b") | ForEach-Object {$_.value} | Group-Object | Sort-Object count -Descending | Select-Object -First 10

###########################################################################################################################################################################################

# Tip 17: Displaying First Or Last Elements

# Select-Object can limit results to only the first or last elements. Simply use -first or -last
dir | Select-Object -First 10
Get-Process | Sort-Object CPU -Descending | Select-Object -Last 10
Get-Process | Sort-Object CPU -Descending | Select-Object -First 10 | Format-Table name,cpu

$a = Get-Process | Sort-Object CPU -Descending
$a[0..9]                                                    # get first 10 elements
$a[-1..-10]                                                 # get last 10 elements

###########################################################################################################################################################################################

# Tip 18: Sort With PS Code

dir | Sort-Object name
dir | Sort-Object length

dir | Sort-Object {$_.Name.Length} -Descending              # sorts your directory listing by length of name

###########################################################################################################################################################################################

# Tip 19: Create Text Reports with Format-Table

dir $env:windir | Group-Object extension

# a report of all the different file types in a folder
dir $env:windir | Group-Object extension -NoElement | Where-Object {$_.Name -ne ""} | Sort-Object Count -Descending | Format-Table -HideTableHeaders Count, {"Files of type $($_.Name)"}

"Currently running:";Get-Process | Group-Object Company -NoElement | Where-Object {$_.Name -ne ""} | 
    Sort-Object count -Descending | Format-Table -HideTableHeaders {"$($_.count) programs made by $($_.Name)."}

# convert this into a text file report
Get-Process | Group-Object Company -noElement | Where-Object { $_.Name -ne ''} | Sort-Object Count -descending | 
    Format-Table -hideTableHeaders { "$($_.Count) Programs made by '$($_.Name)'." } | Out-File $home\report.txt

# Note: You should be aware that Format-* cmdlets must always be the last elements in a pipeline because they convert the pipeline objects into console formatting objects
Get-Process | Group-Object Company -noElement | Where-Object { $_.Name -ne ''} | Sort-Object Count -descending | 
    Select-Object { "$($_.Count) Programs made by '$($_.Name)'." } | ConvertTo-Html | Out-File $home\report.htm

###########################################################################################################################################################################################

# Tip 20: Expanding Group Results

# returns all running software grouped by company name and sorted by frequency
Get-Process | Group-Object Company | Where-Object {$_.Name -ne ""} | Sort-Object Count -Descending | Format-Table -AutoSize

# group-object may returned result like array, use -expandproperty to expand the array and result result as a list
Get-Process | Group-Object Company | Where-Object {$_.Name -ne ""} | Sort-Object Count -Descending | Select-Object -ExpandProperty Group | Format-Table Company,Name

# OR 
Get-Process | Where-Object {$_.Company -ne $null} | Sort-Object Company -Descending | Format-Table Company,Name

# The differents between above two statements:
 # with Group-Object, you get the opportunity to sort on frequency, outputting the company with the most used software first, whereas a simple sort could have only sorted alphabetically

###########################################################################################################################################################################################

# Tip 21: Combining PowerShell And VBScript

& {
       '
       value = inputbox("please input text")
       WScript.Echo value
       '
  } | Out-File $HOME\test.vbs

  cscript.exe $HOME\test.vbs                                                 # call vbs script from PowerShell 

  $result = cscript.exe $HOME\test.vbs                                       # save the result from your VBScript into a variable

  cscript.exe $HOME\test.vbs | ForEach-Object {"processing input $_!"}

 # cscript //logo c:\"sample scripts"\chart.vbs
 # cscript //nologo c:\"sample scripts"\chart.vbs

###########################################################################################################################################################################################

# Tip 22: Cloning Objects

$a = 1..10
$b = $a
$a += 11
$b                      # After updated $a, $b not updated.

$a = @{}
$a.Test = 1
$a.Value = 2
$b = $a
$a.New = 3
$b                      # $a and $b has the same reference, so if one changed the another will be changed too

$a = @{}
$a.Test = 1
$a.Value = 2
$b = $a.Clone()
$a.New = 3
$b                      # $b is a copy for $a, they didn't share the same reference, so updated one will not affect another

###########################################################################################################################################################################################

# Tip 23: Finding CD-ROM Drives

function Get-CDDrives
{
    @(Get-WmiObject Win32_LogicalDisk -Filter "DriveType=5") | ForEach-Object {$_.DeviceID}
}

(Get-CDDrives).Count
(Get-CDDrives).Count -gt 0

# If you want to exclude CD-ROM drives with no media inserted, check the Access property. It is 0 for no media, 1 for read access, 2 for write access and 3 for both
Get-WmiObject Win32_LogicalDisk -Filter "DriveType=5 and Access>0"

###########################################################################################################################################################################################

# Tip 24: Renaming Object Properties

dir | Select-Object @{Name="FileName"; Expression = {$_.Name}}, Length                            # to rename e a directory listing "Name" property to "FileName"

dir | Select-Object @{Name="KB"; Expression = {"{0:0.0} KB" -f ($_.Length / 1KB)}}, Name, Length

###########################################################################################################################################################################################

# Tip 25: Creating New Shares

md $HOME\testFolder
([wmiclass]"Win32_Share").Create("$HOME\testFolder","myShare",0,5,"A new share").RetrunValue      # create a share for c:\testfolder called myShare and allow 5 connections

# Note: Note that you may need admin privileges for this or else the return value of 2 will indicate "access denied" while a return code of 0 will indicate success

# Deleting Shares:
(Get-WmiObject Win32_Share -Filter "Name='myShare'").Delete()

###########################################################################################################################################################################################

# Tip 26: Working Remotely With WMI

Get-WmiObject Win32_BIOS -ComputerName IIS-CTI5052
Get-WmiObject Win32_BIOS -ComputerName IIS-CTI5052 -Credential (Get-Credential)

# Note: If you receive a "RPC server not available" exception, then the system is blocked by a firewall or not online. If you receive an "access denied" exception, 
 # then you do not have local admin rights on the target machine. You can however authenticate yourself with different credentials

###########################################################################################################################################################################################

# Tip 27: List All Group Memberships of Current User

([System.Security.Principal.WindowsIdentity]::GetCurrent()).Groups | ForEach-Object {$_.Translate([System.Security.Principal.NTAccount])}

# on non-en OS, the "Administrators" localized is different, but they have the same security indentifier id, we can tranlate it via this ID
$lag=((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')).Translate( [System.Security.Principal.NTAccount]).Value.Split('\')[1])   # get localized "Admnistrator"
net localgroup $lag redmond\fwtlaba /add                                                                                                                 # add member to administrators

###########################################################################################################################################################################################

# Tip 28: List Local Groups

# to retrieve all Groups where the domain part is equal to your local computer name and returns the Group name and SID
Get-WmiObject Win32_Group -Filter "domain='$env:computername'" | Select-Object Name, SID

###########################################################################################################################################################################################

# Tip 29: Returning Exit Code from Script

# When running a PowerShell script, you may want to return a numeric exit code to the caller to indicate failure or success. 
 # You should use the "exit" statement to return whatever numeric exit code you want.
exit 10

#You will only receive either 0 or 1 when you now check the exit code from within your batch file using the %ERRORLEVEL% environment variable
powershell.exe -noprofile C:\path_to_script\script.ps1

# To actually pass on the return code submitted by your script, you will have to explicitly read it from $LASTEXITCODE and exit PowerShell with this code
powershell.exe -noprofile C:\path_to_script\script.ps1; exit $LASTEXITCODE

###########################################################################################################################################################################################

# Tip 30: Secret -Force Switch

Get-WmiObject Win32_BIOS
Get-WmiObject Win32_BIOS | Format-List *

$Error[0]
$Error[0].Exception                                       #　the strange thing happens when you try and output the error record Exception property
$Error[0].Exception | Format-List *　　　　　　　　　　　　　 #　For some reason, you only get back the error text message. Things do not change when you pass on the result to Format-List
$Error[0].Exception | Format-List * -Force　　　　　　　　　 #　In these situations, you can use the -Force switch to bypass the type system

###########################################################################################################################################################################################

# Tip 31: Append Information to Text Files

Add-Content $home\silence.txt ("{0:dd MM yyyy HH:mm:ss} New Entry" -f (Get-Date))
& "$home\silence.txt"

Add-Content $home\silence.txt ("{0:dd MM yyyy HH:mm:ss} New Entry" -f (Get-Date)) -PassThru   # Use the -passThru parameter if you'd like to see what exactly Add-Content is adding

Add-Content $home\silence.txt ("{0:dd MM yyyy HH:mm:ss} New Entry" -f (Get-Date)) -WhatIf     # use -whatif to see what will happen if you execute the command, just for simulate 

###########################################################################################################################################################################################

# Tip 32: Deleting

Get-Command del                 # Remove-Item

$test = "test"
del variable:test

function test {Get-Command}
del function:test

Set-Alias -Name edit -Value notepad.exe
del alias:edit

del HKCU:\test

# Note: You should remember to use the -force parameter if the item you want to delete is write-protected

###########################################################################################################################################################################################

# Tip 33: Getting Real Paths

New-PSDrive test FileSystem $env:windir
dir test:

Convert-Path test:\System32                # To "translate" virtual paths to real paths, use Convert-Path

###########################################################################################################################################################################################

# Tip 34: Launching Files

# The most important rule: always specify either an absolute or relative file path to whatever you want to launch 
 # - except if the file is located in a folder listed in the PATH environment variable

$env:Path                                    # get all environment path

calc                                         # start calculator

cd "$env:programfiles\Internet Explorer"
.\iexplore.exe                               # start explorer

# Launching Files without Specifying a Path:
$env:Path += "；C:\"

###########################################################################################################################################################################################

# Tip 35: Launching Files with Spaces

# Note: Use single quotes unless you want to resolve variables that are part of your path name. Then, you should use double quotes

& "$env:programfiles\Internet Explorer\iexplore.exe"     # whenever a path contains spaces, you need to quote it

# OR:
. "$env:programfiles\Internet Explorer\iexplore.exe" 

# Note: To invoke the string, add "&" or "." 

###########################################################################################################################################################################################

# Tip 36: Launching Files with Arguments

cd "$env:programfiles\Internet Explorer"
.\iexplore.exe www.powershell.com

# OR:
& "$env:programfiles\Internet Explorer\iexplore.exe" www.powershell.com

###########################################################################################################################################################################################

# Tip 37: Calling PowerShell from other scripts

powershell.exe -noprofile -command "path to .ps1 script file"

# Note: This is necessary because .ps1 files are by default not associated with powershell.exe. The parameter -noprofile skips your PowerShell profiles, which increases startup time 
 # and decreases memory footprint. Omit -noprofile only if you load important things in your profiles that your script depends on, such as loading additional snap-ins

###########################################################################################################################################################################################

# Tip 38: Interaction with PowerShell from other scripts

<# One:  get result from powershell via VBScript:

    pscommand = "$rv = read-host 'Enter return value (numeric value)'; exit $rv"
    cmd = "powershell.exe -noprofile -command " & pscommand
    Set shell = CreateObject("WScript.Shell")
    rv = shell.Run(cmd, , True)
    MsgBox "PowerShell returned: " & rv, vbSystemModal

    # Note: Make sure you use Run() synchronously by specifying True as third argument to read the value PowerShell returned. 
     # This way, Run() waits for the PowerShell command to complete and returns the value PowerShell has handed over to EXIT
#>


<# Two:  Returning Text Information From PowerShell To VBScript
    
    pscommand = "get-command | Foreach-Object {$_.Name}"
    cmd = "powershell.exe -noprofile -command " & pscommand
    Set shell = CreateObject("WScript.Shell")
    Set executor = shell.Exec(cmd)
    executor.StdIn.Close
    MsgBox executor.StdOut.ReadAll
#>

###########################################################################################################################################################################################

#　Tip 39: Calling VBScript From PowerShell

& {
        '
        answer = MsgBox("Do you use powershell yet?", vbYesNo + vbQuestion, "Survey")
        if answer = vbYes then
            WScript.Quit 1
        elseif answer = vbNo then
            WScript.Quit 2
        end if
        '
} | Out-File $HOME\test.vbs                             # in .vbs file, should use "" instead of ''

cscript.exe $HOME\test.vbs                              #　Use cscript.exe to call .vbs file in powershell

switch($LASTEXITCODE)
{
    1 {"That's great!"}
    2 {"It is worth it though!"}
    default {"Unexpected retrun value ..."}
}

# Note: The important part is to call the VBScript explicitly using cscript.exe and not the default wscript.exe. Only the console-based script host cscript.exe can process return values. 
 # Once PowerShell has called the VBScript, it can then retrieve its return value in the automatic variable $LASTEXITCODE. This is true for any external application or script 
 # PowerShell calls and is the equivalent to ERRORLEVEL in batch files.

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

&{
    '
    name = InputBox("Your Name?")
    WScript.Echo name
    '
} | Out-File $HOME\test2.vbs

# wscript //H:CScript                        # set the default script host to "cscript.exe
# WScript //H:WScript                        # set the default script host to "wscript.exe"

$name = cscript.exe $HOME\test2.vbs
"Hello $name"

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Note: When PowerShell calls this script using cscript.exe, the PowerShell pipeline keeps processing the names returned by VBScript until the VBScript is canceled
cscript.exe $HOME\test2.vbs | ForEach-Object {"processing '$_' ..."}

###########################################################################################################################################################################################

# Tip 40: Using String Functions

$str = "Hello"
$str.ToLower()                       # hello
$str.ToUpper()                       # HELLO
$str.EndsWith("lo")                  # True
$str.StartsWith("he")                # False 
$str.ToLower().StartsWith("he")      # True
$str.Contains("l")                   # True
$str.LastIndexOf("l")                # 3
$str.IndexOf("l")                    # 2
$str.Substring(3)                    # lo
$str.Substring(3,1)                  # l
$str.Insert(3,"INSERTED")            # HelINSERTEDlo
$str.Length                          # 5
$str.Replace("l","x")                # Hexxo

$str = "Server1,Server2,Server3"
$str.Split(",")                      # array: Server1   Server2  Server3

$str = "  remove space at ends  "
$str.Trim()                          # "remove space at ends"  
$str.Trim(" rem")                    # "ove space at ends"            

###########################################################################################################################################################################################

# Tip 41: Casting Strings

# Note: Strings represent text information and consist of individual characters. By casting, you can convert strings to individual characters and these into numeric ASCII codes

[char[]]"Hello"                      # array: H  e  l  l  o

[int[]][char[]]"Hello"               # array: "Hello" -> char[] -> int[]     72  101  108  108  111

[int]"H"                             # Exception here: PowerShell cannot directly convert a string into an ascii value

[int][char]"H"                       # 72, cast to a single character and then to an integer. The result is the ascii code

[char]72                             # convert a numeric value to a character

###########################################################################################################################################################################################

# Tip 42: Using VB.NET to Migrate From VBScript

# Simply load the VB.NET assembly that ships with .NET since it comes with almost all of your beloved VBScript commands
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null

[Microsoft.VisualBasic.Interaction] | Get-Member -Static                                 # get a list of all the commands this .NET class provides

[Microsoft.VisualBasic.Interaction]::MsgBox("Isn't that cool?", 1+32, "Survey")        

###########################################################################################################################################################################################

# Tip 43: Using Path Functions

# One: Using Simple Path Functions
Split-Path $HOME\test.txt                                                                    # C:\Users\v-sihe
Split-Path $HOME\test.txt -IsAbsolute                                                        # True
                                                                                             
Split-Path $HOME\test.txt -Leaf                                                              # test.txt
Split-Path $HOME\test.txt -Parent                                                            # C:\Users\v-sihe
                                                                                             
Split-Path $HOME\test.txt -NoQualifier                                                       # \Users\v-sihe\test.txt
Split-Path $HOME\test.txt -Qualifier                                                         # C:
                                                                                             

                                                                                             
# Two: Using Advanced Path Functions                                                         
[System.IO.Path] | Get-Member -Static                                                        # Use Get-Member with the -static parameter to list its members
[System.IO.Path]::ChangeExtension("$HOME\test.vbs", "ps1")                                   # the real file extension is not changed 
                                                                                             
 
                                                                                             
# Three: Checking Paths for Invalid Path Characters                                          
[System.IO.Path]::GetInvalidPathChars()                                                      # find out which characters are considered to be illegal in a path
 
$pattern = "[{0}]" -f ([regex]::Escape([String][System.IO.Path]::GetInvalidPathChars()))     # use regex to match a path is invalid or not

# Now, it is easy to detect whether a path contains any of the illegal characters
"c:\test<>" -match $pattern                                                                  # True

$warning = "Enter a path"
do
{
    $path = Read-Host $warning
    $warning = "Path was illegal. try again!"

}while($path -match $pattern)



# Four: Checking File Names for Invalid Characters
$pattern2 = "[{0}]" -f ([regex]::Escape([String][System.IO.Path]::GetInvalidFileNameChars()))
$warning2 = "Enter a file name"
do
{
    $path = Read-Host $warning2
    $warning2 = "This file name contained '$($Matches[0])'" + "which is illegal in a file name, Try again!"

}while($path -match $pattern2)

###########################################################################################################################################################################################

# Tip 44: Using Test-Path to Validate A Path

do
{
    $path = Read-Host "Enter a path"

}while(Test-Path -IsValid $path)

###########################################################################################################################################################################################

# Tip 45: Removing Illegal Characters

# One: Removing Illegal Path Characters
$path = "C:\<>illegal path"
$path
$pattern = "[{0}]" -f ([regex]::Escape([String][System.IO.Path]::GetInvalidPathChars())) 
$newpath = [regex]::Replace($path,$pattern,"")
$newpath



# Two: Removing Illegal File Name Characters
$file = "this*file\\is_illegal<>.txt"
$file
$pattern2 = "[{0}]" -f ([Regex]::Escape([String][System.IO.Path]::GetInvalidFileNameChars()))   
$newfile = [regex]::Replace($file,$pattern2,"")
$newfile

###########################################################################################################################################################################################

# Tip 46: Finding Alias Names

Get-Alias | Where-Object {$_.Definition -eq "Get-Childitem"}

dir Alias: | Group-Object Definition | Sort-Object Count -Descending           # get a comprehensive list of all Cmdlets and their aliases

Get-Alias -Definition Get-ChildItem                                            # In PowerShell V2, it is much simpler to find alias for a command

###########################################################################################################################################################################################

# Tip 47: Extending Alias Functionality

Set-Alias ie "$env:programfiles\Internet Explorer\iexplore.exe"
ie www.baidu.com

function pingFast {ping.exe $args -n 1 -w 500}
pingFast 127.0.0.1

###########################################################################################################################################################################################

# Tip 48: Conflicting Commands

# Note: PowerShell supports many different command categories and searches for the command in the following order:
    # 1. Alias
    # 2. Function
    # 3. Cmdlet
    # 4. Executable
    # 5. Script
    # 6. Associated Files

Set-Alias ping notepad.exe
ping localhost                  # ping equal notepad.exe here, so it can not be used like ping.exe

ping.exe localhost              # ping.exe can work like cmd ping here

Get-Command -CommandType cmdlet,function,alias | Group-Object name | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Group }          # To list all conflicting commands

###########################################################################################################################################################################################

# Tip 49: Using Switch Parameters

# Switch parameters work like a switch, so they are either "on" or "off" aka $true or $false.
function test([switch]$force)
{
    $force
}

test                                  # False
test -force                           # True


# If you need the opposite result and want $force to be $true when it is omitted, simply turn the result around
function test([switch]$force)
{
    $force = -not $force
    $force
}

test                                  # True
test -force                           # False

###########################################################################################################################################################################################

# Tip 50: Avoid Format-... in Scripts

# Below code will get exception here
&{
    "
    Get-Process
    dir | Format-Table Name
    "

} | Out-File $HOME\test.ps1

&"$HOME\test.ps1"

# Reason: get-process use default formatter, the second statement use format-table formatter, powershell can not mix them

# Solution: to make sure all cmdlets use the same formatter
&{
    "
    Get-Process | Format-Table Name
    dir | Format-Table Name
    "

} | Out-File $HOME\test.ps1

# Note: As a general rule of thumb, in scripts it is often better to replace formatter cmdlets with select-object
Get-Process | Select-Object Name

###########################################################################################################################################################################################

# Tip 51: Assigning Values to Parameters

function test([int]$number = 0, [switch]$force)
{
    "You specified: number = $number, force = $force"
}

test                                                                            # You specified: number = 0, force = False
                                                                                
test 100 $true                                                                  # You specified: number = 100, force = False
test 100 -force                                                                 # You specified: number = 100, force = True
                                                                                                                                                               
test -number 100 -force                                                         
test -force -number 100                                                         
                                                                                
test -number 100                                                                # You specified: number = 100, force = False
test -number:100                                                                # You specified: number = 100, force = False

# switch parameters do not accept a value, the following call would fail
test -force $false                                                              # You specified: number = 0, force = True
                                                                                
# To bind a value to a switch parameter                                         
test -force:$false                                                              # You specified: number = 0, force = False 
     
###########################################################################################################################################################################################

# Tip 52: Virtual Drives With UNC-Paths

New-PSDrive home FileSystem $HOME                                               # to create a drive that points to your user profile
dir home:\

New-PSDrive HKCR Registry HKEY_CLASSES_ROOT                                     # add a new drive pointing to HKEY_CLASSES_ROOT
dir HKCR:\

New-PSDrive NetDrive FileSystem \\127.0.0.1\c$                                  # To map a UNC path to a virtual drive
dir NetDrive:

###########################################################################################################################################################################################

# Tip 53: Adding Multiple Registry Keys

New-Item HKCU:\Software\TestKey                                                 # to create new registry keys

New-Item HKCU:\Software\TestKey\A\B\C\D                                         # Exception here: New-Item can only create one key at a time and fails if the parent key does not exist

New-Item HKCU:\Software\TestKey\A\B\C\D -Force                                  # specify the -force switch, New-Item will happily create even multiple keys

Remove-Item HKCU:\Software\TestKey -Force

###########################################################################################################################################################################################

# Tip 54: Summing Up Multiple Objects

Get-WmiObject Win32_Battery                                                     # try and figure out the battery charge on a notebook

# To find out the overall charge remaining in all of your batteries
Get-WmiObject Win32_Battery | Measure-Object EstimatedChargeRemaining -Average (Get-WmiObject Win32_Battery | Measure-Object EstimatedChargeRemaining -Average).Average

# Note that this call will fail if there is no battery at all in your system because then Measure-Object is unable to find a EstimatedchargeRemaining property

###########################################################################################################################################################################################

# Tip 55: Displaying Battery Charge in your prompt

function prompt
{
    $charge = (Get-WmiObject Win32_Battery | Measure-Object -Property EstimatedChargeRemaining -Average).Average
    $prompt = "PS [$charge%] >"

    if($charge -lt 30)
    {
        Write-Host $prompt -ForegroundColor Red -NoNewline
    }
    elseif($charge -lt 60)
    {
        Write-Host $prompt -ForegroundColor Yellow -NoNewline
    }
    else
    {
        Write-Host $prompt -ForegroundColor Green -NoNewline
    }

    $host.UI.RawUI.WindowTitle = (Get-Location)
    " "
}

###########################################################################################################################################################################################

# Tip 56: Consolidating Information In An Object

$result = 1 | Select-Object Version, Build
$result.Version = (Get-WmiObject Win32_BIOS).Version
$result.Build = (Get-WmiObject Win32_OperatingSystem).BuildNumber
$result

###########################################################################################################################################################################################

# Tip 57: Wait For Key Press

Read-Host "Please press Enter" | Out-Null                             # The easiest way is to use Read-Host

# Note:  Read-Host accepts more than one key and only continues after you press ENTER. To implement a real "one key press only" solution, 
 # you need to access the PowerShell raw UI interface. This interface has a property called KeyAvailable which returns $true once a key was entered

function Wait-KeyPress($prompt = "Press a key!")
{
    Write-Host $prompt

    do
    {
        Start-Sleep -Milliseconds 100    
    }
    until($host.UI.RawUI.KeyAvailable)                                # A loop queries KeyAvailable until it returns $true

    $host.UI.RawUI.FlushInputBuffer()                                 # clears the input buffer and removes the pressed key,to prevent the entered key to show up in subsequent inputs 
}

Wait-KeyPress

# There is an easy way to do this
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")                  # this command will got exception in ISE, but works well in powershell console
$x                                                                    # output the details info for the key which pressed

###########################################################################################################################################################################################

# Tip 58: Using Select-String to Focus on the important stuff

ipconfig | Select-String IP
 # Windows IP Configuration
 #   Link-local IPv6 Address . . . . . : fe80::70cb:39dc:abc:253c%11
 #   IPv4 Address. . . . . . . . . . . : 172.18.32.41

route print | Select-String 127.0.0.1

# Note that select-string compares text case-insensitive unless you specify the -CaseSensitive switch parameter
$result  = whoami /priv | Select-String Disable
$result2 = whoami /priv | Select-String Disable -CaseSensitive
Compare-Object $result $result2

###########################################################################################################################################################################################

# Tip 59: "Grep": Finding PowerShell Scripts with a Keyword

$path = C:\Users\v-sihe\Desktop\Tools\PowershellScripts\*.ps1

Select-String -Path $path  Get-WmiObject -List                                # finds all *.ps1 files in your Documents folder containing Get-WMIObject
# Note the use of -list which lists each file only once. Without this switch parameter, you would get one result for each found keyword

Select-String -Path $path  Get-WmiObject -List | Format-Table Path            # select-string returns rich objects which you can format

###########################################################################################################################################################################################

# Tip 60: Advanced String Filtering

route print | Select-String 127.0.0.1                                         # Select-String does not return the filtered text information but instead wraps this as MatchInfo objects

route print | Select-String 127.0.0.1 | Format-Table *                        

(route print) -like "*127.0.0.1*"                                                                                    # A much better way of filtering string arrays is the -like operator

(route print) -like "*127.0.0.1*" | ForEach-Object {$_.Split(" ")}

(route print) -like "*127.0.0.1*" | ForEach-Object {$_.Split(" ",[StringSplitOptions]"RemoveEmptyEntries")}          # Split() supports an advanced option to remove empty entries

(route print) -like '*127.0.0.1*' | ForEach-Object {$_.Split(' ',[StringSplitOptions]'RemoveEmptyEntries')[-1]}  # Using an index, you specify which array element you'd like to use

###########################################################################################################################################################################################

# Tip 61: Feeding Input Into Native Commands

"list disk" | DiskPart                                            # to get a list of drives without interaction

"list disk" | DiskPart | Where-Object {$_.StartsWith(" ")}        # process the results with Where-Object and filter out any unwanted information

@'
select disk 0
detail disk
'@ | DiskPart                                                     # you can even submit more than one command

# Note: Be aware that DiskPart requires Admin privileges

###########################################################################################################################################################################################

# Tip 62: Finding Out A Drives' FileSystem

function Get-FileSystem($letter = "C:")
{
    if(!(Test-Path $letter))
    {
        Throw "Drive $letter does not exist"
    }

    ([wmi]"Win32_LogicalDisk='$letter'").FileSystem
}

Get-FileSystem C:                                               # NTFS
Get-FileSystem D:                                               # Drive D: does not exist



function Get-FileSystem($letter = "C:")
{
    if(!(Test-Path $letter))
    {
        Throw "Drive $letter does not exist"
    }

    [wmi]"Win32_LogicalDisk='$letter'"                          # Check out what you get when you omit ".FileSystem"! Pipe the result to Format-List * to see all available information
}

Get-FileSystem C: | Format-List *

###########################################################################################################################################################################################

# Tip 63: Reading and Writing Drive Labels

# To read the existing drive label
function Get-DriveLabel($letter = "C:")
{
    if(!(Test-Path $letter))
    {
        Thorw "Drive $letter does not exist."
    }

    ([wmi]"Win32_LogicalDisk='$letter'").VolumeName
}

# To actually change the drive label
function Set-DriveLabel($letter = "C:", $label = "New Label")
{
    if(!(Test-Path $letter))
    {
        Thorw "Drive $letter does not exist."
    }

    $instance = ([wmi]"Win32_LogicalDisk='$letter'")
    $instance.VolumeName = $label
    $instance.Put()
}

# Note: Be aware that changing a drive label requires Admin privileges

Get-DriveLabel
###########################################################################################################################################################################################

# Tip 64: Converting FileSystem To NTFS

function ConvertTo-NTFS($letter = "C:")
{
    if(!(Test-Path $letter))
    {
        Throw "Drive $letter does not exist."
    }

    $drive = [wmi]"Win32_LogicalDisk='$letter'"
    $label = $drive.VolumeName
    $filesystem = $drive.FileSystem

    if($filesystem -eq "NTFS")
    {
        Throw "Drive already uses NTFS filesystem"
    }

    "Label is $label"
    
    $label | convert.exe $letter /FS:NTFS /X             # PowerShell can embrace convert.exe and turn it into a more sophisticated tool that converts drives without manual confirmation
                                                         # the convert.exe utility which can convert a file system to NTFS without data loss
}

# Note: Make sure the drive you want to convert is not in use or else the conversion may be scheduled to take place at next reboot.

###########################################################################################################################################################################################

# Tip 65: List Folders or Files

# One: List Hidden Files
dir $env:SystemDrive -Force                                                                                         # To see hidden files, you need to specify the -force parameter
dir $env:SystemDrive -Force | Where-Object {$_.Mode -like "*h*"}                                                    # just get the hidden files


# Two: List All Folders and Subfolders
dir -Recurse | Where-Object {$_.PSIsContainer} | ForEach-Object {$_.FullName}


# Three: Finding Empty Folders
dir $env:windir -Recurse | Where-Object {$_.PSIsContainer} | Where-Object {$_.GetFiles().Count -eq 0} | ForEach-Object {$_.FullName}      # To find out all folders that contain no files

dir $env:windir -Recurse | Where-Object {$_.PSIsContainer} | Where-Object {$_.GetFiles().Count -eq 0} | 
    Where-Object {$_.GetDirectories().Count -eq 0} | ForEach-Object {$_.FullName}                                   #  to create a list of completely empty folders (no subfolders either)

###########################################################################################################################################################################################

# Tip 66: Listing All Installed Font Families

[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
$families = (New-Object System.Drawing.Text.InstalledFontCollection).Families

$families -contains "Wingdings"                                                                                     # to check whether a specific font family is available

###########################################################################################################################################################################################

# Tip 67: Converting Objects Into Text

[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
$families = (New-Object System.Drawing.Text.InstalledFontCollection).Families

$families | Select-String "Wingdings"                                                                               # output: [FontFamily: Name=Wingdings] ...
$families[0].GetType().FullName                                                                                     # output: System.Drawing.FontFamily
                                                                                                                    
# convert the objects into plain text strings to then process them with Select-String                               
$families | Out-String -Stream | Select-String "Wingdings"                                                          # output: Wingdings ...   
                                                                                                                    
# As it turns out, FontFamily objects only have one property called Name which is a string                          
$families[0] | Get-Member -MemberType *property                                                                     
                                                                                                                    
$families | ForEach-Object {$_.Name} | Select-String "Wingdings"                                                    #　use the Name property to find out the right font
$families | Where-Object {$_.Name -like "*Wingdings*"} | ForEach-Object {$_.Name}
                               
###########################################################################################################################################################################################

# Tip 68: Creating A HTML Font List

#  create a HTML document listing each installed type face on your computer 
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null

$result = (New-Object System.Drawing.Text.InstalledFontCollection).Families | ForEach-Object {$strHTML = ""}{
    $strHTML += "<font size='5' face='{0}'>{0}</font><br>" -f ($_.Name)
}{$strHTML}

Set-Content $HOME\test.htm $result

& "$HOME\test.htm"

###########################################################################################################################################################################################

# Tip 69: Test-Path Can Check More Than Files

# One: To check whether a file or folder exists
Test-Path C:\bootmgr                                               # True
Test-Path C:\notexist                                              # False


# Two: To check whether an environment variable exists
Test-Path Env:\COMPUTERNAME                                        # True


# Three: To check for an alias
Test-Path Alias:\dir                                               # True


# Four: To check for registry keys
Test-Path HKCU:\Software\AppDataLow                                # True


# Five: To check the path point to file or folder
Test-Path C:\bootmgr -PathType Leaf                                # True: point to file,   use -PathType Leaf
Test-Path 'C:\$Recycle.Bin' -PathType Container                    # True: point to folder, use -PathType Container

 # Note: If c:\test was a file (without extension), Test-Path would still return false because with the -pathtype parameter, you explicitly looked for folders. 
  # Likewise, if you want to explicitly check for files, use the parameter -pathtype leaf


# Six: To find specific file
Test-Path $HOME\*.ps1

###########################################################################################################################################################################################

# Tip 70: Does a Folder contain a specific file?

Test-Path $HOME\*.ps1                                              # Test-Path supports wildcards 

@(dir $HOME\*.ps1).Count                                           # To get the actual number of PowerShell scripts

# Note the @() converts any result into an array

###########################################################################################################################################################################################

# Tip 71: Storing Cmd-Results in PowerShell Variables

$result = cmd.exe /c dir
$result

# Note: You can run classic cmd commands from within PowerShell and store the results in variables. All you need to do is invoke cmd.exe with the /c switch 

###########################################################################################################################################################################################

# Tip 72: PowerShell ISE uses Unicode

# all scripts you create are saved in Unicode by default. This was done to support more languages, but you may no longer be able to open these scripts in other editors

# Solution: 
$psISE.CurrentFile.Save([Text.Encoding]::ASCII)

###########################################################################################################################################################################################

# Tip 73: Import-CSV and Types

dir $env:windir | Select-Object Name, Length | Export-Csv $HOME\test.csv

$result = Import-Csv $HOME\test.csv
$result

$result | Sort-Object Length                                                      # As it turns out, Export/Import-CSV does preserve properties but converts them all to string

###########################################################################################################################################################################################

# Tip 74: Persisting Objects with XML

dir $env:windir | Select-Object Name, Length | Export-Clixml $HOME\test.xml

$result = Import-Clixml $HOME\test.xml
$result

$result | Sort-Object Length                                                      # when use export-clixml instead of export-csv, the re-loaded objects are sorted correctly (numerically)

###########################################################################################################################################################################################

# Tip 75: CSVs with Alternative Delimiters

Get-Service | Export-Csv $HOME\test2.csv -Delimiter ";"                      
. "$Home\test2.csv"                                                               # export-csv use "," as default delimiter, it not work well if you use ";" on en OS

# Note: Provided you run this on a system with German locale (or any culture that uses ";" as default separator) 
 # and have excel installed, you will get a nice spreadsheet with all service details

# Solution: -UseCulture can use the right delimiter automaticlly to adjust to your OS culture
Get-Service | Export-Csv $HOME\test3.csv -UseCulture
. "$Home\test3.csv"

###########################################################################################################################################################################################

# Tip 76: Adding Custom Properties

$process = (Get-Process powershell)[0]
$process | Add-Member NoteProperty SayHello Hello
$process.SayHello                                                                  # Hello

# When you try the same with a different object such as a string, it fails
$object = "Hello"
$object | Add-Member NoteProperty SayHello Hello
$object.SayHello                                                                   # nothing output here

# Reason: The truth is PowerShell can add new properties only to its own object types (PSObject). So if the object you want to extend is not of the correct type, 
 # Add-Member cannot manipulate it. To solve this issue, use the -passThru parameter and update the variable. 
 # This way, Add-Member converts your object into the appropriate type, adds the property and updates your variable
$object = "Hello"
$object = $object | Add-Member NoteProperty SayHello Hello -PassThru
$object.SayHello                                                                   # Hello

# Note: 使用PowerShell 中的PassThru参数可以将那些新创建的或者经过更新的对象由默认的隐藏变成输出或返回，以便进行下一步操作，体现的正是PowerShell的灵活性

###########################################################################################################################################################################################

# Tip 77: Adding Custom Methods to Objects

$text = "This is some text"
$text = $text | Add-Member ScriptMethod Words {$this.Split()} -PassThru   # create an external function to manipulate text and to find out the word count in a string
$text.Words()
@($text.Words()).Count

###########################################################################################################################################################################################

# Tip 78: Accessing Current PowerShell Process

#Note: If you ever want to access the process that is executing your current PowerShell session, use the $pid automatic variable which tells you the Process ID, and feed it to Get-Process
Get-Process -Id $pid

###########################################################################################################################################################################################

# Tip 79: Count Your Work: Calculating Process Runtime

(Get-Process -id $pid).StartTime                                          # to know when you started your PowerShell session

$timespane = (New-TimeSpan (Get-Process -Id $pid).StartTime).TotalHours   # to find out how long you have been working
"You worked for $timespane hours now!"


# Display Work Hours in Prompt
function prompt
{
    $work = [int]((New-TimeSpan (Get-Process -Id $pid).StartTime).TotalMinutes)
    "$work min. PS> "
    $host.UI.RawUI.WindowTitle = (Get-Location)
}

###########################################################################################################################################################################################

# Tip 80: Getting System Uptime

# to determine a system's uptime, you should use WMI and convert the WMI date into a more readable format
function Get-SystemUpTime
{
    $os = Get-WmiObject Win32_OperatingSystem
    [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
}

# Note: You can even query remote systems by adding the -computerName parameter to Get-WmiObject. Use the shortcut ETS has created if you don't want to remember the .NET class
function Get-SystemUpTime
{
    $os = Get-WmiObject Win32_OperatingSystem
    $os.ConvertToDateTime($os.LastBootUpTime)
}

###########################################################################################################################################################################################

# Tip 81: Set PowerShell Execute Mode

Set-StrictMode -Version 1
Set-StrictMode -Version 2
Set-StrictMode -Version 3

# Note: Use the -version parameter to tell the cmdlet which PowerShell version rules you would like enforce. 
 # PowerShell will throw a lot of suggestions at you whenever you do something that violates the rules when you enter commands or run scripts in strict mode

###########################################################################################################################################################################################

# Tip 82: Processing Switch Return Value

function Get-Name ($number)
{
    switch($number)
    {
        1 {"One"}
        2 {"Two"}
        3 {"Three"}
        default {"Other"}
    }
}

Get-Name 1
Get-Name 100


function Get-Name($number)
{
    $result = $(switch ($number){        # you can assign the result from Switch to a variable directly by simply placing Switch into $()
        1 {"One"}                        # Note: In PowerShell V2, the $() workaround isn't even necessary
        2 {"Two"}
        3 {"Three"}
        default {"Other"}   
    })

    "The result is: $result"
}

Get-Name 1
Get-Name 100

###########################################################################################################################################################################################

# Tip 83: Some Special for Switch

# One: Use Switch to handle array
$nums = 7..10
switch($nums)
{
    Default {"n = $_"}
}
# Note: Switch 本是多路分支的关键字，但是在Powershell中由于Switch支持集合，所以也可以使用它进行循环处理


# Two: Use switch for condition filter
$nums = 7..10
switch($nums)
{
    {($_ % 2) -eq 0} {"$_ 偶数"}
    {($_ % 2) -ne 0} {"$_ 基数"}
}
# Note: 有时对集合的处理，在循环中还须条件判断，使用Switch循环可以一部到位


# Three: Compare String
$domain = "www.baidu.com"
switch -CaseSensitive ($domain)     # Switch有一个-case 选项，一旦指定了这个选项，比较运算符就会从-eq 切换到 -ceq，即大小写敏感比较字符串
{
    "WWW.Baidu.com" {"OK 1"}
    "www.baidu.COM" {"OK 2"}
    "www.baidu.com" {"OK 3"}
}                                   # output: OK 3
# Note: if no -CaseSensitive, it will output: OK 1   OK 2  OK 3


# Four: Use wildcard
$domain = "www.baidu.com"
switch -Wildcard ($domain)
{
    "*"     {"match '*'"}
    "*.com" {"match '*.com'"}
    "*.*.*" {"match '*.*.*'"}
}


# Five: Use regex
$mail = "hhb@hotmail.com"
switch -Regex ($mail)
{
    "^www"                         {"start with www"}
    "com$"                         {"end with com"}
    "d{1,3}.d{1,3}.d{1,3}.d{1,3}"  {"IP address"}
}                                                                 # output: end with com

###########################################################################################################################################################################################

# Tip 84: Using Hash Tables Instead of Switch

function Get-Name($number)
{
    $translate = @{
        1 = "One"
        2 = "Two"
        3 = "Three"
    }
    
    if($translate.ContainsKey($number))                           # Checking Whether Hash Table Contains Key
    {
        $translate[$number]
    }
    else
    {
        "Other"
    }
}

Get-Name 1
Get-Name 20

# Note: For translation jobs, a hash table could be more efficient than Switch

###########################################################################################################################################################################################

# Tip 85: Check PowerShell is Available or not

# To determine whether any PowerShell version is available, check whether this key is available by manual: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\1
Test-Path registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\1           # This command can be success only if you got powershell available

$PSVersionTable                                                                  # get detail info about the current powershell you used

###########################################################################################################################################################################################

# Tip 86: Trap and Try/Catch

trap
{
    Write-Host "Something terrible happened: $($_.Exception.Message)" -ForegroundColor Yellow

    continue                                     # if ignore continue here, powershell will throw internal exception message except your specified message
}

& {
      dir nonexisttent: -ErrorAction Stop
      "Hello"                                    # any additional lines in this scope will be ignored
  }
"I'm back!"


try
{
    dir nonexisttent: -ErrorAction Stop
}
catch
{
    Write-Host "Something strange occurred: $_" -ForegroundColor Yellow
}

###########################################################################################################################################################################################

# Tip 87: Creating Script Modules

# Note: Before you can use the functions defined in a module, you will need to load the module using import-module. 
 # Secondly, modules can determine which functions should be public and which should be hidden (for internal purposes)
&{
    '
    $internal = 1
    $external = 2

    function Test-Function1
    {
        "I am function 1"
    }
    
    function Test-Function2
    {
        "I am function 2"
    }

    Export-ModuleMember -Function Test-Function1 -Variable external     # export Test-Function1 and $external
    '
}  | Out-File $HOME\example.psm1                                        # You should save this file as example.psm1 and then import it 

Import-Module $HOME\example.psm1            
# Once complete, you can now call Test-Function1 and $external, but you cannot access Test-Function2 and $internal because they were not explicitly published

# Note: Once you have loaded a module, it is important to understand that it is cached until the PowerShell session ends. 
 # So even if you reload a module using Import-Module, you will still work with the module you loaded initially. 
 # While this is acceptable for performance reasons, it may be confusing during development. To actually see changes you made to a module, 
 # you need to re-import the module using Import-Module with the -force switch

###########################################################################################################################################################################################

# Tip 88: Managing PowerShell Modules

&{
    '
    $internal = 1
    $external = 2

    function Test-Function1
    {
        "I am function 1"
    }
    
    function Test-Function2
    {
        "I am function 2"
    }

    Export-ModuleMember -Function * -Variable *                 # export all function and variable in the .psm1 file
    '
}  | Out-File $HOME\example.psm1 

$modules = Import-Module $HOME\example.psm1 -Force -PassThru
$modules | Get-Member                                           # show you all the useful methods you can now access. For example, functions, aliases, variables, etc.

###########################################################################################################################################################################################

# Tip 89: Accessing Hidden Module Members

&{
    '
    $internal = 1
    $external = 2

    function Test-Function1
    {
        "I am function 1"
    }
    
    function Test-Function2
    {
        "I am function 2"
    }

    Export-ModuleMember -Function Test-Function1 -Variable external     # export Test-Function1 and $external
    '
}  | Out-File $HOME\example.psm1                                        # You should save this file as example.psm1 and then import it 

$modules =  Import-Module $HOME\example.psm1 -Force -PassThru

# Now, you can submit a script block to the module, which executes in its private context and has access to all members, regardless of what Export-ModuleMember has declared public
& $modules { "The hidden variable INTERNAL contains: $internal "}       # Test-Function2 and $internal can access now with this method

###########################################################################################################################################################################################

# Tip 90: Loading New Windows 7 Modules

Get-Module -ListAvailable                       # to see which modules are available

Import-Module BitsTransfer                      # to load a module you should use Import-Module and the module name you want to load

Get-Command -Module BitsTransfer                # to see which new cmdlets this gets you

# Note: Windows 7 comes with a number of such modules. Use the pre-defined function ImportSystemModules to load them all.
ImportSystemModules

###########################################################################################################################################################################################

# 91: Getting Process Windows Titles

Get-Process | Where-Object {$_.MainWindowTitle} | Format-Table id, name, mainwindowtitle -AutoSize            # filter all processes and remove those that have no window title

Get-Process | Where-Object {$_.MainWindowTitle -like "*tip*"}                                                 # get the process which main window title contains "tip"

###########################################################################################################################################################################################

# Tip 92: Handling Event Logs with Get-WinEvent

Get-WinEvent -ListLog *                                       # get you all the event logs 

Get-WinEvent -ListLog *powershell*                            # use keywords to find special logs

Get-WinEvent *powershell*                                     # To actually read event log entries from one or more event logs, simply remove the -listLog parameter


# Instead of searching for specific event logs, you can search for specific event providers to determine which event logs they maintain

Get-WinEvent -ListProvider *policy*                           # to find all event logs related to policies

Get-WinEvent -ListProvider *powershell*                       # to find all providers related to PowerShell

Get-WinEvent -ProviderName powershell                         # dump all event log entries created by a specific provider


# Finding Events Supported by an Event Provider
Get-WinEvent -ListProvider *powershell* | ForEach-Object {$_.Events} | Format-Table id, description -AutoSize

###########################################################################################################################################################################################

# Tip 93: Creating Your Own Eventlog

New-EventLog -LogName "Client Login Scripts" -Source "Logonscript"                                                        # create your own event log, this requires Admin privileges 

Write-EventLog -LogName "Client Login Scripts" -Source "Logonscript" -Message "Something bad happened" -EventId 111       #  start writing events to your event log

Get-WinEvent -ProviderName Logonscript                                                                                    # To read all logged events

Remove-EventLog -LogName "Client Login Scripts"                                                                           # remove event log
# Note: this cmdlet removes any event log you specify, including pre-existing system logs. 
 # You should be careful what you delete because it cannot be undone. You will need Admin privileges to safeguard users.

###########################################################################################################################################################################################

# Tip 94: Finding Unwanted Output

# Note: Functions will return more information than you planned if you forget to clean this up by sending it to Out-Null or casting the return value to [void]
function test
{
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

    "This is my return value"
}

function test
{
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null   # OR: [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

    "This is my return value"
}


$null = Set-PSBreakpoint -Command Out-Default                                                # find all the spots that return data so you can check who is responsible for unwanted output
# When you call your function again, the debugger will stop at each instance that returns data.

###########################################################################################################################################################################################

# Tip 95: Restarting Processes

# Note: when restart a process, you should wait a minutes because it will take time to shutdown the process (the process can be shutdown shortly)

Wait-Process notepad -ErrorAction SilentlyContinue; notepad                                                                # re-launch Notepad once you close it
                                                                                                                           
Wait-Process notepad -ErrorAction SilentlyContinue; Write-Host "the notepad process shutdown"                              # give text notification once a process has fully shut down
                                                                                                                           
Wait-Process notepad -ErrorAction SilentlyContinue                                                                         
(New-Object -ComObject SAPI.SpVoice).Speak("OK, the notepad is down!")                                                     # give voice notification once a process has fully shut down 

###########################################################################################################################################################################################

# Tip 96: Picking Random Items

# Note: Get-Random retrieves as many random numbers as you like, allowing you to use this random number generator in many scenarios
"Tip 1", "Tip 2", "Tip 3" | Get-Random

Get-Process | Get-Random -Count 2 | Stop-Process -WhatIf                                                                   # get two random processes and kill them if memory runs out

###########################################################################################################################################################################################

# Tip 97: Getting Image Details

$image = New-Object -ComObject WIA.ImageFile                                                                      # WIA.Image can return all kinds of useful information about images
$image.LoadFile("$([System.Environment]::GetFolderPath('MyPictures'))\10.jpg")                                    # to load an image file from "My Pictures" folder
$image                                                                                                            # show image details

###########################################################################################################################################################################################

# Tip 98: Exploring Cmdlets Added by Snap-ins

Get-Command -CommandType Cmdlet | Sort-Object PSSnapin, Module | Format-Table Name, PSSnapin, Module              #  to see which cmdlets are available and where they come from
# Note: cmdlets available in PowerShell can come from two sources: PowerShell-SnapIns and (new in v.2) modules


Get-PSSnapin                                                                                           # Get-PSSnapin lists all loaded snap-ins

Get-Command -pssnap *management                                                                        # to see which cmdlets are located in the snap-in Microsoft.PowerShell.Management
Get-Command -pssnap *                                                                                  # get all snap-in in current powershell console

###########################################################################################################################################################################################

# Tip 99: Sorting with Sort-Object

Get-Service | Sort-Object -Property Name | Format-Table Name, Status -AutoSize
Get-Service | Sort-Object -Property Name -Descending | Format-Table Name, Status -AutoSize

# sort with two properties at the same time: sort status with descending and sort name with default sort
Get-Service | Sort-Object -Property @{Expression = "Status"; Descending = $true}, @{Expression = "Name"; Descending = $false} | Format-Table Name, Status -AutoSize

###########################################################################################################################################################################################

# Tip 100: Listing Official PowerShell Verbs                              

Get-Verb                                                       # It lists all officially approved PowerShell verbs 

${function:get-verb}                                           # view source code for get-verb

function SourceCodeforGetVerb
{
    param
    (
        [Parameter(ValueFromPipeline=$true)]
        [string[]]
        $verb = '*'
    )
    
    begin 
    {
        $allVerbs = [PSObject].Assembly.GetTypes() | Where-Object {$_.Name -match '^Verbs.'} | Get-Member -type Properties -static |
            Select-Object @{Name='Verb'; Expression = {$_.Name}}, @{ Name='Group'; Expression = {$str = "$($_.TypeName)"; $str.Substring($str.LastIndexOf('Verbs') + 5)}}
    }

    process
    {
        foreach ($v in $verb) 
        {
            $allVerbs | Where-Object { $_.Verb -like $v }
        }
    }
    # .Link
    # http://go.microsoft.com/fwlink/?LinkID=160712
    # .ExternalHelp System.Management.Automation.dll-help.xml
}

###########################################################################################################################################################################################