# Reference site: http://powershell.com/cs/blogs/tips/
###########################################################################################################################################################################################

# Tip 1: Rename Drive Label

$drive = [WMI]"Win32_LogicalDisk='E:'"
$drive.VolumeName = "My HardDrive"
$drive.Put()                                                           # Just make sure to call Put() so your changes will get written back to the original object

###########################################################################################################################################################################################

# Tip 2: Finding Interesting WMI Classes

Get-WmiObject -List *                                                  # lists all WMI classes 

Get-WmiObject -List Win32_*video*                                      # lists all WMI classes related to video

Get-WmiObject Win32_VideoController                                    # specify one to check out the kind of information these classes return

###########################################################################################################################################################################################

# Tip 3: Getting Help for WMI Classes

# Method 1:
$class = [WMIClass]"Win32_VideoController"                             # retrieve Help information for any WMI class
$class.psbase.Options.UseAmendedQualifiers = $true
($class.psbase.Qualifiers["description"]).Value


# Method 2:  to get a list of all WMI methods present inside a specific WMI class – plus a meaningful description and a list of all the supported return values
$class = [wmiclass]"Win32_NetworkAdapterConfiguration"
$class.psbase.Options.UseAmendedQualifiers = $true
$class.psbase.methods | Where-Object {$rv = 1 | Select-Object Name, Help; $rv.Name = $_.Name; $rv.Help = ($_.Qualifiers["Description"]).Value; $rv} | Format-Table -AutoSize -Wrap

###########################################################################################################################################################################################

# Tip 4:　Calling ChkDsk via WMI

# Note: Make sure you have administrator privileges or else you receive a return value of 2 which corresponds to "Access Denied".
([wmi]"Win32_LogicalDisk='D:'").Chkdsk($true, $false, $false, $false, $false, $true).ReturnValue      # initiates a disk check on drive D

([wmi]"Win32_LogicalDisk='D:'").Chkdsk                                                                # to know what the arguments are for, ask for a method signature

###########################################################################################################################################################################################

# Tip 5: Renewing all DHCP Leases

([wmiclass]"Win32_NetworkAdapterConfiguration").RenewDHCPLeaseAll().ReturnValue                       # Some WMI classes contain static methods. Static methods do not require an instance

###########################################################################################################################################################################################

# Tip 6: Lowering Process Priority

Get-Process notepad | ForEach-Object {$_.PriorityClass = "BelowNormal"}                               # lowers priority for all Notepad processes to "below normal"

###########################################################################################################################################################################################

# Tip 7: Working with Path Names

# The .NET System.IO.Path class has a number of very useful static methods that you can use to extract file extensions
[System.IO.Path] | Get-Member -Static                                                                                        # get a list of available methods

[System.IO.Path]::ChangeExtension("test.txt","ps1")                                                                          # test.txt => test.ps1

###########################################################################################################################################################################################

# Tip 8: Reading File "Magic Number"

# Note: File types are not entirely dependent on file extension. Rather, binary files have internal ID numbers called "magic numbers" that tell Windows what type of file it is. 
function Get-MagicNumber($path)
{
    Resolve-Path $path | ForEach-Object {
        
        $magicNumber = Get-Content -Encoding Byte $_ -ReadCount 4 -TotalCount 4
        
        $hex1 = ("{0:x}" -f ($magicNumber[0] * 256 + $magicNumber[1])).PadLeft(4, "0")
        
        $hex2 = ("{0:x}" -f ($magicNumber[2] * 256 + $magicNumber[3])).PadLeft(4, "0")
        
        [string]$chars = $magicNumber | ForEach-Object {
         
            if([char]::IsLetterOrDigit($_)) 
            {
                [char]$_
            }
            else
            {
                "."
            }
        }
        
        "{0} {1} {2}" -f $hex1, $hex2, $chars
    }
}

Get-MagicNumber $env:windir\atiogl.xml                                                                                       # Output: 3c50 524f . P R O

###########################################################################################################################################################################################

# Tip 9: Displaying Hex Dumps

function Get-HexDump($path, $width = 10, $bytes = -1)
{
    $OFS = ""
    Get-Content -Encoding Byte $path -ReadCount $width -TotalCount $bytes | ForEach-Object {
        $byte = $_
        if(($byte -eq 0).Count -ne $width)
        {
            $hex = $byte | ForEach-Object {
                " " + ("{0:x}" -f $_).PadLeft(2, "0")
                $char = $byte | ForEach-Object {
                        if([char]::IsLetterOrDigit($_))
                        {
                            [char]$_
                        }
                        else
                        {
                            "."
                        }
                    }
                    "$hex $char"
                }
        }
    }
}

Get-HexDump $env:windir\explorer.exe -width 15 -bytes 150


# Getting Alphabetical Listings
$OFS = ","
[string][char[]](65..90)                                       # Output: A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z

###########################################################################################################################################################################################

# Tip 10: Bulk-Changing File Extensions

# renames all ps1 PowerShell script files found in your user profile and renames them to the new extension "old.ps1

# One: use .NET method
Dir $home\*.ps1 -recurse | Foreach-Object {Rename-Item $_.FullName ([System.IO.Path]::GetFileNameWithoutExtension($_.FullName) + ".old.ps1") -whatif}

# Two: works in V1 and V2
dir $home\*.ps1 -Recurse | ForEach-Object {Rename-Item $_.FullName (($_.FullName).TrimEnd($_.Extension) + ".old.ps1") -WhatIf}

# Three: works in PowerShell V2 or above only
dir $home\*.ps1 -Recurse | ForEach-Object {Rename-Item $_.FullName ((Join-Path $_.DirectoryName $_.BaseName) + ".old.ps1") -WhatIf}

###########################################################################################################################################################################################

# Tip 11: Windows Special Folder

# Desktop
[Environment]::GetFolderPath("Desktop")                                                        # The .NET Environment class can provide the paths to common folders like a user desktop
[enum]::GetNames([System.Environment+SpecialFolder])                                           #  to get a list of allowed folder names


# Recent
dir ([Environment]::GetFolderPath("Recent"))
del "$([Environment]::GetFolderPath("Recent"))\*.*" -WhatIf                                    # this can be used to clear private data


# Cookies
dir ([Environment]::GetFolderPath("Cookies"))
explorer ([Environment]::GetFolderPath("Cookies"))                                             # to open that folder in your Explorer
dir ([Environment]::GetFolderPath("Cookies")) | Select-String prefbin -List                    # select cookies based on content. lists all cookies that contain the word "prefbin”
dir ([Environment]::GetFolderPath("Cookies")) | del -WhatIf                                    # del all cookies

###########################################################################################################################################################################################

# Tip 12: Finding Cmdlets With a Given Parameter

Get-Command *service* -CommandType Cmdlet                                                      # Finding cmdlets by name 

Get-Help * -Parameter list                                                                     # to use Get-Help with the parameter -parameter to see all cmdlets with a -List parameter


filter Contains-Parameter                                                                      # This filter allows only those to pass the pipeline that supports the given parameter
{
    param($name)

    $number = @($_ | ForEach-Object {$_.ParameterSets | ForEach-Object {$_.Parameters | Where-Object {$_.Name -eq $name}}}).Count

    if($number -gt 0)
    {
        $_
    }
}

Get-Command | Contains-Parameter 'list'                       

###########################################################################################################################################################################################

# Tip 13: About Dates

Get-Date | Get-Member -MemberType *Method                # to see all methods available

(Get-Date).ToShortDateString()                           # Output: 2014/09/03

[DateTime] | Get-Member -Static                          # Finding Static Methods

[DateTime]::IsLeapYear(1904)                             # check whether a year is a leap year

[DateTime]::DaysInMonth(2009, 2)                         # Finding Days in Month
[DateTime]::DaysInMonth(2008, 2)

###########################################################################################################################################################################################

# Tip 14: Identifying 64-Bit-Environments

if([IntPtr]::Size -eq 8)
{
    "Ein 64-Bit-System"
}
else
{
    "Ein 32-Bit-System"
}

###########################################################################################################################################################################################

# Tip 15: Prompting for Passwords

# Note: If you need to prompt for a secret password and do not want it to be visible while entered, you should use Get-Credential. This cmdlet returns a credential object, 
 # which contains the entered password in encrypted format. You should then call GetNetworkCredential() to un-encrypt it to plain text

# Method One:
$cred = Get-Credential Administrator
$cred.Password                                             # Output: System.Security.SecureString

$cred.GetNetworkCredential()
# Output:
   # UserName              Domain                                            
   # --------              ------  
          
$cred.GetNetworkCredential().Password                      # Output: iis6!dfu                       # call GetNetworkCredential() to un-encrypt it to plain text



# Method Two:
# Retrieving Clear Text Password
$cred = Get-Credential
# to restore the clear text password entered into the dialog
$pwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $cred.Password ))
"Password: $pwd"



# Method Three:
$pwd = Read-Host 'Enter Password' -AsSecureString
(New-Object System.Management.Automation.PSCredential("Administrator", $pwd)).GetNetworkCredential().Password



# Method Four: Prompting for Secret Passwords via Console
$cred = $host.UI.PromptForCredential("Log on", "message", ".\administrator", "target")  # use Get-Credential or a low- level PS API function
$cred

# Note: You should nNote that the low- level function will allow you to set dialog captions and a custom message. For a custom message, you should replace $null with your message string

###########################################################################################################################################################################################

# Tip 16: Running Commands Elevated

Start-Process powershell.exe -ArgumentList "-command dir $env:windir" -Verb runas                  # launch a separate PS environment and elevate it as administrator

###########################################################################################################################################################################################

# Tip 17: Write-Output

# Note: When you leave data behind, PowerShell automatically returns it to the caller. This may create strange-looking code. 
 # With Write-Output, you can explicitly assign your return values. These two examples work the same:

function testA
{
    "My Result"
}

function testB
{
    Write-Output "My Result"
}

testA                                                    # Output: My Result
testB                                                    # Output: My Result



# Note: Assigning (multiple) return values with Write-Output works well, but you should keep in mind that Write-Output is picky and returns the exact thing you specified.
 # So, if you want to return a calculated result, you should make sure it is put in parenthesis

function Convert-Dollar2EuroA($amount, $rate = 0.8)
{
    $amount * $rate
}

function Convert-Dollar2EuroB($amount, $rate = 0.8)
{
    Write-Output $amount * $rate
}

function Convert-Dollar2EuroC($amount, $rate = 0.8)
{
    Write-Output ($amount * $rate)
}

Convert-Dollar2EuroA 100                                 # Output: 80
Convert-Dollar2EuroB 100                                 # Output: 100 * 0.8
Convert-Dollar2EuroC 100                                 # Output: 80

###########################################################################################################################################################################################

# Tip 18: Download Web Page Content

# Method 1:
# Note: Probably the easiest way of reading raw Web page content is using the Web client object

$url = "http://blogs.msdn.com/b/powershell/"
$object = New-Object System.Net.WebClient
$object.DownloadString($url)                             # output the Web content of http://blogs.msdn.com/b/powershell/       


# Method 2:  
$website = Invoke-WebRequest -Uri $url  
$website.RawContent       

###########################################################################################################################################################################################

# Tip 19: Scraping Information from Web Pages

# Regular expressions are a great way of identifying and retrieving text patterns. 
 # Take a look at the next code fragment as it defines a RegEx engine that searches for HTML divs with a "post-summary" attribute, 
 # then reads the PowerShell team blog and returns all summaries from all posts in clear text:

$regex = [regex]'<div class="post-summary">(.*?)</div>'
$url = "http://blogs.msdn.com/b/powershell/"
$object = New-Object System.Net.WebClient
$content = $object.DownloadString($url)
$regex.Matches($content) | ForEach-Object {$_.Groups[1].Value}



$regex = [regex]'<div class="post-summary">(.*?)</div>'
$url = "http://blogs.msdn.com/b/powershell/"
$content = (Invoke-WebRequest -Uri $url).RawContent
$regex.Matches($content) | ForEach-Object {$_.Groups[1].Value}


$url = "http://www.baidu.com"
Invoke-WebRequest -Uri $url

###########################################################################################################################################################################################

# Tip 20: Using 'Continue'

# One: use in loops to skip the remainder of a loop
for($x = 1; $x -lt 20; $x++)
{
    if($x % 4)
    {
        continue
    }
    $x
}                                                             # Output: 4   8    12    16



# Two: use continue to catch exception with your own message but not throw inner exception message

# Note: Without "Continue", your handler would trigger, but the exception would continue to bubble up, 
 # so PowerShell’s own handler would also see it and toss in a red PS error message as well.
trap
{
    "Whew, an error: $_"; continue
}

1 / $null

# Note: When you handle errors yourself using Trap or try/catch, you should make sure that you set your ErrorActionPreference to "Stop" or
 # specify -ErrorAction Stop for each cmdlet you want to handle. Otherwise, Cmdlet exceptions would be invisible to your error handlers and handled by PowerShell automatically. 

###########################################################################################################################################################################################

# Tip 21: About Variable


# One: Understanding Variable Inheritance

# PowerShell variables are inherited downstream by default, not upstream
$a = 1
function test
{
    "variable a contains $a"

    $a = 2

    "variable a contains $a"
}

test
     # variable a contains 1
     # variable a contains 2

$a                                                    # Output: 1

# Note: the function test receives variable $a from its parent scope. While it can define its own variable $a as well, 
 # it never affects the parent variable $a because it is only inherited downstream


# Two: Creating "Static" Variables

# Static variables are accessible everywhere. They can be used to collect data from various scopes in one place. 
 # You can use a prefix for a variable with "script:" to create a static variable

function test                                                # This example shows a recursive call that runs 10 nest levels. You can use a static variable to keep track of nest level
{
    $Script:nestLevel += 1                                   # Your static variable will also affect the console if you replace the prefix "script:" by "global:" 

    if($Script:nestLevel -gt 10)
    {
        break
    }

    "I am net level $Script:nestLevel"

    test
}

test



# Three: Working with Private Variables

 # Note: PowerShell inherits by default variables downstream so subsequent scopes can "see" the parent variables. If you want to turn off variable inheritance altogether, 
 # you should use the prefix "private:." This way, variables will only work in the scope in which they are defined, and will neither inherit upstream or downstream
$Private:a = 1
function test
{
    "variable a contains $a"

    $a = 2

    "variable a contains $a"
}

test
    # Output:  variable a contains 
    # Output:  variable a contains 2
$a                                                          # Output: 1

# Note: You should note that defining a variable as private will not overwrite an existing variable. So, if you had previously defined a variable "a" without the private: prefix, 
 # you would not be able to turn it into a private variable. To make sure, you should either start a new fresh PowerShell environment, 
 # or delete the variable before creating your private variable
Remove-Item variable:a

###########################################################################################################################################################################################

# Tip 22: Changing Console Colors

# You can use two different approaches to set the values. Both set the console background color to "Blue":
$host.UI.RawUI.BackgroundColor = "Blue"
[System.Console]::BackgroundColor = "Blue"                  
Clear-Host                                                  # enter Clear-Host to have the entire console background repainted


# Resetting Console Colors
[System.Console]::ResetColor()                              # only works if you set color by [System.Console]::BackgroundColor
Clear-Host

# Note: If you had changed colors using the PS object model through $host.UI.RawUI.BackgroundColor = 'Blue', ResetColor() would not return to the original color.

###########################################################################################################################################################################################

# Tip 23: Listing Available Culture IDs

[System.Globalization.CultureInfo]::GetCultures("AllCultures")                   # get all avaliable list, http://msdn.microsoft.com/en-us/library/system.globalization.culturetypes.aspx

$c = [System.Globalization.CultureInfo]"zh-cn"
[System.Threading.Thread]::CurrentThread.CurrentCulture = $c;Get-Date            # Output: 2014年9月4日 13:08:05

$c = [System.Globalization.CultureInfo]"en-us"
[System.Threading.Thread]::CurrentThread.CurrentCulture = $c;Get-Date            # Output: Thursday, September 04, 2014 13:08:46

# Note: You should note that the console character set is not able to display certain characters.
 # You may want to run that command inside the PowerShell ISE or another Unicode-enabled environment.

###########################################################################################################################################################################################

# Tip 24: Translating Culture IDs to Country Names

#  translate a culture ID to the full culture name

[System.Globalization.CultureInfo]::GetCultureInfoByIetfLanguageTag("en-us")
 #  LCID             Name             DisplayName                                                                                                                                                                        
 #  ----             ----             -----------                                                                                                                                                                        
 #  1033             en-US            English (United States)

[System.Globalization.CultureInfo]::GetCultureInfoByIetfLanguageTag("zh-cn")
 #  LCID             Name             DisplayName                                                                                                                                                                        
 #  ----             ----             -----------                                                                                                                                                                        
 #  2052             zh-CN            Chinese (Simplified, PRC)

 [System.Globalization.CultureTypes]

###########################################################################################################################################################################################

# Tip 25: Running Programs as Different User

# If you ever needed to run a program as someone else, you can use Start-Process and supply alternate credentials. 
 # When you do that, you should also make sure you specify -LoadUserProfile to load the user profile unless you do not need it:

Start-Process powershell -LoadUserProfile -Credential (Get-Credential)           #  to run a program as someone else

# With UAC enabled, this will always launch the program without administrator privileges. 
 # However, you can elevate another console from here if you like, and it will continue to run as the user you specified to launch the first program

Start-Process powershell -Verb runas                                             #  to run a program as administrator

###########################################################################################################################################################################################

# Tip 26: Finding Methods with Specific Keywords

# As such, .NET Framework is huge and full of stars, and it is not easy to discover interesting methods buried inside of it. 
 # You can use the next lines to find all methods with a given keyword:

$key = "kill"
[System.Diagnostics.Process].Assembly.GetExportedTypes() | Where-Object {$_.IsPublic} | 
    Where-Object {$_.IsClass} | ForEach-Object {$_.GetMethods()} | Where-Object {$_.Name -like "*$key*"} | Select-Object DeclaringType, Name

###########################################################################################################################################################################################

# Tip 27: Using OpenFile Dialog

# Method 1: Use .Net method
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null     # to open a standard OpenFile dialog in your PowerShell scripts

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.DefaultExt = ".ps1"
$dialog.Filter = 'PowerShell-Skripts|*.ps1|All Files|*.*'
$dialog.InitialDirectory = $home
$dialog.Multiselect = $false
$dialog.RestoreDirectory = $true
$dialog.Title = "Select a script file"
$dialog.ValidateNames = $true
$dialog.ShowDialog()                                             # Output: OK | Cancel
$dialog.FileName

# Important note: Dialogs only work correctly when you launch PowerShell with the -STA option! So before you enter and run the code, be sure to open the correct PowerShell environment

function Select-FileDialog
{
    param([string]$Title, [string]$Directory, [string]$Filter="All Files (*.*)|*.*")
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    
    $objForm = New-Object System.Windows.Forms.OpenFileDialog
    $objForm.ShowHelp = $true
    $objForm.InitialDirectory = $Directory
    $objForm.Filter = $Filter
    $objForm.Title = $Title
    
    $Show = $objForm.ShowDialog()
    
    If ($Show -eq "OK")
    {   
        Return $objForm.FileName    
    }
    
    Else
    { 
        Write-Error "Operation cancelled by user."   
    }
}

$file = Select-FileDialog -Title "Select a file" -Directory $home -Filter "Powershell Scripts(*.ps1)|*.ps1"
$file



# Method 2: Use COM method
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

# Tip 28: Downloading Internet Files with Dialog

$url = 'http://www.idera.com/images/Tours/Videos/PowerShell-Plus-IDE-1.wmv'
$dest = "$home\video.wmv"

[Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null

$object = New-Object Microsoft.VisualBasic.Devices.NetWork
$object.DownloadFile($url, $dest, "", "", $true, 500, $true, "DoNothing")

Invoke-Item $dest

# There is a great way to download large files from the Internet. This example downloads a tutorial video from Idera that, once downloaded will run in your media player

###########################################################################################################################################################################################

# Tip 29: Why Native Commands Fail In PowerShell

# You probably already know that not all native commands work equally well in PowerShell. Have a look here:
find /I /N "dir" *.ps1                              # Exception here, In cmd.exe, this line finds all occurrences of "dir" in all PowerShell scripts that are located in the current path


# Solution 1: to execute the line in good old cmd.exe and then enclose the command in single quotes
cmd.exe /c 'find /I /N "dir" *.ps1'



# PowerShell parser:

  #   " 	The beginning (or end) of quoted text
  #   # 	The beginning of a comment
  #   $ 	The beginning of a variable
  #   & 	Reserved for future use
  #   ( ) 	Parentheses used for sub-expressions
  #   ; 	Statement separator
  #   { } 	Script block
  #   | 	Pipeline separator
  #   ` 	Escape character

# Why does this work in cmd.exe and not in powershell.exe?
Trace-Command NativeCommandParameterBinder {find /I "dir" *.ps1} -PSHost


# Solution 2: 
# Actually, find.exe expected two parameters, not three. The first two parameters should stay together. You can get this done by removing the space in between:
find /I` /N` "dir" *.ps1


# Solution 3:

find /I /N '"dir"' *.ps1
# As it turns out, the command requires the quotes, and PowerShell interprets them as string delimiter and removes them. So to make the command work, you should use this instead

###########################################################################################################################################################################################

# Tip 30: Running Commands in the Background

# You can also transfer commands into another PowerShell session and run it in the background

$job = Start-Job {dir $env:windir *.log -Recurse -ErrorAction SilentlyContinue}      # find all log files recursively in your Windows folder and all of its sub-folders as background job

$job                                                                                 # To check the status of your job, you can check the $job variable

Receive-Job $job                                                                     # To retrieve the collected results from your background job, you can use Receive-Job

###########################################################################################################################################################################################

# Tip 31: Get Notification When a Background Job is Done

# When you assign long-running commands to a background session, 
 # you may want to get some notification when the job is completed so you don't have to constantly check its status. Here is how:
$job = Start-Job -Name GetLogFiles {dir $env:windir *.log -Recurse -ErrorAction SilentlyContinue}

Register-ObjectEvent $job StateChanged -Action {
    
    [Console]::Beep(1000, 500)

    Write-Host ("Job #{0} ({1}) complete." -f $sender.Id, $sender.Name) -ForegroundColor White -BackgroundColor Red
    Write-Host "Use this command to retrieve the results:"
    Write-Host (prompt) -NoNewline
    Write-Host ("Receive-Job -ID {0}; Remove-Job -ID {0}" -f $sender.Id)
    Write-Host (prompt) -NoNewline
    
    $eventSubscriber | Unregister-Event
    $eventSubscriber.Action | Remove-Job

} | Out-Null

# This will create a temporary event subscriber that deletes itself once the event is triggered. 
 # It will output a message when the background job is completed and also gives a hint on how to retrieve the results from that background job

# A interesting voice for warning
1..30 | ForEach-Object {
    $frequency = Get-Random -Minimum 400 -Maximum 10000 
    $duration  = Get-Random -Minimum 100 -Maximum 500

    [Console]::Beep($frequency,$duration)
  }

###########################################################################################################################################################################################

# Tip 32: Running 32-Bit-Code on 64-Bit Machines

$32bitCode = {[IntPtr]::Size}

& $32bitCode                               # Output: 8                    # run on 64bit machine directly


$job = Start-Job $32bitCode -RunAs32       # Output: 4                    # run in isolated 32bit session
$job | Wait-Job | Receive-Job
Remove-Job $job
"Done."

###########################################################################################################################################################################################

# Tip 33: Create Files and Folders in One Step

if(!(Test-Path $home\subfolder\anothersubfolder\yetanotherone\test.txt))
{
    # Use New-Item like this when you want to create a file plus all the folders necessary to host the file
    New-Item -ItemType file -Force $home\subfolder\anothersubfolder\yetanotherone\test.txt 
}

# Note: This will create the necessary folders first and then insert a blank file into the last folder. You can then edit the file easily using notepad. 
 # However, you should watch out as this line will overwrite the file if it already exists. First, you should use Test-Path to check and run the line only if the file is missing.

###########################################################################################################################################################################################

# Tip 34: Strongly Typed Arrays

$array = [Int[]](1..5)
$array.GetType().FullName                  # Output: System.Int32[]
$array

$array += 6
$array.GetType().FullName                  # Output: System.Object[]
$array
# Note: When you assign strongly typed values to an array, the type declaration will remain intact only as long as you do not add new array content.
 # Once you add new content using +=, the complete array will be copied into a new one, and type information is lost



# You can work around this by simply strongly type the variable that stores the array instead, As you can see, the type information is preserved this way:

[Int[]]$array = 1..5
$array.GetType().FullName                  # Output: System.Int32[]
$array

$array += 6
$array.GetType().FullName                  # Output: System.Int32[]
$array

###########################################################################################################################################################################################

# Tip 35: Get WebClient with Proxy Authentication

# If your company is using an Internet proxy, and you'd like to access Internet with a webclient object, make sure it uses the proxy and supplies your default credentials to it
function Get-WebClient
{
    $wc = New-Object Net.WebClient
    $wc.UseDefaultCredentials = $true
    $wc.Proxy.Credentials = $wc.Credentials
    $wc
}

# retrieve an RSS feed from the Internet
$webClient = Get-WebClient
[xml]$powershellTips = $webClient.DownloadString('http://powershell.com/cs/blogs/tips/rss.aspx')
$powershellTips.rss.channel.item | Select-Object Title, Link

###########################################################################################################################################################################################

# Tip 36: Enter-PSSession - Do's and Dont's

# Enter-PSSession will let you switch your console input to a remote computer—if remoting is enabled on the target computer. Essentially, 
# anything you enter after Enter-PSSession is sent to the remote computer that you specified with -ComputerName, and any result is marshaled back to your console. 
# This, of course, requires PowerShell Remoting to be setup appropriately on the target machine.


# However, you should be aware that Enter-PSSession only sends your interactive commands to the remote machine. 
# So, it does not make sense to pre-pend script code with Enter-PSSession to run your script remotely. The script will not run remotely this way. Instead, 
# it will run on your local machine. Instead, you should use Invoke-Command -ComputerName XYZ -FilePath c:\myscript.ps1 to run scripts remotely.

###########################################################################################################################################################################################

# Tip 37: Create (and Edit) your Profile Script

# Profile scripts are automatically executed whenever PowerShell launches. Your profile script is the perfect place to customize your PowerShell environment, change the prompt, 
 # colors, and make any changes you would like to keep in all of your sessions.

# here is a line that makes sure the profile script exists before it opens it in the ISE PowerShell editor:
if((Test-Path $profile) -eq $false)
{
    New-Item $profile -ItemType file -Force -ErrorAction 0 | Out-Null
    ise $profile
}

# If you are sure the profile already exists, you can shorten the line
ise $profile

# You need to make sure the execution policy allows script execution if you want your profile script to run automatically the next time you launch PowerShell:
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force


# to see the changes that are in effect without having to close and restart PowerShell, you can simply run your profile script dot-sourced
. $profile


# get all profile
$profile.psextended | Format-List *

###########################################################################################################################################################################################

# Tip 38: Determining Function Parameters Supplied by User

# To find out which parameters that a user submitted to a self-defined function, you can use $PSCmdlet like this
function Get-Parameters
{
    [CmdletBinding()]
    param($name, $surname = "Default", $age, $id)

    $PSCmdlet.MyInvocation.BoundParameters.GetEnumerator()
}

Get-Parameters -name "Silence" -id 1

# Output: 
          #  Key                                 Value                                                                                                     
          #  ---                                 -----                                                                                                     
          #  name                                Silence                                                                                                   
          #  id                                  1    

###########################################################################################################################################################################################

# Tip 39: Processing Function Parameters As Hash Table

# If you want to get more control over function parameters, you can treat them as hash table. 
 # So in your function, you can then check whether a certain parameter was specified by the caller, and then act accordingly

function Do-Something
{
    [CmdletBinding()]
    param($name, $surname = "Default", $age, $id)

    $params = $PSCmdlet.MyInvocation.BoundParameters

    if($params.ContainsKey("name"))
    {
        "You specified -name and submitted {0}" -f $params["name"]
    }
    else
    {
        "You did not specify -name"
    }
}

Do-Something                              # Output: You did not specify -name
Do-Something -id 1                        # Output: You did not specify -name
Do-Something -name "Silence"              # Output: You specified -name and submitted Silence

###########################################################################################################################################################################################

# Tip 40: Test Admin Privileges

function is-Admin
{
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal $identity
    $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

is-Admin                                                                       # It will return $true if you do have Administrator rights, else $false                                          

###########################################################################################################################################################################################

# Tip 41: Launching PowerShell Scripts with Admin Privileges

# Note: If you must ensure that a PowerShell script runs with Admin privileges, you can add this to the beginning of your script

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = new-object Security.Principal.WindowsPrincipal $identity

if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)  -eq $false) 
{
	$Args = '-noprofile -nologo -executionpolicy bypass -file "{0}"' -f $MyInvocation.MyCommand.Path
	Start-Process -FilePath 'powershell.exe' -ArgumentList $Args -Verb RunAs
	exit 
}

"Running with Admin Privileges"
Read-Host "PRESS ENTER"

###########################################################################################################################################################################################

# Tip 42: Enumerating Network Adapters

# Finding your network adapters with WMI isn't always easy because WMI treats a lot of "network-like" adapters like network adapters. 
 # To find only those adapters that are also listed in your control panel, you should make sure to filter out all adapters that have no NetConnectionID, 
 # which is the name assigned to a network adapter in your control panel

function Get-NetworkAdapter
{
    Get-WmiObject Win32_NetworkAdapter -Filter "NetConnectionID!=null"
}

Get-NetworkAdapter | Select-Object Caption, NetCon*                       # list all network adapters and check their status
# Output:
         # Caption                                                        NetConnectionID                    NetConnectionStatus
         # -------                                                        ---------------                    -------------------
         # [00000007] Broadcom NetXtreme Gigabit Ethernet                 Local Area Connection                                2

###########################################################################################################################################################################################

# Tip 43: Disabling Network Adapters

# If you need to systematically disable network adapters, all you need is the name of the adapter 
 # (as stated in your control panel or returned by Get-NetworkAdapter, a function featured in another tip). Of course, you will also need Administrator privileges. 

function Disable-NetWorkAdapter
{
    param($name)

    Get-WmiObject Win32_NetworkAdapter -Filter "NetConnectionID='$name'" | ForEach-Object {
        
        $rv = $_.Disable().RetrunValue

        if($rv -eq 0)
        {
            "{0} disabled" -f $_.Caption
        }
        else
        {
            "{0} could not be disabled. Error code {1}" -f $_.Caption, $rv
        }
    }
}

###########################################################################################################################################################################################

# Tip 44: Hiding NetworkAdapter

# Note: When you use VMware, or if you have installed a Microsoft Loopback adapter, these adapters will show up in your network panel as "Unidentified Networks," 
 # and Windows assigns to them a "public network" status. That can be bad for a number of reasons. For example, when there are public networks, you cannot enable PowerShell remoting. 
 # While you could disable all virtual network adapters, a better way is to hide them from network location awareness.

# find a network adapter 
function Get-NetworkAdapter($name = "*")
{
    $key = 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}\*'

    Get-ItemProperty $key -ErrorAction SilentlyContinue | Where-Object {$_.DriverDesc -like $name} | Select-Object DriverDesc, PSPath
}


# hide the adapter from NLA
function Set-NetworkDaapter
{
    param([Parameter(ValueFromPipelineByPropertyName = $true)]$psPath, [switch]$ignoreNLA)

    process
    {
        if($ignoreNLA)
        {
            New-ItemProperty $psPath -Name "*NdisDeviceType" -PropertyType dword -Value 1 | Out-Null
        }
        else
        {
            Remove-ItemProperty $psPath -Name "*NdisDeviceType"
        }
    }
}

Get-NetworkAdapter
Get-NetworkAdapter "Microsoft LoopbackAdapter"
Get-NetworkAdapter "Microsoft LoopbackAdapter" | Set-NetworkDaapter -ignoreNLA

# Note: You should be aware that you will need Administrator privileges to change network adapter settings. Settings will take effect once you restart the network adapter, 
 # such as by disabling and enabling or by rebooting your machine. Be careful: do not hide your regular network adapters! If you did this by accident, 
 # you can remove the registry value by running Set-NetworkAdapter without the -IgnoreNLA switch.

###########################################################################################################################################################################################

# Tip 45: Using Robocopy to Copy Stuff

# You should avoid cmdlets if you need to copy large files or a large number of files. Instead, you should use native commands like robocopy

# This script collects all log files three levels deep into your Windows folder and copies them to some other folder
robocopy $env:windir\ $env:SystemDrive\logfiles\ *.log /R:0 /LEV:3 /S /XD *winsxs* | Where-Object {$_} | ForEach-Object {
    
    Write-Progress "Robocopy is looking for LOG-files in: " $_.split([char]9)[-1]
}

$file = dir $env:SystemDrive\logfiles *.log -Recurse
$file | Move-Item -Destination $env:SystemDrive\logfiles\ -Force
dir $env:SystemDrive\logfiles | Where-Object {$_.PSIscontainer} | del -Recurse -Force

# Note: Robocopy is fast, robust, and highly configurable. Take, for example, the above code where the tool only traverses three levels deep into the Windows sub-folders 
 # and omits the winsxs-subfolder. Although robocopy is a native command, you can see that PowerShell’s pipeline can capture and display its feedback messages in real time. 
 # It is part of Windows Server 2008 and Windows 7 and can also be downloaded separately as part of the resource kit.

###########################################################################################################################################################################################

# Tip 46: Refreshing Web Pages

# Imagine you opened a number of Web pages in Internet Explorer and would like to keep the display current. 
 # Instead of manually reloading these pages in intervals, you can use this script

function Refresh-WebPages($interval = 5)                                                   # It will automatically refresh all opened Internet Explorer pages every five seconds
{
    "Refreshing IE Windows every $interval seconds."
    "Press any key to stop."

    $shell =New-Object -ComObject Shell.Application

    do
    {
        "Refresh All HTML"

        $shell.Windows() | Where-Object {$_.Document.url} | ForEach-Object {$_.Refresh()}

        Start-Sleep -Seconds $interval

    }until([System.Console]::KeyAvailable)

    [System.Console]::ReadKey($true) | Out-Null
}

Refresh-WebPages                                                                           # Not work in ISE environment but works well in powershell console.

###########################################################################################################################################################################################

# Tip 47: Auditing PowerShell Scripts

Get-ChildItem $env:windir *.ps1 -Recurse -ErrorAction SilentlyContinue | Get-AuthenticodeSignature | Where-Object {$_.status -ne "Valid"}

# This will find all PowerShell scripts on drive c:\ and check their digital signature. 
# Any script without a signature or having an invalid signature will get reported back to you. Next, you can double-check those scripts and then sign them if they are OK.

###########################################################################################################################################################################################

# Tip 48: Finding Systems Online (Fast)

# Method 1:

# Using PowerShell Background Jobs, you can find a large number of online systems within only a few seconds

function Check-Online($computerName)
{
    Test-Connection -Count 1 -ComputerName $computerName -TimeToLive 5 -AsJob | Wait-Job | Receive-Job | Where-Object {$_.StatusCode -eq 0} | Select-Object -ExpandProperty Address
}

$ips = 1..255 | ForEach-Object {"10.10.10.$_"}                              # This code pings an IP segment from 10.10.10.1 to 10.10.10.255 and returns only those IPs that respond
$online = Check-Online -computerName $ips
$online

$online | Sort-Object | ForEach-Object {
    
    $ip = $_
    try
    {
        [System.Net.Dns]::GetHostAddresses($ip)                             # Resolving Host Names by IP Address
    }
    catch
    {
        "can not resolve $ip. Reason: $_"
    }
}


# Method 2:

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

###########################################################################################################################################################################################

# Tip 49: Using MsgBox Dialogs

# One: VBScript -- Mark a choice

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null
$rv = [Microsoft.VisualBasic.Interaction]::MsgBox("Do you want this happen?", "YesNoCancel, Exclamation, MsgBoxSetForeground, SystemModal", "Accept or Deny")

switch($rv)
{
    "Yes"     {"OK, we'll do it!"}
    "No"      {"Next time maybe ..."}
    "Cancel"  {"You cancelled ..."}
}



# Two: .Net -- Message Box

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Windows.Forms.MessageBox]::Show("Hello Silence, love you!")

###########################################################################################################################################################################################

# Tip 50: Create Random Passwords

function Get-RandomPassword
{
    param($length = 8, $characters = 'abcdefghkmnprstuvwxyzABCDEFGHKLMNPRSTUVWXYZ123456789!"§$%&/()=?*+#_')

    $random = 1..$length | ForEach-Object {Get-Random -Maximum $characters.length}

    $private:OFS = ""

    [String]$characters[$random]
}


function Randomize-Text($text)
{
    $anzahl = $text.length - 1
    $indizes = Get-Random -InputObject (0..$anzahl) -Count $anzahl

    $private:OFS = ""
    [String]$text[$indizes]
}


function Get-ComplexPassword
{
    $password  = Get-RandomPassword -length 2 -characters "abcdefghiklmnprstuvwxyz"
    $password += Get-RandomPassword -length 2 -characters "~!@#$%^&*()_+<>?;|\."
    $password += Get-RandomPassword -length 2 -characters "0123456789"
    $password += Get-RandomPassword -length 2 -characters "ABCDEFGHKLMNPRSTUVWXYZ"

    Randomize-Text $password
}

Get-RandomPassword                                                                  # create simple random password
Get-ComplexPassword                                                                 # create random passwords that meet certain requirements

# Note: Get-ComplexPassword creates a 14-character password consisting of six lowercase characters, two special characters, two numbers, and four uppercase characters. 
 # This is done by creating four parts and then melting them together with Randomize-Text, which takes a text and randomly re-arranges the characters. With this toolset, 
 # you can then create almost any random password adhering to almost any complexity rule you are likely to run into.

###########################################################################################################################################################################################

# Tip 51: Adding Quotes in Quotes

$text = "Hello ""World"""
$text                              # Output: Hello "World"


$text = 'Hello ''World'''
$text                              # Output: Hello 'World'


# Note: The use of single-quotation marks does not provide for the expansion of variables or expressions to their value. The use of double-quotation marks does.
$a = 'Hello'

$text = "$a ""World"""
$text                              # Output: Hello "World"

$text = '$a "World"'
$text                              # Output: $a "World"

###########################################################################################################################################################################################

# Tip 52: Use Back References with -Replace

# If you need to replace text with some other text and keep a reference to the original text, you can use the backreferenceback reference placeholder $0

'The problem was described in KB123456. Look it up.' -replace 'KB\d{6}', 'KB99999 (was: $0)'
# Output:  The problem was described in KB99999 (was: KB123456). Look it up.



'After each comma,I want a whitespace,but only if there was no whitespace in the first place,of course.' -replace  ',(\S)', ', $1'
# Output: After each comma, I want a whitespace, but only if there was no whitespace in the first place, of course.

# The regular expression is looking for any occurrences of a comma followed immediately by non-whitespace (/S). 
 # If found, replaces it with a comma, a space and whatever the non-whitespace character was (the match in parenthesis, represented by the back reference in $1).


'this text text contains duplicate words words following each other'  -replace '\b(\w+)(\s+\1){1,}\b', '$1'                 # Eliminating Duplicate Words
# Output: this text contains duplicate words following each other

###########################################################################################################################################################################################

# Tip 53: Disabling Automatic Page Files

# If you would like to programmatically control page file settings, you can use WMI but must enable all privileges using -EnableAllPrivileges.

$c = Get-WmiObject Win32_ComputerSystem -EnableAllPrivileges
$c.AutomaticManagedPagefile = $false                               # this will disable automatic page files
$c.Put()

###########################################################################################################################################################################################

# Tip 54: Reading Text Files Fast

get-process | Export-Clixml $home\data.xml
(dir $home\data.xml | Select-Object -ExpandProperty Length) / 1MB


Measure-Command {Get-Content $home\data.xml} | Select-Object -ExpandProperty TotalMilliseconds                                      # Output: 1758.7565 s
Measure-Command {[System.IO.File]::ReadLines("$home\data.xml")}  | Select-Object -ExpandProperty TotalMilliseconds                  # Output: 152.5976  s


Measure-Command {1..25 | %{Get-Content -ReadCount 0 $home\data.xml}} | Select-Object -ExpandProperty TotalMilliseconds              # Output: 5563.7991 s
Measure-Command {1..25 | %{[System.IO.File]::ReadAllLines("$home\data.xml")}} | Select-Object -ExpandProperty TotalMilliseconds     # Output: 4270.915  s

# Note: According to the above result, [System.IO.File]::ReadAllLines() is much faster than get-content


# retrun [array]
(Get-Content $home\data.xml) -is [array]                                  # Output: True
([System.IO.File]::ReadAllLines("$home\data.xml")) -is [array]            # Output: True


# retrun [String]
(Get-Content $home\data.xml -Raw) -is [array]                             # Output: False
([System.IO.File]::ReadAllText("$home\data.xml")) -is [array]             # Output: False

###########################################################################################################################################################################################

# Tip 55: Checking Whether User or Group Exists

[ADSI]::Exists("WinNT://./Tobias1")                                       # check whether there is a local account named "Tobias1." 

[ADSI]::Exists("LDAP://CN=Testuser,CN=Users,DC=YourDomain,DC=Com")        # To check domain accounts, you can simply replace "." with your domain name, or use LDAP

[ADSI]::Exists("WinNT://fareast/v-sihe")                                  # Output: True

###########################################################################################################################################################################################

# Tip 56: Comparing Services

# to compare the service status on two machines and find out where services are configured differently:
$local = Get-Service -ComputerName localhost
$remote = Get-Service -ComputerName IIS-CTI5052

Compare-Object -ReferenceObject $local -DifferenceObject $remote -Property Name,Status,MachineName -PassThru | Sort-Object Name | Select-Object Name,Status,MachineName

# Note: Compare-Object can help when troubleshooting computers


# Comparing Hotfixes
$local = Get-HotFix -ComputerName localhost
$remote = Get-HotFix -ComputerName IIS-CTI5052

Compare-Object -ReferenceObject $local -DifferenceObject $remote -Property HotFixID -IncludeEqual

###########################################################################################################################################################################################

# Tip 57: Setting Mouse Position

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(500,100)                      # PowerShell can place the mouse cursor anywhere on your screen.

###########################################################################################################################################################################################

# Tip 58: Sending Mails With Outlook

# You can use PowerShell to automatically prepare your e-mails and send them via Outlook. 
 # The code requires that you first launch Outlook as an application. It will not work when Outlook has not yet been started.

# start-process outlook
$outlook = New-Object -ComObject Outlook.Application

$mail = $outlook.CreateItem(0)
$mail.subject = "Test Message"
$mail.body    = "First Line `n Second Line `n Third Line"
$mail.To      = "v-sihe@microsoft.com"
$mail.Attachments.Add("$home\test.txt")

$mail.Send()
$outlook.Quit()

###########################################################################################################################################################################################

# Tip 59: Formatting a Drive

# (Get-WmiObject Win32_Volume -Filter "DriveLetter='D:'").Format("NTFS", $true, 4096, $false)       # format drive D:\ using NTFS file system, provided you have appropriate permissions

# Note: In Windows Vista and higher, there is a new WMI class called Win32_Volume that you can use to format drives. However, you should be careful when formatting data on a drive

###########################################################################################################################################################################################

# Tip 60: Identifying Computer Hardware

# Note: Unless the hard drive is exchanged, this number provides a unique identifier
Get-WmiObject Win32_DiskDrive | Select-Object -ExpandProperty SerialNumber                     # If you must identify computer hardware, you could do so on the hard drive serial number
# Output: 2020202057202d44435756413239363430393737


Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID                 # That is the id that windows uses to identify computers when booting via PXE for example
# Output: 0B2130FF-9238-11DE-BBD8-05B89BD418A9

###########################################################################################################################################################################################

# Tip 61: Playing Sound in PowerShell

# If you would like to catch a user’s attention, you can make PowerShell beep like this:
[System.Console]::Beep()
[System.Console]::Beep(1000, 300)

# A nicer sound can be played this way
[System.Media.SystemSounds]::Beep.Play()
[System.Media.SystemSounds]::Asterisk.Play()
[System.Media.SystemSounds]::Exclamation.Play()
[System.Media.SystemSounds]::Hand.Play()

###########################################################################################################################################################################################

# Tip 62: Get Logged On User

Get-WmiObject Win32_ComputerSystem -ComputerName localhost | Select-Object -ExpandProperty UserName

# Note that this will always return the user logged on to the physical machine. It will not return terminal service user or users inside a virtual machine. 
 # You will need administrator privileges on the target machine. Get-WmiObject supports the -Credential parameter if you must authenticate as someone else.




# Get All Logged On Users: discover logged-on terminal service users or users inside virtual machines
$computername = 'localhost'
Get-WMIObject Win32_Process -filter 'name="explorer.exe"' -computername $ComputerName |	
    ForEach-Object { $owner = $_.GetOwner(); '{0}\{1}' -f $owner.Domain, $owner.User } | Sort-Object | Get-Unique

###########################################################################################################################################################################################

# Tip 63: Automated Authentication

# You will not want a credential dialog to pop up if you need to run scripts unattended that need to authenticate using credentials. 
 # Here is an example of how to hard-code credentials into your scripts

$user = ".\Administrator"
$password = "password" | ConvertTo-SecureString -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($user, $password)
Start-Process notepad.exe -Credential $cred -LoadUserProfile                                        # launches a process automatically as a different user

###########################################################################################################################################################################################

# Tip 64: Filtering Multiple File Types

filter Where-Extension([string[]]$extension = (".bmp", ".jpg", ".wmv"))
{
    $_ | Where-Object {$extension -contains $_.Extension}
}

dir $env:windir -Recurse -ErrorAction SilentlyContinue | Where-Extension .log, .txt




dir $env:windir -Recurse -ErrorAction SilentlyContinue -Include *.log, *.txt                         # the -Include parameter is that it qualifies the -Path parameter
                                                                                                     
                                                                                                     
dir $env:windir -ErrorAction SilentlyContinue -Include *.log, *.txt                                  # return nothing

# So if you don't use the -Recurse parameter, you must specify the content of the directory, like 
dir $env:windir\* -ErrorAction SilentlyContinue -Include *.log, *.txt
# Note: Note the '*' wildcard after windir.


# Note: The function doesn’t run until all data has been stored in the $input variable,
 # whereas the filter has immediate access to the $_ variable and begins processing elements as they become available. 
 # Filters are more efficient when large amounts of data are passed through the object pipeline.

###########################################################################################################################################################################################

# Tip 65: Checking -STA Mode

# PowerShell needs to run in STA mode to display Windows Presentation Foundation (WPF) windows. 
 # ISE runs in STA mode by default whereas the console will need to be launched explicitly with the -STA switch
$isSTAEnabled = $host.Runspace.ApartmentState -eq "STA"
$isSTAEnabled                                                    # Output: True, ISE runs in STA mode by default

if($isSTAEnabled -eq $false)
{
    "You need to run this script with -STA switch or inside ISE"
}


# If you want to have the script automatically launch itself in STA mode, then you can use the following modified code:
$isSTAEnabled = $host.Runspace.ApartmentState -eq "STA"
$host.UI.RawUI.WindowTitle = "Window Title"

if($isSTAEnabled -eq $false)
{
    "Script is not running in STA mode. Wwitching to STA mode ..."

    $script = $MyInvocation.MyCommand.Definition
    Start-Process powershell.exe -ArgumentList "-sta $script"
    
    exit
}

# Place the code at the beginning of the script processing before everything else, or after parameter declaration if you do that sort of thing 

###########################################################################################################################################################################################

# Tip 66: Checking Loaded Assemblies

$host.Runspace.RunspaceConfiguration.Assemblies                                             # to check which .NET assemblies are currently loaded into PowerShell


# Checking Loaded Formats
$host.Runspace.RunspaceConfiguration.Formats | Select-Object -ExpandProperty FileName       # to check which format files have been loaded

###########################################################################################################################################################################################

# Tip 67: Loading .EVT/.EVTX Event Log Files

# If customers send in dumped event log files, there is an easy way to open them in PowerShell and analyze content: Get-WinEvent! 
 # The -Path parameter will allow you to read in those binary dumps and display the content as an object. 

Get-WinEvent -Path C:\sample.evt | Where-Object {$_.Level -eq 2} | Select-Object Message, TimeCreated, ProviderName, TimeCreated | 
    Export-Csv $env:TEMP\list.csv -UseCulture -Encoding UTF8 -NoTypeInformation

Invoke-Item $env:TEMP\list.csv

###########################################################################################################################################################################################

# Tip 68: Opening Excel Reports in a New Window

# When opening CSV files with Excel from PowerShell, you may receive exceptions if the particular file was opened by Excel already
Invoke-Item C:\files\report.csv

# You can work around this by telling Excel to open a new window for the CSV file
Start-Process excel -ArgumentList c:\file\report.csv

# This will load c:\files\report.csv in a new Excel window, so you do not get an exception if the file was already open in another instance of Excel. 
 # Instead, you will get a friendly dialog asking if you want to open the new file as write-protected copy

###########################################################################################################################################################################################

# Tip 69: Backing Up Event Log Files

# WMI provides a method to backup event log files as *.evt/*.evtx files. The code below creates backups of all available event logs:
Get-WmiObject Win32_NTEventLogFile | ForEach-Object {
    
    $filename = "$home\" + $_.LogfileName + ".evtx"
    del $filename -ErrorAction SilentlyContinue
    $_.BackupEventLog($filename).ReturnValue
}

# By the way, you can read in the *.evt/*.evtx files created by this approach using Get-WinEvent -Path.

###########################################################################################################################################################################################

# Tip 70: Launching Scripts Externally

# to launch a *.ps1 PowerShell script externally from outside PowerShell via desktop shortcut or from inside a batch file
powershell.exe -nologo -excutionpolicy bypass -noprofile -file "D:\myscript.ps1" 


# to simply execute a PowerShell command
powershell.exe -nologo -noprofile -command get-process

###########################################################################################################################################################################################

# Tip 71: Accessing Web Services

# PowerShell can access public and private Web services

$web = New-WebServiceProxy 'http://www.webservicex.net/globalweather.asmx?WSDL'
$web.GetCitiesByCountry("China")
$web.GetWeather('Beijing', 'China')                                      # connect to a global weather service providing airport weather reports from around the globe

###########################################################################################################################################################################################

# Tip 72: Speed Up The Reading of Large Text Files

Get-Content $env:windir\windowsupdate.log | Where-Object {$_ -like "*successfully installed*"}

$txt = Get-Content $env:windir\windowsupdate.log -ReadCount 0
$txt | Where-Object {$_ -like "*successfully installed*"}

# When you use Get-Content to read in text files, you may initially be disappointed by its performance. 
 # However, Get-Content is slow only because it emits each line to the pipeline as it reads the file, which is time-consuming. 
 # You can dramatically speed up reading large text files by adding the parameter -ReadCount 0 to Get-Content. This way, the file is read and only then passed on to the pipeline.

###########################################################################################################################################################################################

# Tip 73: Sorting IP Addresses

# Sorting or comparing IP addresses won't initially work because PowerShell uses alphanumeric comparison. 
 # However, you can compare or sort them correctly by casting IP addresses temporarily to the type System.Version

$iplist = "10.10.10.1", "10.10.10.3", "10.10.10.230"

$iplist
# Output:
#        10.10.10.1
#        10.10.10.3
#        10.10.10.230


$iplist | Sort-Object
# Output:
#        10.10.10.1
#        10.10.10.230
#        10.10.10.3


$iplist | ForEach-Object {[System.Version]$_} | Sort-Object | ForEach-Object {$_.ToString()}
# Output:
#        10.10.10.1
#        10.10.10.3
#        10.10.10.230


# Here's a simpler solution:
$iplist | Sort-Object @{ Expression = {[System.Version]$_} }
# Output:
#        10.10.10.1
#        10.10.10.3
#        10.10.10.230

###########################################################################################################################################################################################

# Tip 74: Find Out If A Machine Is Connected To The Internet

$networkListManager = [Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]"{DCB00C01-570F-4A9B-8D69-199FDBA5723B}"))

$connections = $networkListManager.GetNetworkConnections()

$connections | ForEach-Object {$_.isConnectedToInternet}

###########################################################################################################################################################################################

# Tip 75: Entering Passwords Securely

$password = Read-Host "Password" -AsSecureString                      # use Read-Host -asSecureString to be able to type a password with hidden characters

$password = (New-Object System.Management.Automation.PSCredential("dummy", (Read-Host "Password" -AsSecureString))).GetNetworkCredential().Password  #get the password back into plain text



# Mandatory Password Parameters
function Test-Password
{
    param([System.Security.SecureString][Parameter(Mandatory = $true)]$password)

    $plain = (New-Object System.Management.Automation.PSCredential("dummy", $password)).GetNetworkCredential().Password

    "You entered: $plain ."
}

# Note: If you mark a parameter as mandatory and set its type to "SecureString," PowerShell will automatically prompt for the password with masked characters

###########################################################################################################################################################################################

# Tip 76: Validating Input: use regular expressions to validate user input

do
{
    $result = Read-Host "3-7-digit number"
    $result = $result.Trim("0")

}while($result -notmatch "^\d{3,7}$")              # only accepts three seven digit numbers

"Entered: $result"

###########################################################################################################################################################################################

# Tip 77: Getting Non-Expiring Passwords


function Get-NonExpiringPasswords
{
	$filter = '(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=65536))'
	
    $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")	
    $searcher = New-Object System.DirectoryServices.DirectorySearcher $filter	

    $SearchRoot = $root.defaultNamingContext	
    $searcher.SearchRoot = "LDAP://$SearchRoot"	
    $searcher.SearchScope = 'SubTree'	
    $searcher.SizeLimit = 0	
    $searcher.PageSize = 1000	

    $searcher.FindAll() | Foreach-Object { $_.GetDirectoryEntry() }
}

Get-NonExpiringPasswords                 # use a searcher object like this if you need to find all user accounts in your Active Directory with non-expiring passwords

###########################################################################################################################################################################################

# Tip 78: Control Media Player from PowerShell

# Start-MediaPlayer accepts any one of them and will launch Media Player to play the song or playlist
function Start-MediaPlayer
{
    param([Parameter(Mandatory = $true)]$name)

    $player = New-Object -ComObject WMPLAYER.OCX
    $item = $player.mediaCollection.getByName($name)

    if($item.count -gt 0)
    {
        $filename = $item.item(0).sourceurl
        $player.openPlayer($filename)
    }
    else
    {
        "$name not found"
    }
}

# Get-MediaPlayerItems will return all playlists and multimedia items accessible
function Get-MediaPlayerItems
{
    $player = New-Object -ComObject WMPLAYER.OCX
    $items = $player.mediaCollection.getAll()

    0..($items.count - 1) | ForEach-Object {$items.Item($_).Name}
}

###########################################################################################################################################################################################

# Tip 79: Reading Twitter News

$xml = New-Object XML
$xml.Load("http://search.twitter.com/search.atom?q=powershell&rpp=100")         # to return the latest 100 Twitter entries related to PowerShell
$xml.feed.entry | Select-Object title

###########################################################################################################################################################################################

# Tip 80: Reading Password Age

function Get-PwdAge 
{	
    $filter = '(&(objectCategory=person)(objectClass=user))'	

    $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")	
    $searcher = New-Object System.DirectoryServices.DirectorySearcher $filter	

    $SearchRoot = $root.defaultNamingContext	
    $searcher.SearchRoot = "LDAP://CN=Users,$SearchRoot"	
    $searcher.SearchScope = 'SubTree'	
    $searcher.SizeLimit = 0	
    $searcher.PageSize = 1000	
    $searcher.FindAll() | Foreach-Object { 	
    	
        $account = $_.GetDirectoryEntry()		
        $pwdset = [datetime]::fromfiletime($_.properties.item("pwdLastSet")[0]) 		
        $age = (New-TimeSpan $pwdset).Days				
        $info = 1 | Select-Object Name, Age, LastSet		
        $info.Name = $account.SamAccountName[0]		
        $info.Age = $age		
        $info.LastSet = $pwdset		
        $info	
    }
}



# Example with the ActiveDirectory Module
function Get-PwdAge 
{

 Import-Module ActiveDirectory -ea 0                            # import failed here: Make sure you have ADWS installed and have the service started
 # Reference: http://social.technet.microsoft.com/Forums/windowsserver/zh-CN/25f28c33-43c7-42db-bb9d-9073e6920e66/importmodule-activedirectory-fails

 $users = Get-ADUser -Properties pwdLastSet -Filter *

 $users | % {

   $pwdSet = [DateTime]::FromFileTime($_.pwdLastSet)
   $age = (New-TimeSpan $pwdset).Days
   $info = $true | Select-Object Name, Age, LastSet
   $info.Name = $_.SamAccountName
   $info.Age = $age
   $info.LastSet = $pwdset

   $info
  }
}

###########################################################################################################################################################################################

# Tip 81: Accessing Registry Remote

$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", "IIS-CTI5052")
$key = $reg.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')

$key.GetSubKeyNames() | ForEach-Object {

    $subkey = $key.OpenSubKey($_)

    $i = @{}
    $i.Name = $subkey.GetValue("DisplayName")
    $i.Version = $subkey.GetValue("DisplayVersion")

    New-Object PSObject -Property $i
    $subkey.Close()
}

$key.Close()
$reg.Close()

###########################################################################################################################################################################################

# Tip 82: Get Installed Software

# Get-InstalledSoftware, it will return the locally installed software. By adding a ComputerName, you can then use PowerShell Remote to get the same information from a remote machine
function GetInstalledSoftware($computername = $null)
{
    $code = {
        
        Get-ItemProperty 'hklm:\software\microsoft\windows\currentversion\uninstall\*' | Where-Object {$_.DisplayName} | Select-Object DisplayName, DisplayVersion, Publisher
    }

    if($computername)
    {
        Invoke-Command -ScriptBlock $code -ComputerName $computername | Select-Object DisplayName, DisplayVersion, Publisher
    }
    else
    {
        Invoke-Command -ScriptBlock $code
    }
}

GetInstalledSoftware
GetInstalledSoftware IIS-CTI5052

# Note: You can get a list of installed software right from the registry as long as the target system runs PowerShell v2 
 # and is set up for PowerShell Remote, which also works for remote machines.

###########################################################################################################################################################################################

# Tip 83: Finding 32-Bit Processes

# On a 64-bit machine that not all processes are 64-bit, to filter out only 32-bit processes
Get-Process | Where-Object { ($_ | Select-Object -ExpandProperty Modules -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ModuleName) -Contains "wow64.dll"}

###########################################################################################################################################################################################

# Tip 84: Getting Up-to-Date Exchange Rates

# If you would like to get currency exchange rates, you can simply access a Web service from one of the major banks. Here is how to get the USD exchange rate
$xml = New-Object XML
$xml.Load('http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml')
$rates = $xml.Envelope.Cube.Cube.Cube

"Current USD exchange rate:"

$usd = $rates | Where-Object {$_.currency -eq "USD"} | Select-Object -ExpandProperty rate
$usd

###########################################################################################################################################################################################

# Tip 84: Finding MAC Addresses

Get-WmiObject Win32_NetworkAdapter | Where-Object {$_.MacAddress} | Select-Object Name, MacAddress


# Sending Magic Packet

# You can also send a "magic packet" with the machines MAC address if you'd like to wake up a machine. 
 # Here is how to send such a packet-- just make sure you adjust $mac to the MAC address your machine is using
$mac = [byte[]](0x00, 0x11, 0x85, 0x83, 0xad, 0x3b)

$UDPClient = New-Object System.Net.Sockets.UdpClient
$UDPClient.Connect(([System.Net.IPAddress]::Broadcast), 4000)

$packet = [byte[]](,0xFF * 102)
6..101 | ForEach-Object { $packet[$_] = $mac[($_ % 6)] }

"Send: "
$packet

$UDPClient.Send($packet, $packet.Length)

###########################################################################################################################################################################################

# Tip 85: Finding Out Video Resolution

function Get-Resolution
{
    param(
    [Parameter(ValueFromPipeline = $true)]
    [Alias("cn")]
    $computerName = "."
    )

    process
    {
        Get-WmiObject Win32_VideoController -ComputerName $computerName | Select-Object *resolution*, __Server
    }
}

Get-Resolution
# Output:
#        CurrentHorizontalResolution        CurrentVerticalResolution         __SERVER                                                              
#        ---------------------------        -------------------------         --------                                                              
#                               1600                              900         IIS-V-SIHE-01

# Note: When you run Get-Resolution, you will retrieve your own video resolution. By adding a computer name, 
 # you will find the same information is returned from a remote system (provided you have sufficient privileges to access the remote system).

###########################################################################################################################################################################################

# Tip 86: Storing a Picture in Active Directory

# When you need to store a picture into an AD account, the picture will have to be converted to byte values before it can be stored. 
 # Just make sure you adjust the path to the picture you want to store and the LDAP path of the AD object you want the picture to be stored in

# Method One:
$file = "C:\pic.jpg"

$bild = New-Object System.Drawing.Bitmap($file)
$ms   = New-Object System.IO.MemoryStream

$bild.Save($ms, "jpeg")
$ms.Flush()

$byte = $ms.ToArray()

$user = New-Object System.DirectoryServices.DirectoryEntry ("LDAP://10.17.141.219/CN=Tobias,CN=Users,DC=powershell,DC=local")
$user.Properties["jpegPhoto"].Value = $byte
$user.SetInfo()

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Method Two:
$byte = Get-Content -Encoding Byte $file

$user = New-Object System.DirectoryServices.DirectoryEntry ("LDAP://10.17.141.219/CN=Tobias,CN=Users,DC=powershell,DC=local")
$user.Properties["jpegPhoto"].Value = $byte
$user.SetInfo()

###########################################################################################################################################################################################

# Tip 87: Print All PDF Files in a Folder

dir $env:windir\*.pdf | ForEach-Object { Start-Process -FilePath $_.FullName -Verb Print }        # to print out all PDF documents you have stored in one folder

###########################################################################################################################################################################################

# Tip 88: Determining Service Start Mode

Get-WmiObject Win32_Service | Select-Object Name, StartMode                                       # To get a list of all services

([wmi]"Win32_Service.Name='Spooler'").StartMode                                                   # to find out the start mode of one specific service


# Changing Service Startmode

# Method One:
([wmi]"Win32_Service.Name='Spooler'").ChangeStartMode("Automatic").ReturnValue
([wmi]"Win32_Service.Name='Spooler'").StartMode 


([wmi]"Win32_Service.Name='Spooler'").ChangeStartMode("Manual").ReturnValue
([wmi]"Win32_Service.Name='Spooler'").StartMode 



# Method Two:
Get-Service spooler | Set-Service -StartupType Automatic

# Note that a return value of 0 indicates success. You will need Administrator privileges to change the start mode.

###########################################################################################################################################################################################

# Tip 89: Out-GridView Requirements

Get-Process | Out-GridView                        # Out-GridView is a great way to present results in a “mini-Excel” sheet

# However, Out-GridView has two requirements:.NET Framework 3.51 and the built-in script editor ISE must both be installed. 
 # ISE is not installed by default on Windows Servers. So, if you want  to use Out-GridView on server products, you will need to make sure you install the ISE feature.


# On a Server 2008 R2, you could enable ISE by using PowerShell
Import-Module ServerManager
Add-WindowsFeature PowerShell-ISE

###########################################################################################################################################################################################

# Tip 90: Writing Registry Key Default Value

# If you need to set the default value for a registry key, you can  use either of these approaches:

Set-ItemProperty -Path HKCU:\Software\Somekey -Name "(Default)" -Value MyValue

Set-Item -Path HKCU:\Software\SomeKey -Value MyValue

###########################################################################################################################################################################################

# Tip 91: List Registry Hives

dir Registry::                                               # to get a list of all registry hives

###########################################################################################################################################################################################

# Tip 92: Analyzing Windows Launch Time

function Get-WindowsLaunch
{
    $filter = @{ logname = 'Microsoft-Windows-Diagnostics-Performance/Operational'; id = 100 }

    Get-WinEvent -FilterHashtable $filter | ForEach-Object {
        
        $info = 1 | Select-Object Date, Startduration, Autostarts, Logonduration

        $info.Date = $_.Properties[1].Value
        $info.Startduration = $_.Properties[5].Value
        $info.Autostarts = $_.Properties[18].Value
        $info.Logonduration = $_.Properties[43].Value

        $info
    }
}

Get-WindowsLaunch

Get-WindowsLaunch | Measure-Object StartDuration -Minimum -Maximum -Average




# Finding Software Updates

# The function Get-SoftwareUpdates will create a list of software updates that your machine received.  
 # While this does not cover the usual Hotfixes (use Get-Hotfix for those), the list will clearly show which installed software packages received updates
function Get-SoftwareUpdates
{
    $filter = @{ logname = 'Microsoft-Windows-Application-Experience/Program-Inventory'; id = 905}

    Get-WinEvent -FilterHashtable $filter | ForEach-Object {
    
        $info = 1 | Select-Object Date, Application, Version, Publisher

        $info.Date = $_.TimeCreated
        $info.Application = $_.Properties[0].Value
        $info.Version = $_.Properties[1].Value
        $info.Publisher = $_.Properties[2].Value

        $info
    }
}

Get-SoftwareUpdates


# In Windows Vista/Server 2008, Microsoft introduced many new service and application specific log files. PowerShell can access those with Get-WinEvent.

###########################################################################################################################################################################################

# Tip 93: Output Data in Color

Get-Process | Write-Host -ForegroundColor Yellow                      # The result is colorized, but Write-Host has converted the processes into very simplistic string representations

Get-Process | Out-String -Stream | Write-Host -ForegroundColor Yellow # convert objects to string manually with Out-String to use the same rich conversion you normally see in the console

###########################################################################################################################################################################################

# Tip 94: Filter PowerShell Results Fast and Text-Based

filter grep($keyword)
{
    if(($_ | Out-String) -like "*$keyword*")
    {
        $_
    }
}

Get-Service | grep running
dir $env:windir | grep .exe
dir $env:windir | grep 14.07.2009
Get-Alias | grep child

# As you can see, while the filtering is based on simple plain text keywords, the results are still rich objects!

###########################################################################################################################################################################################

# Tip 95: Use Hash Tables To Convert Numeric Return Values

$cleartext = @{ 0 = "success"; 5 = "access denied" }

$rv = (dir).count                                         # You can use the return value as an index to get back the clear text representation of a numeric return value
$cleartext[$rv]

###########################################################################################################################################################################################

# Tip 96: Split Special Characters

"1,2,3,4" -split ","                                             # PowerShell’s new –split operator can split text into parts.             

# However, you can also submit a script block and calculate where to split. If you take advantage of the many specific character tests available in the [Char] type, 
 # you can then split based on punctuation or other character groups
"Hello, this is a test" -split { [char]::IsPunctuation($_) }


[char] | Get-Member -Static -MemberType Method -Name is*         # to view all available character groups
# Output:
#        Name            MemberType Definition                                                                                                          
#        ----            ---------- ----------                                                                                                          
#        IsControl       Method     static bool IsControl(char c), static bool IsControl(string s, int index)                                           
#        IsDigit         Method     static bool IsDigit(char c), static bool IsDigit(string s, int index)                                               
#        IsHighSurrogate Method     static bool IsHighSurrogate(char c), static bool IsHighSurrogate(string s, int index)                               
#        IsLetter        Method     static bool IsLetter(char c), static bool IsLetter(string s, int index)                                             
#        IsLetterOrDigit Method     static bool IsLetterOrDigit(char c), static bool IsLetterOrDigit(string s, int index)                               
#        IsLower         Method     static bool IsLower(char c), static bool IsLower(string s, int index)                                               
#        IsLowSurrogate  Method     static bool IsLowSurrogate(char c), static bool IsLowSurrogate(string s, int index)                                 
#        IsNumber        Method     static bool IsNumber(char c), static bool IsNumber(string s, int index)                                             
#        IsPunctuation   Method     static bool IsPunctuation(char c), static bool IsPunctuation(string s, int index)                                   
#        IsSeparator     Method     static bool IsSeparator(char c), static bool IsSeparator(string s, int index)                                       
#        IsSurrogate     Method     static bool IsSurrogate(char c), static bool IsSurrogate(string s, int index)                                       
#        IsSurrogatePair Method     static bool IsSurrogatePair(string s, int index), static bool IsSurrogatePair(char highSurrogate, char lowSurrogate)
#        IsSymbol        Method     static bool IsSymbol(char c), static bool IsSymbol(string s, int index)                                             
#        IsUpper         Method     static bool IsUpper(char c), static bool IsUpper(string s, int index)                                               
#        IsWhiteSpace    Method     static bool IsWhiteSpace(char c), static bool IsWhiteSpace(string s, int index)                                     


[char]::IsPunctuation
# Output:
#        OverloadDefinitions                                                                                                                                                                                                  
#        -------------------                                                                                                                                                                                                  
#        static bool IsPunctuation(char c)                                                                                                                                                                                    
#        static bool IsPunctuation(string s, int index) 

###########################################################################################################################################################################################

# Tip 97: Encrypt Files With EFS

$file = Get-Item $home\test.txt    # access the file using Get-Item to encrypt a file with EFS

$file.Encrypt()                    # Provided that EFS is available on your system, the file will be  encrypted, and in Windows Explorer, the file will now get a green label

$file.Decrypt()                    # to undo encryption

###########################################################################################################################################################################################

# Tip 98: Filter by More Than One Criteria

dir $env:windir -Filter *.log -Recurse -ErrorAction SilentlyContinue         # use Get-Childitem with its -filter parameter to list specific files because it is much faster than -include

dir $env:windir -Recurse -Include *.wav, *.bmp -ErrorAction SilentlyContinue # Use -include instead of -filter. Although it is slower, it is  also more versatile

###########################################################################################################################################################################################

# Tip 99: Use -include instead of -filter. Although it is slower, it is  also more versatile

function Open-File([Parameter(Mandatory = $true)]$path)
{
    $path = Resolve-Path $path -ErrorAction SilentlyContinue

    if($path -ne $null)
    {
        $path | ForEach-Object { Invoke-Item $_ }
    }
    else
    {
        "No file matched $path ."
    }
}

Open-File $home\*.txt

###########################################################################################################################################################################################

# Tip 100: Use Multiple Wildcards

Resolve-Path $home\*\*\*.dll -ErrorAction SilentlyContinue     # find all DLL-files in all sub-folders up to two levels below your Windows folder

dir C:\Users\*\Desktop\*                                       # lists the Desktop folder content for all user accounts on your computer-provided that you have sufficient privileges

###########################################################################################################################################################################################