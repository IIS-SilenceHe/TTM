# Reference site: http://powershell.com/cs/blogs/tips/
###########################################################################################################################################################################################

# Tip 1: Use Splatting to Encapsulate WMI Calls


# Splatting is a great way of forwarding parameters to another cmdlet. 
# Here is an example that can be used to encapsulate WMI calls and make them available under different names:

function Get-BIOSInfo
{
    param
    (
        $ComputerName,
        $Credential,
        $SomethingElse
    )

    $null = $PSBoundParameters.Remove("SomethingElse")

    Get-WmiObject -Class Win32_BIOS @PSBoundParameters
}

Get-BIOSInfo

# Get-BIOSInfo gets BIOS information from WMI, and it works locally, remotely and remotely with credentials. 
# This is possible because only the parameters that a user actually submits to Get-BIOSInfo are forwarded to the same parameters in Get-WmiObject.
# So when a user does not submit -Credential, then no -Credential parameter is submitted to Get-WmiObject either.

# Splatting typically uses a self-defined hash table where each key represents a parameter, and each value is the argument assigned to that parameter. 
# In this example, a predefined hash table called $PSBoundParameters is used instead. It is prefilled with the parameters submitted to the function.

# Just make sure you do not forward parameters that are unknown to the destination cmdlet. 
# To illustrate this, the function Get-BIOSInfo defines a parameter called "SomethingElse". 
# Get-WmiObject does not have such a parameter, so before you can splat, you must call Remove() method to remove this key from the hash table.

###########################################################################################################################################################################################

# Tip 2: Getting Database Connection String


# Have you ever been puzzled just what the connection string would look like for a given database? 
# When you create a new data source in Control Panel, a wizard guides you through the creation process. 
# Here is a way to utilize this wizard and get back the resulting connection string.

# Note that the wizard choices depend on the installed database drivers on your machine.

function Get-ConnectionString
{
    $path = Join-Path -Path $env:TEMP -ChildPath "dummy.dul"
    $null = New-Item -Path $path -ItemType File -Force
    $commandArg = """$env:CommonProgramFiles\System\OLE DB\oledb32.dll"",OpenDSLFile " + $Path

    Start-Process -FilePath Rundll32.exe -ArgumentList $commandArg -Wait

    $connectionString = Get-Content -Path $path | Select-Object -Last 1
    $connectionString | clip.exe

    Write-Warning 'Connection String is also available from clipboard'
    
    $connectionString
}

Get-ConnectionString

# When you call Get-ConnectionString, a dummy udl file is created and opened by the Control Panel wizard. 
# You can then walk through the wizard. Once done, 
# PowerShell examines the results in the dummy file and returns the connection string to you.

# This is possible because Get-Process uses -Wait, effectively halting the script until the wizard exists. 
# At that point, the script can safely examine the udl file.

###########################################################################################################################################################################################

# Tip 3: gpupdate on Remote Machines


# To run gpupdate.exe remotely, you could use a script like this:

function Start-GPUpdate
{
    param([String[]]$ComputerName)

    $code = {   
      
        $rv = 1 | Select-Object -Property ComputerName, ExitCode

        $null = gpupdate.exe /force

        $rv.Exitcode = $LASTEXITCODE
        $rv.ComputerName = $env:COMPUTERNAME

        $rv  
    }

    Invoke-Command -ScriptBlock $code -ComputerName $ComputerName | Select-Object -Property ComputerName, ExitCode
} 

Start-GPUpdate -ComputerName iis-cti5052
# Output:
#         ComputerName           ExitCode
#         ------------           --------
#         Server01               0


# Start-GPUpdate accepts one or more computer names and will then run gpupdate.exe on all of them. The result is transferred back to you.

# This script takes advantage of PowerShell remoting, so it does require PowerShell remoting to be enabled on target machines, 
# and you need to have local Admin privileges on these machines.

###########################################################################################################################################################################################

# Tip 4: Reading Installed Software Remotely


# Most software registers itself in the Registry. 
# Here is a piece of code that reads all installed software from the 32-bit and 64-bit hive and works locally and remotely as well. 
# It can serve as a good example on how to remotely read Registry keys, too.


$Hive = "LocalMachine"

$key = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'

$value = 'DisplayName', 'DisplayVersion', 'UninstallString'

$ComputerName = $env:COMPUTERNAME

# add the value "RegPath" which will contain the actual Registry path the value came from
$value = @($value) + "RegPath"

$key | ForEach-Object {

    $RegHive = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $ComputerName)

    $regKey = $RegHive.OpenSubKey($_)

    $regKey.GetSubKeyNames() | ForEach-Object {
    
        $subKey = $regKey.OpenSubKey($_)

        $returnValue = 1 | Select-Object -Property $value

        $value | ForEach-Object {
        
            $returnValue.$_ = $subKey.GetValue($_)
        }

        $returnValue.RegPath = $subKey.Name

        $returnValue

        $subKey.Close()
    }

    $regKey.Close()
    $RegHive.Close()

} | Out-GridView

###########################################################################################################################################################################################

# Tip 5: Getting DateTaken Info from Pictures


# If you'd like to reorganize your picture archive, then here is a piece of code that reads the "DateTaken" information from picture files.

# The example uses a system function to find out the MyPictures path, then recursively searches for all files in that folder or its subfolders. 
# The result is piped to Get-DateTaken which returns the picture file name, the folder name, and the date the picture was taken.

function Get-DateTaken
{
    param
    (
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias("FullName")]
        [String]
        $Path
    )

    begin { $shell = New-Object -ComObject Shell.Application }

    process
    {
        $returnValue = 1 | Select-Object -Property Name, DateTaken, Folder

        $returnValue.Name = Split-Path $Path -Leaf
        $returnValue.Folder = Split-Path $Path

        $shellfolder = $shell.NameSpace($returnValue.Folder)
        $shellfile = $shellfolder.ParseName($returnValue.Name)
        $returnValue.DateTaken = $shellfolder.GetDetailsOf($shellfile, 12)

        $returnValue
    }
}

$picturePath = [System.Environment]::GetFolderPath("MyPictures")
Get-ChildItem -Path $picturePath -Recurse -ErrorAction SilentlyContinue | Get-DateTaken

# Output:
#        Name                                          DateTaken              Folder                                                                
#        ----                                          ---------              ------                                                                
#        10.jpg                                                               C:\Users\v-sihe\Pictures                                                                                           
#        62f87eb4jw1ee2eze41axj207q0ki40c.jpg                                 C:\Users\v-sihe\Pictures                                              
#        6628711bjw1edeb470j3yj20c8096dhb.jpg                                 C:\Users\v-sihe\Pictures                                              
#        Cedar%2527s%252BStars.jpg                                            C:\Users\v-sihe\Pictures                                              
#        DSC_3970.jpg                                                         C:\Users\v-sihe\Pictures                                              
#        iis_logo.jpg                                                         C:\Users\v-sihe\Pictures                                              
#        ok.jpg                                        ‎2008/‎10/‎23 ‏‎05:12 PM    C:\Users\v-sihe\Pictures                                              
#        SAM_1413.JPG                                  ‎2012/‎10/‎28 ‏‎08:53 AM    C:\Users\v-sihe\Pictures                                               

###########################################################################################################################################################################################

# Tip 6: Bulk File Renaming


# Let's assume you have a bunch of scripts (or pictures or log files or whatever) in a folder, and you'd like to rename all files. 
# The new file name should, for example, have a prefix, then an incrementing number.

# This example would rename all PowerShell scripts with the extension .ps1 inside the folder you specified. 
# The new name would be powershellscriptX.ps1 where "X" is an incrementing number.

# Note that the actual rename is disabled in this script. Remove the -WhatIf parameter to actually rename the files, 
# but be extremely careful. If you mistype a variable or use the wrong folder path, 
# then your script might happily rename thousands of wrong files that you did not intend to rename.

$path = "C:\temp"
$filter = "*.ps1"
$prefix = "powershellscript"
$counter = 1

Get-ChildItem -Path -Filter $filter -Recurse | Rename-Item -NewName {

    $extension = [System.IO.Path]::GetExtension($_.Name)

    "{0}{1}.{2}" -f $prefix, $Script:Counter, $extension

    $Script:Counter++

} -WhatIf

###########################################################################################################################################################################################

# Tip 7: Be Aware of Side Effects


# There are plenty of low level system functions that PowerShell can use. 
# This one, for example, creates a temporary file name:

[System.IO.Path]::GetTempFileName()

# However, it does not only do that. It also actually creates the file. So if you use this function to create temporary file names, 
# you might end up with a lot of orphaned files in your file system. Use it only if you actually want a temporary file to be created.

###########################################################################################################################################################################################

# Tip 8: Using Profile Scripts


# You probably know that PowerShell supports profile scripts. Simply make sure the file found in $profile exists. 
# It's a plain script that gets executed each time the PowerShell host launches.

# So it's easy to configure your PowerShell environment, load modules, add snap-ins, and do other adjustments.
# This would shorten your PowerShell prompt and instead display the current location in the title bar:

function prompt
{
    "PS> "
    $host.UI.RawUI.WindowTitle = Get-Location
}

# Note also that the profile script found in $profile is host specific. 
# It works only for a given host (either the PowerShell console, or the ISE editor, or whatever else PowerShell host you are using).

# To execute code automatically on launch of any host, use this file instead:
$profile.CurrentUserAllHosts

# It is basically the same path, except the file name is now not using a host name anymore. 
# Instead, the file name is called "profile.ps1" only.


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Skipping Profile on Keystroke

# Maybe you'd like to be able to occasionally skip certain parts of your profile script.
# For example, in the ISE editor, simply add this construction to your profile script 
# (path to your profile script is found in $profile, it may not yet exist):

if([System.Windows.Input.Keyboard]::IsKeyDown("Ctrl")){ return }
# This will skip all remaining lines in your profile script if you hold CTRL while launching the ISE editor.

# Or, you can use it like this:

if([System.Windows.Input.Keyboard]::IsKeyDown('Ctrl') -eq $false) 
{ 
    Write-Warning 'You DID NOT press CTRL, so I could execute things here.'
}

# This would run the code in the braces section only if you did not hold down CTRL while launching the ISE.

# If you want to use the code in the PowerShell console, too, then make sure you load the appropriate assembly. 

# This would work in all profile files:

Add-Type -AssemblyName PresentationFramework
if([System.Windows.Input.Keyboard]::IsKeyDown('Ctrl') -eq $false) 
{ 
    Write-Warning 'You DID NOT press CTRL, so I could execute things here.'
}

###########################################################################################################################################################################################

# Tip 9: Handling Cmdlet Errors without Interruption


# When you want to use error handlers with errors that occur inside a cmdlet, 
# then you can only catch such exceptions when you set the -ErrorAction of that cmdlet to "Stop". 
# Else, the cmdlet handles the error internally.

# That's bad because setting -ErrorAction to "Stop" will also stop the cmdlet at the first error.

# So if you want to not interrupt a cmdlet but still get all errors caused by the cmdlet, then use -ErrorVariable instead. 
# This line gets all PowerShell scripts recursively inside your Windows folder (which can take some time).
# Errors are suppressed but logged to a variable:

Get-ChildItem -Path c:\Windows -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue -ErrorVariable myErrors

# When the cmdlet is done, you can examine the variable $myErrors. It contains all the errors that occurred. 
# This would give you a list of subfolders, for example, that Get-ChildItem was unable to look into:
$myErrors.TargetObject


# It uses automatic unrolling (introduced in PowerShell 3.0). So in PowerShell 2.0, you'd have to write:
$myErrors | Select-Object -ExpandProperty TargetObject

###########################################################################################################################################################################################

# Tip 10: Reading Registry Values the Easy Way


# With PowerShell, it can be a piece of cake to read out Registry values. Here is your simple code template:

$regPath = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion'
$key = Get-ItemProperty -Path "Registry::$regPath"

# Now, simply replace the value of $RegPath with any Registry key path. 
# You can even copy and paste the key path from regedit.exe.

# Once you run the code and $key is filled, simply enter $key and a dot. 
# In the ISE editor, IntelliSense will show all value names of that key, and you can simply pick the ones you want to read. 
# In the console, press TAB after you entered the "." to see the available value names:

$key.CommonFilesDir         # Output: C:\Program Files\Common Files
$key.MediaPathUnexpanded    # Output: C:\Windows\Media
$key.ProgramW6432Dir        # Output: C:\Program Files

###########################################################################################################################################################################################

# Tip 11: Speeding Up Arrays


# When you assign new items to an array often, you may experience a performance problem.
# Here is a sample that illustrates how you should not do it:

Measure-Command {

    $array = @()

    for($i = 1; $i -lt 10000; $i++)
    {
        $array += $i
    }
}                                          # TotalSeconds      : 7.6956287

# In a loop, an array receives a lot of new items using the operator "+=". 
# This takes a long time, because PowerShell needs to create a new array each time you change its size.



# Here is an approach that is many, many times faster--using an ArrayList which is optimized for size changes:

Measure-Command {

    $array = New-Object -TypeName System.Collections.ArrayList
   
    for($i = 1; $i -lt 10000; $i++)
    {
        $array.Add($i) | Out-Null          # void returns the index at which the new element was added
    }   
}                                          # TotalSeconds      : 0.0815746

# Both pieces of code achieve the same thing. The second approach is a lot faster.



# From Comments:

Measure-Command {

    $array = @(
    
        for($i = 1; $i -lt 10000; $i++)
        {
            $i
        }
    )
}                                          # TotalSeconds      : 0.0576392

###########################################################################################################################################################################################

# Tip 12: Using Nested Hash Tables


# Nested hash tables can be a great alternative to multidimensional arrays. 
# They can be used to store data sets in an easy-to-manage way. Have a look:

$person = @{}

$person.Name = "Silence"
$person.Id = 25

$person.Address = @{}
$person.Address.Street = "Wujing"
$person.Address.City = "Shanghai"

$person.Address.Details = @{}
$person.Address.Details.Story = 4
$person.Address.Details.ScenicView = $false

# This would define a person. You can always view the entire person:

$person
# Output:
#        Name                           Value                                                                                                                                                                                 
#        ----                           -----                                                                                                                                                                                 
#        Name                           Silence                                                                                                                                                                               
#        Id                             25                                                                                                                                                                                    
#        Address                        {Street, Details, City}


# You can just as easily retrieve individual pieces of information:

$person.Address.City                    # Output: Shanghai

$person.Address.Details.Story           # Output: 4



# From Comments:

$person = @{

    Name = "Silence"
    Id = 25

    Address = @{
    
        Street = "Wujing"
        City = "Shanghai"

        Details = @{
        
            Story = 4
            ScenicView = $false
        }
    }
}

$person

###########################################################################################################################################################################################

# Tip 13: Dealing with Environment Variables


$env:windir            # Output: C:\Windows
$env:USERNAME          # Output: sihe

# Actually, "env:" is a drive, so you can use it to find all (or some) environment variables. 
# This would list all environment variables that have "user" in its name:

dir env:\*user*

# Output:
#        Name                           Value                                                                                                                                                                                 
#        ----                           -----                                                                                                                                                                                 
#        USERDNSDOMAIN                  FAREAST.CORP.MICROSOFT.COM                                                                                                                                                            
#        USERPROFILE                    C:\Users\v-sihe                                                                                                                                                                       
#        ALLUSERSPROFILE                C:\ProgramData                                                                                                                                                                        
#        USERNAME                       v-sihe                                                                                                                                                                                
#        USERDOMAIN                     FAREAST                                                                                                                                                                               
#        GIT_USERNAME                   silence                                                                                                                                                                               
#        USERDOMAIN_ROAMINGPROFILE      FAREAST



dir env:      # list all environment variables

###########################################################################################################################################################################################

# Tip 14: Speeding Up Background Jobs


# Background jobs can be a great thing to speed up scripts because they can do things in parallel. 
# However, background jobs only work well if the code you run does not produce large amounts of data 
# - because transporting back the data via XML serialization often takes more time than you can save by executing things in parallel.

$start = Get-Date

$code1 = { Get-Hotfix }
$code2 = { Get-ChildItem $env:windir\system32\*.dll }
$code3 = { Get-Content -Path C:\Windows\WindowsUpdate.log }

$job1 = Start-Job -ScriptBlock $code1 
$job2 = Start-Job -ScriptBlock $code2
$result3 = & $code3 

$alljobs = Wait-Job $job1, $job2 

Remove-Job -Job $alljobs
$result1, $result2 = Receive-Job $alljobs

$end = Get-Date

$timespan = $end - $start
$seconds = $timespan.TotalSeconds
Write-Host "This took me $seconds seconds."

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


$start = Get-Date

$result1 = Get-Hotfix 
$result2 = Get-ChildItem $env:windir\system32\*.dll 
$result3 = Get-Content -Path C:\Windows\WindowsUpdate.log 

$end = Get-Date
$timespan = $end - $start
$seconds = $timespan.TotalSeconds
Write-Host "This took me $seconds seconds."

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# So background jobs really have made the code more complex and increased the script runtime. 
# Only when you start and optimize the return data will background jobs become useful. 
# The fewer data they emit, the better.

$start = Get-Date

$code1 = { Get-Hotfix | Select-Object -ExpandProperty HotfixID }
$code2 = { Get-Content -Path C:\Windows\WindowsUpdate.log | Where-Object { $_ -like '*successfully installed*' }}
$code3 = { Get-ChildItem $env:windir\system32\*.dll | Select-Object -ExpandProperty Name }

$job1 = Start-Job -ScriptBlock $code1 
$job2 = Start-Job -ScriptBlock $code2
$result3 = & $code3 

$alljobs = Wait-Job $job1, $job2 

Remove-Job -Job $alljobs
$result1, $result2 = Receive-Job $alljob 

# This time, the background jobs only return as much data as is really needed, 
# and the job with the most output is moved to the foreground PowerShell. 
# This will cut down execution time considerably.

# Generally, background jobs work best when they simply do something (for example a configuration)
# but do not return anything or return just a small amount of data.

###########################################################################################################################################################################################

# Tip 15: Finding Working Days


# To find all working days in a given month, here is a neat little one-liner:

$month = 12

1..31 | ForEach-Object { Get-Date -Day $_ -Month $month } | Where-Object { $_.DayOfWeek -gt 0 -and $_.DayOfWeek -lt 6 }


# With a couple of more commands, the pipeline returns the number of working days as a number, too:

1..31 | ForEach-Object { Get-Date -Day $_ -Month $month } | Where-Object { $_.DayOfWeek -gt 0 -and $_.DayOfWeek -lt 6 } | 
    Measure-Object | Select-Object -ExpandProperty Count                                                                    # Output: 23




$month = 12
$year = 2014

1..[DateTime]::DaysInMonth($year,$month) | ForEach-Object { Get-Date -Day $_ -Month $month -Year $year } | 
    Where-Object { $_.DayOfWeek -gt 0 -and $_.DayOfWeek -lt 6 }

###########################################################################################################################################################################################

# Tip 16: Speeding Up Scripts with StringBuilder


# Often, scripts add new text to existing text. Here is a piece of code that may look familiar to you:

Measure-Command {

    $text = "Hello"

    for($i = 0; $i -lt 100000; $i++)
    {
        $text += "status $x"
    }

    $text
}                                         # TotalSeconds      : 63.717262 


# This code is particularly slow because whenever you add text to a string, the complete string needs to be recreated. 
# There is, however, a specialized object called StringBuilder. It can do the very same, but at lightning speed:

Measure-Command {

    $sb = New-Object -TypeName System.Text.StringBuilder
    $null = $sb.Append("hello")

    for($i = 0; $i -lt 100000; $i++)
    {
        $null = $sb.Append("status $x")
    }

    $sb.ToString()
}                                          # TotalSeconds      : 0.5231941

###########################################################################################################################################################################################

# Tip 17: Using Default Parameters


# In PowerShell 3.0, an option was added to define default values for arbitrary cmdlet parameters.

# This line, for example, would set the default value for the parameter -Path of all cmdlets to a given path:

$PSDefaultParameterValues.Add("*:Path", "C:\Windows")


# So when you now run Get-ChildItem or any other cmdlet that has a parameter -Path, 
# it behaves as if you had specified the given path for this parameter.



# Instead of the "*", you can of course add the name of a specific cmdlet. 
# So if you wanted to set the parameter -ComputerName for Get-WmiObject to a specific remote system, this is how you would do that:

$PSDefaultParameterValues.Add("Get-WmiObject:ComputerName", "server12")


# All of these defaults are valid only in the current PowerShell session. 
# If you want to keep them, then simply define the default values in one of your profile scripts.



# To remove all custom default values again, use this:

$PSDefaultParameterValues.Clear()

###########################################################################################################################################################################################

# Tip 18: Finding Dates Between Two Dates


# If you must know how many days are between two dates, you can easily find out by using New-TimeSpan:

$start = Get-Date                                  # Output: Monday, December 01, 2014 09:50:29
$end = Get-Date -Date "2015/1/1"                   # Output: Thursday, January 01, 2015 00:00:00

$timespan = New-TimeSpan -Start $start -End $end
# Output:
#        Days              : 30
#        Hours             : 14
#        Minutes           : 9
#        Seconds           : 30
#        Milliseconds      : 932
#        Ticks             : 26429709323892
#        TotalDays         : 30.5899413470972
#        TotalHours        : 734.158592330333
#        TotalMinutes      : 44049.51553982
#        TotalSeconds      : 2642970.9323892

$timespan.Days                                     # Ouput: 30



# However, if you do not just want to know how many days there are, but if you actually need the exact days, here is another approach:

$days = [Math]::Ceiling($timespan.TotalDays) + 1   # Output: 32

1..$days | ForEach-Object {

    $start

    $start = $start.AddDays(1)
}

# This time, PowerShell outputs all dates between the two specified dates.



# Since you now know the exact dates (and not just the number of days), you can filter, for example by weekday name, 
# and find out how many Sundays or how many work days there are before you can go on vacation or retire.

$days = [Math]::Ceiling($timespan.TotalDays) + 1   # Output: 32

1..$days | ForEach-Object {

    $start

    $start = $start.AddDays(1)

} | Where-Object { $_.DayOfWeek -gt 0 -and $_.DayOfWeek -lt 6}

###########################################################################################################################################################################################

# Tip 19: Copying Command History



# If you played with PowerShell and suddenly notice that some of the lines of code you just entered actually work, 
# then you may want to copy and paste them into a script editor, save them, or show them to friends.

Get-History -Count 5 | Select-Object -ExpandProperty CommandLine | clip.exe

# This will copy the last five commands you entered to the clipboard.


(Get-History).CommandLine | clip.exe  # It copies all commands from your command history to the clipboard

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Copying Command History as a Tool


# You can copy the previously entered interactive PowerShell commands to your favorite script editor. 
# Here is a function that makes this even easier. If you like it, you may want to put it into your profile script so you have it handy any time:

function Get-MyGeniusInput
{
    param($Count, $Minutes = 10000)

    $cutoff = (Get-Date).AddMinutes(- $Minutes)

    $null = $PSDefaultParameterValues.Remove("Minutes")

    $result = Get-History @PSBoundParameters | Where-Object { $_.StartExecutionTime -gt $cutoff } | Select-Object -ExpandProperty CommandLine

    $count = $result.Count
    $result | clip.exe

    Write-Warning "Copied $count command lines to the clipboard!"
}

# Get-MyGeniusInput by default copies the entire command history to the clipboard. 
# With the parameter -Count you can limit the results to a given number, for example the last 5 commands. 
# And with the parameter -Minute, you can specify the number of minutes that you would like to go back in history.

Get-MyGeniusInput -Minutes 25
Get-MyGeniusInput -Minutes 25 -Count 5

###########################################################################################################################################################################################

# Tip 20: Accepting Multiple Input


# When you create PowerShell functions, here is a template that defines a InputObject parameter that will accept multiple values both via parameter and via pipeline:

function Get-Something
{
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Object[]]
        $InputObject
    )

    process
    {
        $InputObject | ForEach-Object {
        
            $element = $_

            "processing $element"
        }
    }
}

Get-Something -InputObject 1,2,3,4
# Output:
#        processing 1
#        processing 2
#        processing 3
#        processing 4
        
1,2,3,4 | Get-Something
# Output:
#        processing 1
#        processing 2
#        processing 3
#        processing 4

# Note how the parameter is defined as an object array (so it can accept multiple values). 
# Next, the parameter value runs through ForEach-Object to process each element individually. 
# This takes care of the first example call: assigning comma-separated (multiple) values.

# To be able to accept multiple values via pipeline, make sure you assign ValueFromPipeline to the parameter that is to accept pipeline input. 
# Next, add a Process script block to your function. It serves as a loop, very similar to ForEach-Object, and runs for each incoming pipeline item.

###########################################################################################################################################################################################

# Tip 21: Creating Great Reports


# You can change all properties of objects when you clone them. 
# Cloning objects can be done to “detach” the object data from the underlying real object and is a great idea. 
# Once you cloned objects, you can do whatever you want with the object, for example, change or adjust its properties.

# To clone objects, run them through Select-Object. That is all.

# This example takes a folder listing, runs it through Select-Object, and then prettifies some of the data:

Get-ChildItem -Path C:\Windows | Select-Object -Property LastWriteTime, "Age(days)", Length, Name, PSIsContainer | ForEach-Object {

    $_."Age(days)" = (New-TimeSpan -Start $_.LastWriteTime).Days

    if($_.PSIsContainer -eq $false)
    {
        $_.Length = ("{0:N1} MB" -f ($_.Length / 1MB))
    }

    $_

} | Select-Object -Property LastWriteTime, "Age(days)", Length, Name | Sort-Object -Property LastWriteTime -Descending | Out-GridView

# The result shows file size in MB rather than bytes, and a new column called “Age(days)” with the file and folder age in days.

###########################################################################################################################################################################################

# Tip 22: Finding AD Accounts Easily


# You do not necessarily need additional cmdlets to search for user accounts or computers in your Active Directory. 
# Provided you are logged on to the domain, simply use this:

$ldap = "(&(objectClass=Computer)(sanAccount=iis*))"
$searcher = [ADSISearcher]$ldap
$searcher.FindAll()

# This would find all computer accounts that start with “dc”. 
# $ldap can be any valid LDAP query. To find users, replace “computer” by “user”.

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Finding AD Users

# Searching the AD can be done with simple calls provided you are logged on an Active Directory domain. 
# Here is an extension that allows you to define a search root (starting point for your search), 
# as well as a flat search (rather than recursing into a container).

$SAMAccountName = $env:USERNAME
$SearchRoot = 'LDAP://OU=customer,DC=company,DC=com'
$SearchScope = 'OneLevel'

$ldap = "(&(objectClass=user)(samAccountName=*$SAMAccountName*))"
$searcher = [adsisearcher]$ldap
$searcher.SearchRoot = $SearchRoot
$searcher.PageSize = 999
$searcher.SearchScope = $SearchScope

$searcher.FindAll() | ForEach-Object { $_.GetDirectoryEntry()  } | Select-Object -Property *

###########################################################################################################################################################################################

# Tip 23: Delete Aliases


# While you can easily create new aliases with New-Alias or Set-Alias, there is no cmdlet to delete aliases.

Set-Alias -Name devicemanager -Value devmgmt.msc
devicemanager


# To delete an alias, you would typically restart your PowerShell. Alternatively, you can delete them using the Alias: drive:

del Alias:\devicemanager
devicemanager

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Test-Driving Scripts without Aliases


# Aliases can be cool in interactive PowerShell but should not be used in scripts. 
# In scripts, use the underlying commands (so use “Get-ChildItem” instead of “dir” or “ls”).

# To test drive a script, you can delete all aliases and then try and see if your script still runs. 
# This is how you’d delete all aliases for the particular PowerShell session
# (it won’t affect other PowerShell sessions and will not delete built-in aliases permanently).

Get-Alias | ForEach-Object { Remove-Item -Path ("Alias:\" + $_.Name) -Force }

# or

Get-Item Alias:* | Remove-Item -Force

# As you see, all aliases are cleaned out, and if a script now uses an alias, it will raise an exception. 
# Once you close and restart PowerShell, all built-in aliases are restored.


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Aliases Can Be Dangerous


# Aliases enjoy the highest priority among executable commands in PowerShell, 
# so if you have ambiguous commands, PowerShell always picks the alias.

# This can be dangerous: if you allow others to change your PowerShell environment, 
# and possibly add aliases you do not know about, your scripts behave completely different.

# Here is a simple call that adds the alias Get-ChildItem and lets it point to ping.exe:

Set-Alias -Name Get-ChildItem -Value ping

# This will change everything: Not only will Get-ChildItem now ping instead of list folder content. 
# Also, all aliases (like “dir” and “ls”) now ping.
# Just imagine the alias would point to format.exe instead, and think what your scripts would now do.

###########################################################################################################################################################################################

# Tip 24: Converting Special Characters


# Sometimes it becomes necessary to replace special characters with other characters. Here is a simple function that does the trick:

function ConvertTo-PrettyText($Text)
{
    $Text.Replace('ü','ue').Replace('ö','oe').Replace('ä', 'ae' ).Replace('Ü','Ue').Replace('Ö','Oe').Replace('Ä', 'Ae').Replace('ß', 'ss')
}

ConvertTo-PrettyText -Text "Mr. Össterßlim"         # Output: Mr. Oesstersslim

# Simply add as many Replace() calls as you need to process the text. 
# Note that Replace() is case-sensitive which is great: you can do case-correct replacements that way.

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Here is another approach that may be a bit slower but is easier to maintain. It also features a case-sensitive hash table:

function ConvertTo-PrettyText($Text)
{
    $hash = New-Object -TypeName HashTable

    $hash.'ä' = 'ae'
    $hash.'ö' = 'oe'
    $hash.'ü' = 'ue'
    $hash.'ß' = 'ss'
    $hash.'Ä' = 'Ae'
    $hash.'Ö' = 'Oe'
    $Hash.'Ü' = 'Ue'

    foreach($key in $hash.Keys)
    {
        $Text = $Text.Replace($key, $hash.$key)
    }

    $Text
}

ConvertTo-PrettyText -Text "Mr. Össterßlim"            # Output: Mr. Oesstersslim

# Note that the function won’t define a hash table via “@{}” but rather instantiates a HashTable object. 
# While the hash table delivered by PowerShell is case-insensitive, the hash table created by the function is case-sensitive. 
# That’s important because the function wants to differentiate between lower- and uppercase letters.

# To add replacements, simply add the appropriate “before”-character to the hash table, and make its replacement text its value.



# If you’d rather specify ASCII codes, here is a variant that uses ASCII codes as key:

function ConvertTo-PrettyText($Text)   
{  
    $hash = @{

      228 = 'ae'
      246 = 'oe'
      252 = 'ue'
      223 = 'ss'
      196 = 'Ae'
      214 = 'Oe'
      220 = 'Ue'   
    }
    
    foreach($key in $hash.Keys)
    {
        $Text = $text.Replace([String][Char]$key, $hash.$key)
    }

    $Text
}

ConvertTo-PrettyText -Text "Mr. Össterßlim" 

###########################################################################################################################################################################################

# Tip 25: Managing Printers Low-Level


# Recent Windows operating systems like Windows 8 and Server 2012 come with great printing support, 
# but if you run older Windows versions, then this call may help:

rundll32.exe PRINTUI.DLL,PrintUIEntry

# Note that this call is case-sensitive! Do not add spaces, and do not change casing.

# The call opens a help window with a great number of sample calls that illustrate how you can install, 
# remove and copy printer drivers, among other things. 
# This tool will also work remotely, provided you allowed remote access via appropriate group policies.

###########################################################################################################################################################################################

# Tip 26: Recursing a Given Depth


# When you use Get-ChildItem to list folder content, you can add the –Recurse parameter to dive into all subfolders. 
# However, this won’t let you control the nesting depth. Get-ChildItem will now search in all subfolders, no matter how deeply they are nested.

Get-ChildItem -Path $env:windir -Filter *.log -Recurse -ErrorAction SilentlyContinue


# Sometimes, you see solutions like this one, in an effort to limit the nesting depth:

Get-ChildItem -Path $env:windir\*\*\* -Filter *.log -ErrorAction SilentlyContinue


# However, this would not limit nesting depth to three levels.
# Instead, it would search in all folders three level deeper. It would not, for example, search any folder in level 1 and 2.

function Get-MyChildItem
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Path,

        $Filter = "*",

        [System.Int32]
        $MaxDepth = 3,

        [System.Int32]
        $Depth = 0
    )

    $Depth++

    Get-ChildItem -Path $Path -Filter $Filter -File

    if($Depth -le $MaxDepth)
    {
        Get-ChildItem -Path -Directory | ForEach-Object { 
        
            Get-MyChildItem -Path $_.FullName -Filter $Filter -Depth $Depth -MaxDepth $MaxDepth 
        }
    }
}

Get-MyChildItem -Path C:\Windows -Filter *.log -MaxDepth 2 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName

###########################################################################################################################################################################################

# Tip 27: Hibernate System


# Here is a simple system call that will hibernate a system (provided of course that hibernation is enabled):

function Start-Hibernation
{
    rundll32.exe PowrProf.dll, SetSupendState 0,1,0   # Note that this call is case-sensitive!
}


# or from comments:

shutdown -h

###########################################################################################################################################################################################

# Tip 28: Case-Correct Name Lists


# Let’s assume it’s your job to update a list of names. 
# Here is an approach that will make sure that only first letter in a name is capitalized. 
# This approach works with double-names as well:

$names = 'some-wILD-casING','frank-PETER','fred'

foreach($name in $names)
{
    $corrected = foreach ($part in $name.Split("-"))
    {
        $firstChar = $part.SubString(0,1).ToUpper()
        $remaining = $part.SubString(1).toLower()

        "$firstChar$remaining"
    }

    $corrected -join "-"
}

# Output:
#        Some-Wild-Casing
#        Frank-Peter
#        Fred

###########################################################################################################################################################################################

# Tip 29: A Fun Beeping Prompt


# If your computer has a sound card, here is a code snippet that will drive your colleagues nuts:

function prompt
{
    1..3 | ForEach-Object {
    
        $frequency = Get-Random -Minimum 400 -Maximum 10000
        $duration = Get-Random -Minimum 400 -Maximum 400

        [Console]::Beep($frequency, $duration)
    }

    "PS> "

    $host.UI.RawUI.WindowTitle = Get-Location
}

# This function will shorten the PowerShell prompt and instead display the current location in the title bar. That’s the productivity part. 
# The destructive part is an irritating triple-beep with varying frequencies each time PowerShell completes a command.

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Have PowerShell Cheer You Up!


# Writing PowerShell code is fun but can be frustrating at times. 
# Here’s a function that makes PowerShell cheer you up. Just turn on your sound, and PowerShell will comment each command with a new remark.

function prompt
{
    $text = 'You are great!', 'Hero!', 'What a checker you are.', 
            'Champ, well done!', 'Man, you are good!', 'Guru stuff I would say.', 'You are magic!'
    
    'PS> '

    $host.UI.RawUI.WindowTitle = Get-Location
    
    (New-Object -ComObject SAPI.SPVoice).Speak(($text | Get-Random))
}

###########################################################################################################################################################################################

# Tip 30: Correcting PowerShell Paths


# Occasionally, you might stumble across strange path formats like this one:

Microsoft.PowerShell.Core\FileSystem::C:\Windows\explorer.exe

# This is a full PowerShell path name which includes the module name and provider name that is attached to this path. 
# To get the pure path name, use this:

Convert-Path -Path Microsoft.PowerShell.Core\FileSystem::C:\Windows\explorer.exe   # Output: C:\Windows\explorer.exe

###########################################################################################################################################################################################

# Tip 31: Finding and Dumping Registry Key Paths


# This code recursively searches through HKEY_CURRENT_USER and dumps all Registry keys that contain the word “powershell” 
# (simply replace the search word with anything else that you may be looking for):

Get-ChildItem -Path HKCU:\ -Include *Powershell* -Recurse -ErrorAction SilentlyContinue | Select-Object -Property *Path* | Out-GridView



# The code outputs all properties that have “Path” in it, and as you will see, 
# Registry keys have two properties that contain the key location: PSPath and PSParentPath. Both use the internal PowerShell path format.

# To simply dump the Registry paths for all keys meeting your search criteria, try this:

Get-ChildItem -Path HKCU:\ -Include *Powershell* -Recurse -ErrorAction SilentlyContinue | ForEach-Object {

    Convert-Path -Path $_.PSPath
}

###########################################################################################################################################################################################

# Tip 32: Edit Network “hosts” File


# If you find yourself editing the “hosts” file regularly, 
# then it may be tedious to manually open the file in an elevated instance of the Notepad. 
# Since this file can only be edited by Administrators, a normal Notepad instance won’t do.

# Here is some code that you can use as-is or easily adjust to open any program with elevated privileges.

function Show-HostsFile
{
    $Path = "$env:windir\system32\drivers\etc\hosts"

    Start-Process -FilePath notepad -ArgumentList $Path -Verb runas
}

Show-HostsFile

###########################################################################################################################################################################################

# Tip 33: Logging What a Script Does


# You probably know that in a PowerShell console (not the ISE editor), you can turn on logging:

Start-Transaction

# This will log all entered commands and all command results to a file. 
# Unfortunately it is of limited use when you run a script, because you cannot see the actual script commands.


# Here is a radical trick that also includes all commands that your script executed.
# Before you try this trick, be aware that this will increase the size of your log file and can slow down script execution, 
# because in loops, each iteration of a loop would also be logged.

# This is all you need to enable logging script commands:

Set-PSDebug -Trace 1

###########################################################################################################################################################################################

# Tip 34: Use Group-Object to Create Hash Tables


# Group-Object can pile objects up, putting objects with the same property together in one pile.

# This can be quite useful, especially when you ask Group-Object to return hash tables. 
# This would generate a hash table with piles for all available service status modes:

$hash = Get-Service | Group-Object -Property Status -AsHashTable -AsString
$hash
# Output:
#        Name                           Value                                                                                                                                                                                 
#        ----                           -----                                                                                                                                                                                 
#        Stopped                        {AdobeFlashPlayerUpdateSvc, AeLookupSvc, ALG, AppMgmt...}                                                                                                                             
#        Running                        {AdobeARMservice, AMD External Events Utility, AppHostSvc, AppIDSvc...}

# You could now get back the list of all running (or stopped) services like this:

$hash.Running
$hash.Stopped




# Use any object property you want to create the piles. 
# This example will pile up files in three piles: one for small, one for medium, and one for large files.

$code = {

    if($_.Length -gt 1MB){ "Huge" }

    elseif($_.Length -gt 1MB){ "Average" }

    else{ "Tiny" }
}

$hash = Get-ChildItem -Path C:\Windows | Group-Object -Property $code -AsHashTable -AsString
$hash
# Output:
#        Name                           Value                                                                                                                                                                                 
#        ----                           -----                                                                                                                                                                                 
#        Huge                           {explorer.exe, MEMORY.DMP, WindowsUpdate.log}                                                                                                                                         
#        Tiny                           {%LOCALAPPDATA%, addins, AppCompat, AppPatch...} 

$hash.Tiny
$hash.Huge

###########################################################################################################################################################################################

# Tip 35: Using the OpenFile Dialog


# Here’s a quick function that works both in the ISE editor and the PowerShell console in PowerShell 3.0 and above

function Show-OpenFileDialog
{
    param
    (
        $StartFolder = [Environment]::GetFolderPath("MyDocument"),
        $Title = "Open what?",
        $Filter = 'All|*.*|Scripts|*.ps1|Texts|*.txt|Logs|*.log'
    )

    Add-Type -AssemblyName PresentationFramework

    $dialog = New-Object -TypeName Microsoft.Win32.OpenFileDialog

    $dialog.Title = $Title
    $dialog.InitialDirectory = $StartFolder
    $dialog.Filter = $Filter

    $result = $dialog.ShowDialog()

    if($result -eq $true)
    {
        $dialog.FileName
    }
}

Show-OpenFileDialog

# It opens a OpenFile dialog. The user can select a file, and the selected file is returned to PowerShell. 
# So next time your script needs to open a CSV file, you may want to use the additional luxury of opening a selection dialog.

###########################################################################################################################################################################################

# Tip 36: Correcting ISE Encoding



# When you run a console application inside the ISE editor, 
# non-standard characters such as “ä” or “ß” do not show correctly in results. 
# To correct the encoding ISE uses to communicate with its hidden console, run this:

cmd.exe /c echo ÄÖÜäöüß      # Output: Ž™š„”á



# Repair encoding. This REQUIRES a console app to run first because only
# then will ISE actually create its hidden background console

$null = cmd.exe /c echo
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Now all is fine
cmd.exe /c echo ÄÖÜäöüß      # Output: ÄÖÜäöüß

###########################################################################################################################################################################################

# Tip 37: Filtering results with regular expressions


# Get-ChildItem does not support advanced file filtering. While you can use simple wildcards, you cannot use regular expressions.

# To work around this, add a cmdlet filter and use the operator -match.

# This example will find all files within your Windows directory structure 
# that have file names with at least 2-digit numbers in them, and a maximum file name length of 8 characters:

Get-ChildItem -Path $env:windir -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.BaseName -match "\d{2}" -and $_.Name.Length -le 8 }

# Note the use of the property "BaseName". It returns the filename without an extension. This way, numbers in file extensions won't count.

###########################################################################################################################################################################################

# Tip 38: Getting Shutdown Information


# Windows logs all shutdown events in its System event log. From there, you can extract and analyze the information.

# Here is a function that looks for the appropriate event log entries, 
# reads the relevant information from the ReplacementStrings array, and returns the shutdown information as objects.

function Get-ShutDownInfo
{
    Get-EventLog -LogName System -InstanceId 2147484722 -Source User32 | ForEach-Object {
    
        $result = "dummy" | Select-Object -Property Computer, TimeWritten, User, Reason, Action, Executable
        
        $result.TimeWritten = $_.TimeWritten
        $result.User = $_.ReplacementStrings[6]
        $result.Reason = $_.ReplacementStrings[2]
        $result.Action = $_.ReplacementStrings[4]
        $result.Executable = Split-Path -Path $_.ReplacementStrings[0] -Leaf

        $result.Computer = $_.MachineName

        $result
    }
}

# Now it is easy to check for shutdown problems:

Get-ShutDownInfo | Out-GridView

###########################################################################################################################################################################################

# Tip 39: Filtering Hotfix Information


# Get-HotFix is a built-in cmdlet that returns the installed hotfixes. It has no parameter to filter on hotfix ID, though.

# With a cmdlet filter, you can easily focus on the hotfixes you are after. 
# This example returns only hotfixes with an ID that starts with "KB25":

Get-HotFix | Where-Object { $_.HotFix -like "KB25" }

# Note that Get-HotFix has a parameter -ComputerName, so assuming you have the appropriate permissions, 
# you can retrieve hotfix information from a remote computer as well.

Get-HotFix -ComputerName Server02 | Where-Object { $_.HotFix -like "KB25" }

###########################################################################################################################################################################################

# Tip 40: Get Sleep and Hibernation Times


# If you want to find out whether a computer is frequently put into sleep or hibernation mode, 
# here is a function that reads the appropriate event log entries and returns a table with details, 
# reporting when the computer was put into sleep mode, and how long the computer remained in sleep mode:

Get-EventLog -LogName System -InstanceId 1 -Source Microsoft-Windows-Power-TroubleShooter | select -First 1 |ForEach-Object {$_.ReplacementStrings}

# Output:
#          2014-11-27T19:04:31.828768800Z
#          2014-11-28T00:56:34.603063100Z
#          9783
#          1101
#          700
#          1488
#          8042
#          0
#          54665
#          21505
#          4
#          5
#          4
#          30
#          PCI standard PCI-to-PCI bridge
#          0
#          0



function Get-HibernationTime
{
    Get-EventLog -LogName System -InstanceId 1 -Source Microsoft-Windows-Power-TroubleShooter | ForEach-Object {

        $result = "dummy" | Select-Object -Property ComputerName, SleepTime, WakeTime, Duration
        
        [DateTime]$result.SleepTime = $_.ReplacementStrings[0]
        [DateTime]$result.WakeTime = $_.ReplacementStrings[1]

        $time = $result.WakeTime - $result.SleepTime

        $result.Duration = ([int]($time.TotalHours * 100))/100
        $result.ComputerName = $_.MachineName
        
        $result                # return result
    }
}

Get-HibernationTime

# Output:
#        ComputerName      SleepTime                 WakeTime                 Duration
#        ------------      ---------                 --------                 --------
#        SIHE              2014/11/28 03:04:31       2014/11/28 08:56:34          5.87
#        SIHE              2014/11/27 19:23:16       2014/11/28 03:01:23          7.64
#        SIHE              2014/11/26 12:30:54       2014/11/26 12:32:43          0.03
#        SIHE              2014/11/26 04:37:25       2014/11/26 08:54:21          4.28

###########################################################################################################################################################################################

# Tip 41: WMI Device Inventory 


# The WMI service can report plenty of details about the computer hardware. 
# Typically, each type of hardware is represented by its own WMI class. 
# It's not easy to find out the names of such hardware classes, though.

# Since all hardware classes descend from the same root WMI class (CIM_LogicalDevice), 
# you can use this root class to find all hardware:

Get-WmiObject -Class CIM_LogicalDevice |Out-GridView

# This will return a basic hardware inventory. But you can do even more. 
# With a little extra code, you get a list of hardware class names used by WMI:

Get-WmiObject -Class CIM_LogicalDevice | Select-Object -Property __Class, Description | Sort-Object -Property __Class -Unique | Out-GridView

# Output:
#        __CLASS                         Description                                                                                               
#        -------                         -----------                                                                                               
#        Win32_Bus                       Bus                                                                                                       
#        Win32_CacheMemory               Cache Memory                                                                                              
#        Win32_CDROMDrive                CD-ROM Drive                                                                                              
#        Win32_DesktopMonitor            Generic PnP Monitor                                                                                       
#        Win32_DiskDrive                 Disk drive                                                                                                
#        Win32_DiskPartition             Installable File System                                                                                   
#        Win32_Fan                       Cooling Device                                                                                            
#        Win32_IDEController             IDE Channel                                                                                               
#        Win32_Keyboard                  Standard PS/2 Keyboard                                                                                    
#        Win32_LogicalDisk               Local Fixed Disk                                                                                          
#        Win32_MemoryArray               Memory Array                                                                                              
#        Win32_MemoryDevice              Memory Device                                                                                             
#        Win32_MotherboardDevice         Motherboard                                                                                               
#        Win32_NetworkAdapter            WAN Miniport (PPTP)                                                                                       
#        Win32_PnPEntity                 Ancillary Function Driver for Winsock                                                                     
#        Win32_PointingDevice            USB Input Device                                                                                          
#        Win32_Printer                                                                                                                             
#        Win32_Processor                 AMD64 Family 16 Model 4 Stepping 2                                                                        
#        Win32_SerialPort                Communications Port                                                                                       
#        Win32_SoundDevice               High Definition Audio Device                                                                              
#        Win32_TemperatureProbe          CPU Temperature                                                                                           
#        Win32_USBController             Standard OpenHCD USB Host Controller                                                                      
#        Win32_USBHub                    USB Root Hub                                                                                              
#        Win32_VideoController           ATI Radeon HD 4200                                                                                        
#        Win32_Volume  

# You can now use any of these class names to query for a particular type of hardware and find out hardware details:

Get-WmiObject -Class Win32_SoundDevice       

# Output:
#        Manufacturer      Name                               Status       StatusInfo
#        ------------      ----                               ------       ----------
#        Microsoft         High Definition Audio Device       OK                    3                                                                                                                     

###########################################################################################################################################################################################

# Tip 42: Test Service Responsiveness


# To test whether a particular service is still responding, use a clever trick. 
# First, ask WMI for the service you want to check. WMI will happily return the process ID of the underlying process.

# Next, look up this process, and the process object will tell you whether the process is frozen or responding:

function Test-ServiceResponding($ServiceName)
{
    $service = Get-WmiObject -Class Win32_Service -Filter "Name='$ServiceName'"
    $processID = $service.processID

    $process = Get-Process -Id $processID

    $process.Responding
}

# This example would check whether the Windows Update service is still responding:

Test-ServiceResponding -ServiceName wuauserv

#　Note that the example code assumes that the service is running. 
# If you wanted to, you could add a check to exclude non-running services yourself.

###########################################################################################################################################################################################

# Tip 43: Finding Attached USB Sticks


# If you'd like to know whether there are currently USB storage devices attached to your computer, WMI can help:

Get-WmiObject -Class Win32_PnpEntity | Where-Object { $_.DeviceID -like "USBSTOR*" }

# This returns all plug and play devices with a device class of "USBSTOR".



# If you are willing to use the WMI query language (WQL), you could even do this in a cmdlet filter:

Get-WmiObject -Query 'Select * from Win32_PnPEntity where DeviceID like "USBSTOR%"'

###########################################################################################################################################################################################

# Tip 44: System Uptime


$uptime = [Environment]::TickCount
$uptime                             # return how long it goes after your machine is turn on, it changed all the time

"I am up for $uptime milliseconds!"


$uptime = [Environment]::TickCount
$timespan = New-TimeSpan -Seconds ($uptime / 1000)

$timespan
# Output:
#        Days              : 0
#        Hours             : 2
#        Minutes           : 21
#        Seconds           : 22
#        Milliseconds      : 0
#        Ticks             : 84820000000
#        TotalDays         : 0.0981712962962963
#        TotalHours        : 2.35611111111111
#        TotalMinutes      : 141.366666666667
#        TotalSeconds      : 8482
#        TotalMilliseconds : 8482000

$hours = $timespan.TotalHours
"System is up for {0:n0} hours now." -f $hours

# As a special treat, New-Timespan cannot take milliseconds directly, so the script had to divide the milliseconds by 1000, introducing a small inaccuracy.

# To turn milliseconds in a timespan object without truncating anything, try this:

$timespan = [TimeSpan]::FromMilliseconds($uptime)

# It won't make a difference in this example, but can be useful elsewhere. 
# For example, you also have a FromTicks() method available that can turn ticks 
# (the smallest unit of time intervals on Windows systems) into intervals.

###########################################################################################################################################################################################

# Tip 45: Playing WAV Sounds


# To play a WAV sound file in a background process, PowerShell can use the built-in SoundPlayer class. 
# It accepts a path to a WAV file and lets you then decide whether you want to play the sound once or repeatedly.


$player = New-Object -TypeName System.Media.SoundPlayer
$player.SoundLocation = "C:\Windows\Media\chimes.wav"
$player.Load()
$player.PlayLooping()   # This would play a sound repeatedly

# do something here...

$player.Stop()          # Once your script is done, it can stop playback using this call


# playback the file mySound.wav that is located in the same folder as your script
$player.SoundLocation = "$PSScriptRoot\mySound.wav"

# Note that $PSScriptRoot requires PowerShell version 3.0 or later. It also requires your script to be saved to a file, of course.

###########################################################################################################################################################################################

# Tip 46: Using -f Operator to Combine String and Data


# Strings enclosed by double-quotes will resolve variables so it is common practice to use code like this:

$name = $Host.Name
"Your host is called $name."           # Output: Your host is called Windows PowerShell ISE Host.


# However, this technique has limitations. If you wanted to display object properties and not just variables, it fails:

"Your host is called $host.Name." 
# Output: Your host is called System.Management.Automation.Internal.Host.InternalHost.Name.

# Note: "Your host is called $($host.Name)." can get the right one also. 

# This is because PowerShell only resolves the variable (in this case $host), not the rest of your code.




# And you cannot control number formats, either. This line works, but the result simply has too many digits to look good:

$freeSpace = ([WMI]"Win32_LogicalDisk.DeviceID='C:'").FreeSpace
$freeSpaceMB = $freeSpace / 1MB

"Your C: drive has $freeSpaceMB MB space available." 
# Output: Your C: drive has 11956.76171875 MB space available.



# The –f operator can solve both problems. It takes a static text template to the left, and the values to insert to the right:

'Your host is called {0}.' -f $host.Name                          # Output: Your host is called Windows PowerShell ISE Host.

'Your C: drive has {0:n1} MB space available.' -f $freeSpaceMB    # Output: Your C: drive has 11,956.8 MB space available.

# As you see, using -f gives you two advantages: the placeholders (braces) tell PowerShell where insertion should start and end, 
# and the placeholders accept formatting hints, too. "n1" stands for numeric data with 1 digit. Simply change the number to suit your needs.

###########################################################################################################################################################################################

# Tip 47: Countdown Hours


# Whether it is a birthday or an important game, you may want PowerShell to tell you just how many hours it is till the event starts. Here is how:

$result = New-TimeSpan -End "2014/12/25 00:00:00"
$hours = [Int]$result.TotalHours

'Another {0:n0} hours to go...' -f $hours      # Output: Another 538 hours to go...

# This example calculates the hours to go until Christmas 2014. 
# Simply replace the date in this piece of code to find out just how many more hours you need to wait for your favorite event to start.

###########################################################################################################################################################################################

# Tip 48: Combining Results


# Let's assume you want to identify suspicious service states like services 
# that are stopped although their start mode is set to "Automatic", or identify services with ExitCodes that you know are bad.

# Here is some example code that illustrates how you can query these scenarios and combine the results in one variable.

# Sort-Object makes sure you do not have duplicates in your list before all the collected results are output to just one grid view window.

$list = @()

$list += Get-WmiObject -Class Win32_Service -Filter 'State="Stopped" and StartMode="Auto" and ExitCode!=0' | 
            Select-Object -Property Name, DisplayName, ExitCode, Description, PathName, DesktopInteract 

$list += Get-WmiObject -Class Win32_Service -Filter 'ExitCode!=0 and ExitCode!=1077' | 
            Select-Object -Property Name, DisplayName, ExitCode, Description, PathName, DesktopInteract 

$list | Sort-Object -Unique -Property Name | Out-GridView



# From comments feedback:

$list = @(

    Get-WmiObject -Class Win32_Service -Filter 'State="Stopped" and StartMode="Auto" and ExitCode!=0'
    Get-WmiObject -Class Win32_Service -Filter 'ExitCode!=0 and ExitCode!=1077'

) | Select-Object -Property Name, DisplayName, ExitCode, Description, PathName, DesktopInteract

$list | Sort-Object -Unique -Property Name | Out-GridView

###########################################################################################################################################################################################

# Tip 49: Finding Minimum and Maximum Values


# To find the smallest and largest item in a range of numbers, use Measure-Object:

$list = 1,5,9,52,564,1,5,68

$result = $list | Measure-Object -Minimum -Maximum
$result.Minimum                                     # Output: 1
$result.Maximum                                     # Output: 564

# This works for any input data and any data type. 



# Here is a slight modification that returns the oldest and newest file in your Windows folder:

$list = Get-ChildItem -Path C:\Windows

$result = $list | Measure-Object -Property LastWriteTime -Minimum -Maximum

$result.Minimum                  # Output: Thursday, February 09, 2006 17:50:00                              
$result.Maximum                  # Output: Tuesday, December 02, 2014 13:52:26

# Simply add the –Property parameter if your input data has multiple properties, and pick the one you want to examine.

###########################################################################################################################################################################################

# Tip 50: Useful Path Manipulation Shortcuts


# Here are a bunch of useful (and simple to use) system functions for dealing with file paths:

[System.IO.Path]::GetFileNameWithoutExtension("file.ps1")         # Output: file
[System.IO.Path]::GetExtension("file.ps1")                        # Output: .ps1
[System.IO.Path]::ChangeExtension("file.ps1", ".copy.ps1")        # Output: file.copy.ps1


# Get methods from [System.IO.Path]
[System.IO.Path] | Get-Member -MemberType Method

# All of these methods accept either file names or full paths, 
# and return different aspects of the path, or change things like the extension.

###########################################################################################################################################################################################

# Tip 51: Important Math Functions


[Math]::Floor(4.9)       # Output: 4
[Math]::Ceiling(3.2)     # Output: 4
[Math]::Max(3, 8)        # Output: 8
[Math]::Min(3, 8)        # Output: 3

# To explore further, Get-Member is your friend.


# Related Reference: http://msdn.microsoft.com/zh-cn/library/system.math_methods(v=vs.80).aspx

###########################################################################################################################################################################################

# Tip 52: Optional and Mandatory at the Same Time


# A parameter can be mandatory when other parameters are present, and optional otherwise.

function Connect-SomeWhere
{
    [CmdletBinding(DefaultParameterSetName = "A")]
    param
    (
        [Parameter(ParameterSetName = "A", Mandatory = $false)]
        [Parameter(ParameterSetName = "B", Mandatory = $true)]
        $ComputerName,

        [Parameter(ParameterSetName = "B", Mandatory = $false)]
        $Credential
    )

    $chosen = $PSCmdlet.ParameterSetName

    "You have chosen $chosen parameter set."
}

Connect-SomeWhere                     # -Computername is optional here
Connect-SomeWhere -Credential test    # -Computername is mandatory here

###########################################################################################################################################################################################

# Tip 53: Discarding Results


# Since PowerShell returns anything that commands leave behind, 
#　it is particularly important in PowerShell scripts to discard any result that you do not want to return.

# There are many way to achieve this, and here are the two most popular. 
# Note that both lines try and create a new folder on your drive C:. New-Item would return the folder object, 
# but if all you want is create a new folder, then you may want to discard the result:

$null = New-Item -Path C:\newfolder -ItemType Directory
[void](New-Item -Path C:\newfolder -ItemType Directory)

# Note: [void] is equivalent to assigning to $null, 
# but assigning to $null is more "powershell-like" whereas [void] is more .Netish. 



New-Item -Path C:\newfolder -ItemType Directory | Out-Null

# So which approach is better? Definitely the first one. 
# Piping unwanted results to Out-Null is expensive and takes about 40x the time. 
# You won't notice on single calls, but if this happens within a loop, it may become significant.


# So better get into the habit of using $null rather than Out-Null!

###########################################################################################################################################################################################

# Tip 54: Removing Illegal Path Characters



# In path names, some characters like colons or quotes are illegal. 
# If your script derives path names from some pieces of information, 
# you may want to make sure that the resulting path name is legal.

# Here is a function that takes any path and replaces illegal path characters with an underscore:

function Get-LegalPathName($Path)
{
    $illegalChars = [System.IO.Path]::GetInvalidFileNameChars()

    foreach($illegalChar in $illegalChars)
    {
        $Path = $Path.Replace($illegalChar, "_")
    }

    $Path
}

Get-LegalPathName 'some:"illegal"\path<chars>.txt'    # Output: some__illegal__path_chars_.txt

###########################################################################################################################################################################################

# Tip 55: Getting the Number of Lines in a String


# Here is a clever trick how to find out how many lines a string (not a string array!) contains:

$text = @"
This is some
sample text
Let's find out
the number of lines.
"@

$lines = $text.Length - $text.Replace("`n", "").Length + 1
$lines   # Output: 4 

# Technically, the example uses a here-string to create the multi-line string, 
# but this is just an example. It works for all kinds of strings, regardless of origin.

###########################################################################################################################################################################################

# Tip 56: Testing Whether Text Contains Upper Case



$text1 = 'this is all lower-case'
$text2 = 'this is NOT all lower-case'


# Use regular expressions to check whether a string contains at least one upper case letter:

$text1 -cmatch '[A-Z]'              # Output: False
$text2 -cmatch '[A-Z]'              # Output: True



# To check whether a text contains purely lower-case letters, try this:

$text1 -cmatch '^[a-z\s-]*$'       # Output: True
$text2 -cmatch '^[A-Z\s-]*$'       # Output: False


# Basically, this test is harder because you need to include all characters that you consider legal. 
# In this example, I chose the lower-case letters from a to z, whitespace, and the minus sign.

# These "legal" characters are embedded within "^" and "$" (line start and line end). 
# The star is a quantifier (any number of "legal" characters).

###########################################################################################################################################################################################

# Tip 57: Finding Errors in Scripts


# It’s never been easier to find scripts with syntax errors in them. Just use this filter:

filter Test-SyntaxError
{
    $text = Get-Content -Path $_.FullName

    if($text.Length -gt 0)
    {
        $err = $null

        $null = [System.Management.Automation.PSParser]::Tokenize($text, [ref]$err)

        if($err){ $_ }
    }
}

# With it, you can quickly scan folders or even entire computers and list all PowerShell files that have syntax errors in them.


# This would search and find all PowerShell scripts in your user profile and list only those that have syntax errors:

dir $HOME -Filter *.ps1 -Recurse -Exclude *.ps1xml | Test-SyntaxError

###########################################################################################################################################################################################

# Tip 58: Waiting for a Keystroke


# To keep the PowerShell console open when a script is done, 
# you may want to add a “Press Any Key” statement. Here is a way how you can implement this:

Write-Host "Press any key!" -NoNewline
$null = [Console]::ReadKey("?")

# Or

$keypress = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")


# This will work in a real PowerShell console only. 
# It will not work in the ISE editor or any other PowerShell editor 
# that does not use a real console with a real interactive keyboard buffer.

###########################################################################################################################################################################################

# Tip 59: Download PowerShell Language Specification

# the link refer to Powershell 2.0
$link = 'http://download.microsoft.com/download/3/2/6/326DF7A1-EE5B-491B-9130-F9AA9C23C29A/PowerShell%202%200%20Language%20Specification.docx'

$outfile = "$env:temp\languageref.docx"

Invoke-WebRequest -Uri $link -OutFile $outfile


# From comments, here a solution for "Cache Access Denied."  -->> proxy issue

$Global:PSDefaultParameterValues = @{

    'Invoke-RestMethod:Proxy'='http://proxyServer:proxyPort'
    
    'Invoke-WebRequest:Proxy'='http://proxyServer:proxyPort'
    
    '*:ProxyUseDefaultCredentials'=$true
}

# Windows PowerShell Language Specification Version 3.0:
# site: http://www.microsoft.com/en-us/download/details.aspx?id=36389

###########################################################################################################################################################################################

# Tip 60: Comparing Service Configuration 


# Provided you have PowerShell remoting up and running on servers, 
# here is a simple script that illustrates how you can get the state of all services from each server
# and then calculate the differences between the two servers.

$server = "IIS-CTI5052"

$serviceForServer = Invoke-Command { Get-Service } -ComputerName $server | Sort-Object -Property Name, Status
$serviceForLocal = Get-Service | Sort-Object -Property Name, Status

Compare-Object -ReferenceObject $serviceForLocal -DifferenceObject $serviceForServer -Property Name, Status | Sort-Object -Property Name

# Output:
#        Name                                     Status SideIndicator                                                         
#        ----                                     ------ -------------                                                         
#        AdobeARMservice                         Running <=                                                                    
#        AdobeFlashPlayerUpdateSvc               Stopped <=                                                                    
#        AMD External Events Utility             Running <=                                                                    
#        AnonymizedLogUploader                   Stopped =>                                                                    
#        AppIDSvc                                Stopped =>    

# The result is a list with only the differences in service configuration.

###########################################################################################################################################################################################

# Tip 61: Dumping Service State Information 


# If you would like to save the results of a PowerShell command to disk
# so that you can take it with you to another machine, here is a simple way:

$path = "$env:temp\mylist.xml"

# This will get all services using Get-Service. 
# The results are tagged with a new column called "ComputerName" that will show the computer name on which the data was taken.

Get-Service | Add-Member -MemberType NoteProperty -Name ComputerName -Value $env:COMPUTERNAME -PassThru | Export-Clixml -Depth 1 -Path $path

explorer.exe "/select,$path"

# Then, the results are saved to disk as serialized XML. 
# Explorer will open the destination folder and select the created XML file 
# so you can easily copy it to a USB stick and take it wherever you want.


Import-Clixml -Path $path      # To "rehydrate" the results elsewhere back into real objects

###########################################################################################################################################################################################

# Tip 62: Finding PowerShell Functions 


# To quickly scan your PowerShell script repository and find all files that have a given function in them, try this filter:

filter Find-Function
{
    $path = $_.FullName
    $lastWriteTime = $_.LastWriteTime
    $text = Get-Content -Path $path

    if($text.Length -gt 0)
    {
        $token = $null
        $errors = $null

        $ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref] $token, [ref] $errors)

        $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | 

            Select-Object -Property Name, Path, LastWriteTime | ForEach-Object {
            
                $_.Path = $path
                $_.LastWriteTime = $lastWriteTime

                $_
            }
    }
}

dir $HOME -Filter *.ps1 -Recurse -Exclude *.ps1xml | Find-Function

###########################################################################################################################################################################################

# Tip 63: Creating TinyURLs


$originalURL = "http://powershell.com/cs/blogs/tips/default.aspx?PageIndex=4"
$url = "http://tinyurl.com/api-create.php?url=$originalURL"


# Method One: 
$webClient = New-Object -TypeName System.Net.WebClient
$webClient.DownloadString($url)                     # Output: http://tinyurl.com/n8f5xkv


# Method Two: 
(Invoke-WebRequest -Uri $url).Content               # Output: http://tinyurl.com/n8f5xkv



# From comments: if you need use proxy

$webClient = New-Object -TypeName System.Net.WebClient

$webClient.Headers.Add("User-Agent", "Mozilla/4.0+")
$webClient.Proxy = [System.Net.WebRequest]::DefaultWebProxy
$webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

$webClient.DownloadString($url)

###########################################################################################################################################################################################

# Tip 64: Replacing Duplicate Spaces


$str = '[  Man, it    works!   ]'

$str -replace "\s{2,}", " "        # Output: [ Man, it works! ]




# You can use this approach also to convert fixed-width text tables to CSV data:

qprocess

# Output:
#        >sihe                console              1  16264  notepad.exe
#        >sihe                console              1  15032  msspellcheck...
#        >sihe                console              1   7376  conhost.exe
#        >sihe                console              1   9248  qprocess.exe

(qprocess) -replace "\s{2,}", ","

# Output:
#        >sihe,console,1,16264,notepad.exe
#        >sihe,console,1,15032,msspellcheck...
#        >sihe,console,1,7376,conhost.exe
#        >sihe,console,1,14068,qprocess.exe




# Once it is CSV, you can use ConvertFrom-Csv to turn text data into objects:

(qprocess) -replace "\s{2,}", "," | ConvertFrom-Csv -Header Name, Session, ID, Pid, Process

# Output:
#        Name    : >v-sihe
#        Session : console
#        ID      : 1
#        Pid     : 16264
#        Process : notepad.exe

###########################################################################################################################################################################################

# Tip 65: Text Splitting


# With the –split operator, you can split text at given locations. The operator expects a regular expression, 
# so if you just want to split using plain text expressions, you need to escape your split text.

# Here is an example that splits a path at the backslash:

$originalText = 'c:\windows\test\file.txt'
$splitText = [regex]::Escape("\")

$parts = $originalText -split $splitText
$parts
# Output:
#        c:
#        windows
#        test
#        file.txt

$parts[0]
$parts[-1]

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Advanced Text Splitting

$str = 'Hello, this is a text, and it has commas' 

# When you use the –split operator to split text, then the split text is consumed:

$str -split ","
# Output:
#        Hello
#         this is a text
#         and it has commas



# The split text can be longer than just one character. This would split at any location that has a comma plus a space:

$str -split ", "
# Output:
#        Hello
#        this is a text
#        and it has commas



# And since –split expects really a regular expression, this would split at any location that is a comma and then at least one space:

'Hello,    this is a    text, and it has commas' -split ',\s{1,}'
# Output:
#        Hello
#        this is a    text
#        and it has commas



# If you want, you can keep the split text with the results by enclosing the split text in “(?=…)”:

'Hello,    this is a    text, and it has commas' -split '(?=,\s{1,})'
# Output:
#        Hello
#        ,    this is a    text
#        , and it has commas
      
###########################################################################################################################################################################################

# Tip 65: Getting MAC Addresses


# Getting the MAC of a network adapter is rather simple in PowerShell. Here is one of many ways:

getmac.exe /FO CSV | ConvertFrom-Csv
# Output:
#        Physical Address          Transport Name                                                                                            
#        ----------------          --------------                                                                                            
#        28-A9-05-B8-1B-D0         \Device\Tcpip_{C2CAEA64-BBCB...} 



# The challenge might be that the actual column names are localized and can vary from culture to culture. 
# Since the raw information comes from CSV data emitted by getmac.exe, there is a simple trick though: 
# rename the columns to whatever you like by skipping the first line (containing the CSV headers), 
# and then submitting your own unique header names:

getmac.exe /FO CSV | Select-Object -Skip 1 | ConvertFrom-Csv -Header MAC, Transport
# Output:
#        MAC                       Transport                                                                                            
#        ----------------          --------------                                                                                            
#        28-A9-05-B8-1B-D0         \Device\Tcpip_{C2CAEA64-BBCB...}

# This will always produce columns named “MAC” and “Transport”.

# Of course there are object-oriented approaches, too, like asking the WMI or using special cmdlets in Windows 8.1 or Server 2012/2012 R2. 
# However, we believe the illustrated approach is a fun alternative and shows how to turn raw CSV data into really useful culture-invariant information.

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Identifying Network Vendors by MAC Address


# Each MAC address uniquely identifies a network device. The MAC address is assigned by a network equipment vendor. 
# So you can backtrack the vendor from any MAC address.

# All you need is the official IEEE vendor list which is more than 2MB in size. 
# Here is a script that downloads the list for you:

$url = 'http://standards.ieee.org/develop/regauth/oui/oui.txt'
$outfile = "$home\vendorlist.txt"

Invoke-WebRequest -Uri $url -OutFile $outfile



# Next, you can use the list to identify the vendor. Get some MAC addresses, for example like this:

getmac
# Output:
#        Physical Address          Transport Name                                                                                            
#        ----------------          --------------                                                                                            
#        28-A9-05-B8-1B-D0         \Device\Tcpip_{C2CAEA64-BBCB...}

# Take the first 3 octets of any MAC address, for example 5c-51-4f, and run these against the file you just downloaded:

Get-Content -Path $outfile | Select-String 18-A9-05 -Context 0,6
# Output:
#        >   18-A9-05   (hex)        Hewlett-Packard Company
#            18A905     (base 16)        Hewlett-Packard Company
#                            11445 Compaq Center Drive
#                          Houston Texas 77070
#                          UNITED STATES
#          
#            18-A9-58   (hex)        PROVISION THAI CO., LTD.

# Not only will you get the vendor (Intel in this case), but also its address and location. 

###########################################################################################################################################################################################

# Tip 66: Normalizing Line Endings


# When you download files from the Internet, you may run into situations where the file won’t open correctly in editors. 
# Most likely, this is caused by non-default line endings.


# When you do that and open the download file vendorlist.txt in the Notepad, all line breaks are gone

$url = 'http://standards.ieee.org/develop/regauth/oui/oui.txt'
$outfile = "$home\vendorlist.txt"

Invoke-WebRequest -Uri $url -OutFile $outfile
start $outfile


# To repair the file, simply use this code:
$oldFile = "$home\vendorlist.txt"
$newFile = "$home\vendorlistGood.txt"

Get-Content $outfile | Set-Content -Path $newFile
notepad $newFile

# Get-Content is capable of identifying even non-standard line breaks, so the result is a string array of lines. 
# When you write these back to a new file, then all is good because Set-Content will use default line endings.

###########################################################################################################################################################################################

# Tip 67: Renaming Variables


# Here is a simple variable renaming function that you can use in the built-in ISE editor that ships with PowerShell 3 and later.

# It will identify any instance of a variable and then replace it with a new name.

function Rename-Variable
{
    param
    (
        [Parameter(Mandatory = $true)]
        $OldName,

        [Parameter(Mandatory = $true)]
        $NewName
    )

    $inputText = $psISE.CurrentFile.Editor.Text
    $token = $null
    $errors = $null

    $ast = [System.Management.Automation.Language.Parser]::ParseInput($InputText, [ref] $token, [ref] $errors)

    $token | Where-Object { $_.Kind -eq "Variable" } | Where-Object { $_.Name -eq $OldName } | 

             Sort-Object { $_.Extent.StartOffset } -Descending | ForEach-Object {
             
                $start = $_.Extent.StartOffset + 1
                $end = $_.Extent.EndOffset
                $inputText = $inputText.Remove($start, $end - $start).Insert($start, $NewName)
             }
    
    $psISE.CurrentFile.Editor.Text = $inputText
}

# Run the function, and you now have a new command called Rename-Variable.

# Next, open a script in the ISE editor, and in the console pane, enter this 
# (and of course, replace the old variable name “oldVariableName” with the name of a variable that actually exists in your currently opened ISE script).

Rename-Variable -OldName oldVariableName -NewName theNEWVariableName

# Immediately, all occurrences of the old variable are replaced with the new variable name.

# Important: this is a very simple variable renaming function. Always make a backup of your scripts. 
# This is not a production-ready variable refactoring solution.

# When you rename variables, there may be other parts of your script that would also need to be updated. 
# For example, when a variable is a function parameter, 
# then all calls to that function would also need to change their parameter name.

###########################################################################################################################################################################################

# Tip 68: Getting a Variable Inventory


# For documentation purposes, you may want to get a list of all variables that a PowerShell script uses.

# Here is a function called Get-Variable:

function Get-Variable
{
    $token = $null
    $errors = $null

    $inputText = $psise.CurrentFile.Editor.Text

    $ast = [System.Management.Automation.Language.Parser]::ParseInput($inputText, [ref] $token, [ref] $errors)
  
    # not complete, add variables you want to exclude from the list:
    $systemVariables = '_', 'null', 'psitem', 'true', 'false', 'args', 'host'

    $null = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)

    $token | Where-Object { $_.Kind -eq "Variable" } | Select-Object -ExpandProperty Name | 
             Where-Object { $systemVariables -notcontains $_ } | Sort-Object -Unique
}

# Simply load a script into the built-in ISE editor, then from the interactive console, run Get-Variable.

# You will get a sorted list of all variables used by the currently opened script.

# If you replace “$psise.CurrentFile.Editor.Text” by a variable that contains script code, 
# you can run this function outside the ISE editor, as well. 
# Simply use Get-Content to load the contents of any script into a variable, and then use this variable in place of the code above.

###########################################################################################################################################################################################

# Tip 69: Finding Changeable Properties


# When you get back results from PowerShell cmdlets, the results are objects and have properties. 
# Some properties can be changed, others are read-only.

# Here is a simple trick to find out the object properties that you can actually change. 
# The code uses the process object of the current PowerShell host, but you can use any cmdlet result.

$myProcess = Get-Process -Id $pid

$myProcess | Get-Member -MemberType Properties | Out-String -Stream | Where-Object { $_ -like "*set;*" }

# Output:
#        EnableRaisingEvents        Property       bool EnableRaisingEvents {get;set;}                                                
#        MaxWorkingSet              Property       System.IntPtr MaxWorkingSet {get;set;}                                             
#        MinWorkingSet              Property       System.IntPtr MinWorkingSet {get;set;}                                             
#        PriorityBoostEnabled       Property       bool PriorityBoostEnabled {get;set;}                                               
#        PriorityClass              Property       System.Diagnostics.ProcessPriorityClass PriorityClass {get;set;}                   
#        ProcessorAffinity          Property       System.IntPtr ProcessorAffinity {get;set;}                                         
#        Site                       Property       System.ComponentModel.ISite Site {get;set;}                                        
#        StartInfo                  Property       System.Diagnostics.ProcessStartInfo StartInfo {get;set;}                           
#        SynchronizingObject        Property       System.ComponentModel.ISynchronizeInvoke SynchronizingObject {get;set;}

###########################################################################################################################################################################################

# Tip 70: Finding Files plus Errors


# When you use Get-ChildItem to recursively search directory paths for files, 
# you may stumble across subfolders where you do not have enough privileges. To suppress errors, you may use –ErrorAction SilentlyContinue.

# That’s fine and good practice, but maybe you’d like to get a list of the folders that you actually had no access to, too.

# Here is a script that searches for all PowerShell scripts within the Windows folder. It stores these files in $PSScripts. 
# At the same time, it logs all errors in the variable $ErrorList and lists all folders that were inaccessible: 

$PSScripts = Get-ChildItem -Path c:\windows -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue -ErrorVariable ErrorList

$ErrorList
# Output:
#        Get-ChildItem : Access to the path 'C:\windows\AppCompat\Programs' is denied.
#        At line:1 char:14
#        + $PSScripts = Get-ChildItem -Path c:\windows -Filter *.ps1 -Recurse -ErrorAction  ...
#        + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#            + CategoryInfo          : PermissionDenied: (C:\windows\AppCompat\Programs:String) [Get-ChildItem], UnauthorizedAccessException
#            + FullyQualifiedErrorId : DirUnauthorizedAccessError,Microsoft.PowerShell.Commands.GetChildItemCommand


$ErrorList | ForEach-Object {

    Write-Warning ('Access denied: ' + $_.CategoryInfo.TargetName)
}

# Output: 
#        WARNING: Access denied: C:\windows\AppCompat\Programs

###########################################################################################################################################################################################

# Tip 71: Reading Registry Values with Type


# Reading all registry values is simple when you do not need the data type: simply use Get-ItemProperty:

Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run




# If you do need the data type, a little more effort is needed:

$key = Get-Item -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run

$a = $key.GetValueNames() | ForEach-Object {

    $ValueName = $_

    $rv = 1 | Select-Object -Property Name, Type, Value

    $rv.Name = $ValueName
    $rv.Type = $key.GetValueKind($ValueName)
    $rv.Value = $key.GetValue($ValueName)

    $rv
}



# Method Two:

($key = Get-Item -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run).GetValueNames() | ForEach-Object {

    New-Object PSObject -Property @{
    
        Name = $_
        Type = $key.GetValueKind($_)
        Value = $key.GetValue($_)

    } | Select-Object Name, Type, Value

} | Format-Table -AutoSize


# Another way to create new objects

$object = [PSCustomObject]@{

    Name = "Silence"
    ID = 123
    Active = $true
}
$object


# Method Three:

(Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run).PSObject.Properties | 
    Where-Object { $_.Name -notlike "PS*" } | Select-Object Name, TypeNameOfValue, Value | Format-Table -AutoSize

###########################################################################################################################################################################################

# Tip 72: Getting US ZIP Codes


$webservice = New-WebServiceProxy -Uri 'http://www.webservicex.net/uszip.asmx?WSDL'

$webservice.GetInfoByCity("New York").Table
# Output:
#        CITY      : New York
#        STATE     : NY
#        ZIP       : 10001
#        AREA_CODE : 212
#        TIME_ZONE : E
#        ...


$webservice.GetInfoByZIP("10286").Table
# Output:
#        CITY      : New York
#        STATE     : NY
#        ZIP       : 10286
#        AREA_CODE : 212
#        TIME_ZONE : E

###########################################################################################################################################################################################

# Tip 73: WMI Search Tool


# WMI is a great and powerful technique: simply specify a WMI class name, and back you get all the instances of that class:

Get-WmiObject -Class Win32_BIOS

# How do you know the WMI classes, though? Here is a search tool:

function Find-WMIClass
{
    param
    (
        [Parameter(Mandatory = $true)]
        $SearchTerm = "Resolution"
    )

    Get-WmiObject -Class * -List | Where-Object { $_.Properties.Count -gt 3 } | 

        Where-Object { $_.Name -notlike "Win32_Perf*" } | Where-Object {
    
            $ListOfNames = $_.Properties | Select-Object -ExpandProperty Name

            ($ListOfNames -like "*$SearchTerm*") -ne $null

        } | Sort-Object -Property Name
}

# Simply specify a search term you are after. 
# The code will find all WMI classes that contain a property with the search term in its name (use wildcards to widen the search).

# This will find all relevant WMI classes that have a property that ends with “resolution”:

Find-WMIClass -SearchTerm *resolution

# Output:
#           NameSpace: ROOT\cimv2
#        
#        Name               Methods                 Properties                                                                                                                                                  
#        ----               -------                 ----------                                                                                                                                                  
#        CIM_CacheMemory    {SetPowerState, R...    {Access, AdditionalErrorData, Associativity...}                                                                                               
#        CIM_CurrentSensor  {SetPowerState, R...    {Accuracy, Availability, Caption, ConfigManagerErrorCode...}                                                                                                
#        CIM_FlatPanel      {SetPowerState, R...    {Availability, Caption, ConfigManagerErrorCode...}                                                                                 
#        CIM_Memory         {SetPowerState, R...    {Access, AdditionalErrorData, Availability...}

# Next, pick a class name and look at the actual data

Get-WmiObject -Class CIM_CacheMemory | Select-Object -Property *

###########################################################################################################################################################################################

# Tip 74: Reading System Logs from File


# Sometimes, you may have to evaluate system log files that have been exported to disk, 
# or you want to read a system log file in “evtx” format directly from a file.

# Here is how you do this:

$path = "$env:windir\System32\Winevt\Logs\Setup.evtx"

Get-WinEvent -Path $path

###########################################################################################################################################################################################

# Tip 75: Enabling and Disabling PowerShell Remoting


# If you want to access a computer remotely via PowerShell, 
# then on the destination side (on the computer you want to visit), 
# run this line of code with Administrator privileges:

Enable-PSRemoting -SkipNetworkProfileCheck -Force

# Once you did, you can now visit this computer from another box – provided you have local Administrator privileges on the target machine, 
# you do specify the computer name and not its IP address, and both computers are joined in the same domain.

# To connect interactively, use this line:
Enter-PSSession -ComputerName $targetComputerName

# To run code remotely, try this:
Invoke-Command -ScriptBlock { Get-Service } -ComputerName $targetComputerName


# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Using PowerShell Remoting without Domain


# By default, when you enable PowerShell remoting via Enable-PSRemoting, then only Kerberos authentication is enabled. 
# This requires both computers to be in the same domain (or trusted domains), 
# and it only works when you specify computer names (possibly including domain suffixes).
# It will not work across domains, outside domains, or with IP addresses.

# To make this work, you need to make one change on the computer that initiates the remoting. 
# In a PowerShell console with Administrator privileges, enter this:

Set-Item WSMan:\localhost\Client\TrustedHosts -Value * -Force

# If that path is not available, you may have to first (temporarily) enable PowerShell remoting on that machine
# (using Enable-PSRemoting –SkipNetworkProfileCheck –Force).

# Once you made the change, you now can authenticate using NTLM, too. 
# Just remember that now, with domain-joined computers, 
# you need to always submit the –Credential parameter and specify a username and enter a password. 


# a simple way to undo this change
Clear-Item WSMan:\localhost\Client\TrustedHosts -Confirm

###########################################################################################################################################################################################

# Tip 76: Faking Object Type


# The internal PowerShell ETS is responsible for converting objects to text. 
# To do this, it looks for a property called “PSTypeName”. 
# You can add this property to your own objects to mimic another object type and make the ETS display your object in the same way:

$object = [PSCustomObject]@{

    ProcessName = "notepad"
    ID = -1
}

$object | ft -AutoSize
# Output:
#        ProcessName ID
#        ----------- --
#        notepad     -1


# The object pretends to be a process object, and ETS will format it accordingly:

$object2 = [PSCustomObject]@{

    ProcessName = "notepad"
    ID = -1
    PSTypeName = "System.Diagnostics.Process"
}

$object2 | ft -AutoSize
# Output:
#        Handles NPM(K) PM(K) WS(K) VM(M) CPU(s) Id ProcessName
#        ------- ------ ----- ----- ----- ------ -- -----------
#                     0     0     0     0        -1 notepad

# When cmdlets accept objects via pipeline, it can be useful to create just the object type they want, 
# and submit dynamic data or anything really you want. So faking objects can be useful to create cmdlet food.

###########################################################################################################################################################################################

# Tip 77: Controlling Execution of Executables


# PowerShell treats executables (files with extension EXE) like any other command. 
# You can, however, make sure that PowerShell will not execute any or execute only a list of approved applications.

# The default setting allows any EXE to be executed:

$ExecutionContext.SessionState.Applications       # Output: *

# This setting would make sure only ping.exe and regedit.exe can run:

$ExecutionContext.SessionState.Applications.Clear()
$ExecutionContext.SessionState.Applications.Add("ping.exe")
$ExecutionContext.SessionState.Applications.Add("regedit.exe")

$ExecutionContext.SessionState.Applications
# Output:
#        ping.exe
#        regedit.exe


# Obviously, you can simply revert this setting to get back the default behavior:
$ExecutionContext.SessionState.Applications.Add("*")

$ExecutionContext.SessionState.Applications
# Output:
#        ping.exe
#        regedit.exe
#        *

# So as-is, this setting will just make it harder to execute EXEs (or prevent accidental execution of unwanted EXEs). 
# To use it as a security measure, you would also need to turn off the so-called “Language Mode”. 

#　When turned off, you no longer can access .NET objects directly, 
# thus you would not be able to revert the change anymore in the current PowerShell session. 


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Turning Off “FullLanguage” Mode


# PowerShell can be restricted in many ways. One is to set the language mode from “FullLanguage” to “RestrictedLanguage”. 
# It is a way of no return, at least unless you close and re-open PowerShell:

$Host.Runspace.SessionStateProxy.LanguageMode      # Output: FullLanguage  [default value for powershell]

$Host.Runspace.SessionStateProxy.LanguageMode = "RestrictedLanguage"


function abc
{
    write-host "212"
}
# Exception here: Function declarations are not allowed in restricted language mode or a Data section.


$Host.Runspace.SessionStateProxy.LanguageMode = "FullLanguage"
# Exception here: Property references are not allowed in restricted language mode or a Data section.

# Once set to “RestrictedLanguage”, PowerShell will only execute commands. 
# It will no longer execute object methods or access object properties, and you can no longer define new functions.

# So RestrictedLanguage basically is a safe lockdown where PowerShell can execute commands 
# but cannot dive into low level .NET or override existing commands with newly created functions.


# Note: re-open Powershell to undo the change.

###########################################################################################################################################################################################

# Tip 78: Creating Colorful HTML Reports


# To turn results into colorful custom HTML reports, 
# simply define three script blocks: one that writes the start of the HTML document, 
# one that writes the end, and one that is processed for each object you want to list in the report.

# Then, hand over these script blocks to ForEach-Object. It accepts a begin, a process, and an end script block.

# Here is a sample script that illustrates this and creates a colorful service state report:

$path = "$env:temp\report.hta"

$begin = {

 @'
    <html>
    <head>
    <title>Report</title>
    <STYLE type="text/css">
        h1 {font-family:SegoeUI, sans-serif; font-size:20} 
        th {font-family:SegoeUI, sans-serif; font-size:15} 
        td {font-family:Consolas, sans-serif; font-size:12} 

    </STYLE>

    </head>
    <image src="http://www.yourcompany.com/yourlogo.gif" />
    <h1>System Report</h1>
    <table>
    <tr><th>Status</th><th>Name</th></tr>
'@
}

$process = {

    $status = $_.Status
    $name = $_.DisplayName

    if ($status -eq 'Running')
    {
        '<tr>'
        '<td bgcolor="#00FF00">{0}</td>' -f $status
        '<td bgcolor="#00FF00">{0}</td>' -f $name
        '</tr>'
    }
    else
    {
        '<tr>'
        '<td bgcolor="#FF0000">{0}</td>' -f $status
        '<td bgcolor="#FF0000">{0}</td>' -f $name
        '</tr>'
    }    
}

$end = {

 @'
    </table>
    </html>
    </body>
'@    
}

Get-Service | ForEach-Object -Begin $begin -Process $process -End $end | Out-File -FilePath $path -Encoding utf8

Invoke-Item -Path $path

###########################################################################################################################################################################################

# Tip 79: Accessing SQLServer Database


# You are running an SQL Server? Then here is a PowerShell script template you could use to run an SQL query and retrieve the data. 
# Simply make sure you fill in the correct user details, server address, and SQL statement:

$Database                       = 'Name_Of_SQLDatabase'
$Server                         = '192.168.100.200'
$UserName                         = 'DatabaseUserName'
$Password                       = 'SecretPassword'

$SqlQuery                       = 'Select * FROM TestTable'

# Accessing Data Base
$SqlConnection                  = New-Object -TypeName System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Data Source=$Server;Initial Catalog=$Database;user id=$UserName;pwd=$Password"

$SqlCmd                         = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText             = $SqlQuery
$SqlCmd.Connection              = $SqlConnection

$SqlAdapter                     = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand       = $SqlCmd

$set                            = New-Object data.dataset

# Filling Dataset
$SqlAdapter.Fill($set)

# Consuming Data
$Path = "$env:temp\report.hta"
$set.Tables[0] | ConvertTo-Html | Out-File -FilePath $Path

Invoke-Item -Path $Path  


# A useful function for sql handle: https://github.com/RamblingCookieMonster/PowerShell/blob/master/Invoke-Sqlcmd2.ps1

###########################################################################################################################################################################################

# Tip 80: Reading In PFX-Certificate


# When you use Get-PfxCertificate, you can read in PFX certificate files and use the certificate to sign scripts. 
# However, the cmdlet will always interactively ask for the certificate password.

# Here is some logic that enables you to submit the password by script:

$path = "C:\temp\test.pfx"
$password = "password"

Add-Type -AssemblyName System.Security
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import($path, $password, "Exportable")

$cert

###########################################################################################################################################################################################

# Tip 81: Changing PowerShell Priority


# Maybe you’d like a PowerShell script to work in the background, 
# for example copy some files, but you do not want the script to block your CPU or interfere with other tasks.

# One way to slow down PowerShell scripts is to assign them a lower priority. Here is a function that can do the trick:

function Set-Priority
{
    param
    (
        [CmdletBinding()]
        [System.Diagnostics.ProcessPriorityClass]
        $Priority
    )

    $process = Get-Process -Id $pid
    $process.PriorityClass = $Priority
}

# To lower the priority of a script, call it like this:

Set-Priority -Priority BelowNormal

# You can change priority back to Normal anytime later, or even increase the priority to assign more power to a script
# – at the expense of other tasks, though, which may make your UI less responsive.

[System.Enum]::GetNames([System.Diagnostics.ProcessPriorityClass])

# Output:
#        Normal
#        Idle
#        High
#        RealTime
#        BelowNormal
#        AboveNormal

# About ProcessPriorityClass: http://msdn.microsoft.com/en-us/library/system.diagnostics.processpriorityclass

###########################################################################################################################################################################################

# Tip 82: Creating New Shares


# WMI can easily create new shares. Here is sample code that will create a local share:

$shareName = "NewShare"
$path = "C:\123"

if(!(Get-WmiObject -Class Win32_Share -Filter "name='$shareName'"))
{
    $share = [WMIClass]"Win32_Share"
    $share.Create($path, $shareName, 0).ReturnValue
}
else
{
    Write-Warning "Share $shareName exists already."
}


# You can also create shares on remote machines, 
# provided you have Admin privileges on the remote machine. To do that, simply add the complete WMI path like this:

$shareName = "NewShare"
$path = "C:\123"
$server = "remoteServer"

if(!(Get-WmiObject -Class Win32_Share -Filter "name='$shareName'" -ComputerName $server))
{
    $share = [WMIClass]"\\$server\root\cimv2:Win32_Share"
    $share.Create($path, $shareName, 0).RetrunValue
}
else
{
    Write-Warning "Share $shareName exists already."
}

###########################################################################################################################################################################################

# Tip 83: Using Notepad to Print Things


# To print a text-based file with the Notepad, try using this line 
# (replace the path to the text file with some path that is meaningful to you, or else you will print a rather long system log file):

Start-Process -FilePath notepad -ArgumentList "/P C:\Windows\WindowsUpdate.log"

###########################################################################################################################################################################################

# Tip 84: Importing and Installing Certificate


# To programmatically load a certificate from a file and install 
# it in a specific location inside the certificate store, have a look at this script:

$pfxPath = "C:\temp\test.pfx"
$password = "password"
[System.Security.Cryptography.X509Certificates.StoreLocation]$StoreLocation = "CurrentUser"
$StoreName = "root"

Add-Type -AssemblyName System.Security
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import($pfxPath, $password, "Exportable")

$Store = New-Object System.Security.Cryptography.X509Certificates.X509Store($StoreName, $StoreLocation)
$Store.Open("ReadWrite")
$Store.Add($cert)
$Store.Close()

# You can configure the script and specify the location and password of the certificate file to import. 
# You can also specify the store location (CurrentUser or LocalMachine), 
# and the container to put the certificate into 
# (for example “root” for trustworthy root certificates, or “my” for personal certificates).

###########################################################################################################################################################################################

# Tip 85: Invoke-Expression is Evil


# Try and avoid Invoke-Expression in your scripts. This cmdlet takes a string and executes it as if it was a command. 
# In most scenarios, it is not needed, but introduces many risks.

# Here is a--somewhat constructed--show case:

function Test-BadBehavior($Path)
{
    Invoke-Expression "Get-Childitem -Path $Path"
}

# This function uses Invoke-Expression to run a command and append a parameter value. 
# So it would return a folder listing for the path entered as parameter.

# Since Invoke-Expression accepts any string, you open up your environment to “SQL-Injection”-like attacks. 
# Try running the function like this:

Test-BadBehavior -Path "C:\;Get-Process"

# It now would also run the second command and list all running processes. 
# Invoke-Expression is often used by attackers when they download the evil payload as string from some outside URL and then execute the code on-the-fly.

# Of course, Invoke-Expression would not have been necessary in the first place. 
# It is hardly ever needed in typical production scripts. Make sure you hardcode the commands you want to execute:

function Test-BadBehavior($Path)
{
    Get-ChildItem -Path $Path
}

###########################################################################################################################################################################################

# Tip 86: Use Out-Host instead of More


# Note that any of this will only work in a “real” console. It will not work in the PowerShell ISE.

# To output data page by page, in the PowerShell console many users pipe the result to more.com, like in the old days:

dir C:\Windows | more

# This seems to work well, until you start and pipe more than just a couple of data sets. Now PowerShell appears to hang:

dir C:\Windows -Recurse -ErrorAction SilentlyContinue | more

# That’s because more.com cannot work in real-time. It first collects all incoming data, then starts outputting it page by page.

# A much better way is to use the cmdlet Out-Host with the parameter –Paging:

dir C:\Windows -Recurse -ErrorAction SilentlyContinue | Out-Host -Paging

# It yields results immediately because it processes pipeline data as it comes in.

###########################################################################################################################################################################################

# Tip 87: Encrypting and Decrypting Files with EFS


# Provided EFS (Encrypting File System) is enabled on your system, and you are saving files to a NTFS location,
# then this is how you can encrypt any file and make sure only you can read it:

(Get-Item -Path "C:\temp\test.txt").Encrypt()

# If encryption succeeds, the file will now be displayed with green instead of black labels in explorer.exe. 



# Use Decrypt() instead of Encrypt() to undo the encryption.

(Get-Item -Path "C:\temp\test.txt").Decrypt()

# Note that EFS may have to be set up first, and that your company may require a centrally stored backup key for encryption.

###########################################################################################################################################################################################

# Tip 88: Functions Always Beat Cmdlets


# Functions always have higher rank than cmdlets, so if both are named alike, the function wins. 

# This function would effectively change the behavior of Get-Process: 

function Get-Process
{
    "go away"
}

Get-Process            # Output: go away


function Microsoft.PowerShell.Management\Get-Process
{
  'go away'
}

Microsoft.PowerShell.Management\Get-Process -Id $pid    # Output: go away


# The same applies to Aliases. They rank even above functions.

# The only way of making sure you are running the cmdlet would be accessing the module, 
# picking the wanted cmdlet, and directly invoking it:

$module = Get-Module Microsoft.PowerShell.Management
$cmdlet = $module.ExportedCmdlets["Get-Process"]
& $cmdlet

# Or, simply make sure no one has fiddled with your PowerShell environment, 
# by starting a fresh PowerShell and making sure you use the -noprofile parameter.

###########################################################################################################################################################################################

# Tip 89: Watch Rick Astley Dance and Sing!


# Before you try this, you may want to click the icon in the upper left corner of the PowerShell title bar, go to properties, and choose a small font.

# Next, try and run this command:

(New-Object Net.WebClient).DownloadString("http://bit.ly/e0Mw9w")


# As you will see, it downloads an entire PowerShell script. 
# Now – if you trust that code – you could make PowerShell immediately execute it by using the (dangerous) Invoke-Expression. 
# It’s dangerous because it can allow hackers to download and immediately execute all kinds of stuff. 
# This example is benign, though. Just make sure you run it in a PowerShell console, not the PowerShell ISE! You may also crank up the volume.

Invoke-Expression (New-Object Net.WebClient).DownloadString("http://bit.ly/e0Mw9w") 

###########################################################################################################################################################################################

# Tip 90: Getting Help


# Provided you have downloaded PowerShell help via Update-Help, 
# you can create yourself an excellent help topic viewer with just one line of code:

Get-Help about* | Out-GridView -PassThru | Get-Help -ShowWindow

# This will display a grid view with all about topics to choose from. Select one, and click OK, to view the help file.

###########################################################################################################################################################################################

# Tip 91: Creating HTML Colors


# To convert decimal color values to a hexadecimal representation, 
# like the one used in HTML, try this line:

'#{0:x2}{1:x2}{2:x2}{3:x2}' -f 255,202,81,0        # Output: #ffca5100

'#{0:x2}{1:x2}{2:x2}{3:x2}' -f 255,0,121,204       # Output: 

# The first value sets the transparency, followed by the byte values for red, green, and blue.

###########################################################################################################################################################################################

# 92: Converting Error Numbers


# Error numbers that are returned by Windows API calls often appear as very large negative numbers. 
# To give meaning to these numbers, convert them to hexadecimal values like this:

$errornumber = -2146828235

'0x{0:x}' -f $errornumber           # Output: 0x800a0035

# To find out the reason for such an error, 
# when you now search for the hexadecimal value, chances are much higher that you find a match.


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Looking Up Cryptic Error Codes


# Often, WMI and API calls return cryptic numeric error codes. To find out what went wrong, try this little helper function:

function Get-HelpForErrorCode([String]$Code)
{
    if($Code.StartsWith("-"))
    {
        $Code = "{0:X8}" -f ([Int32]$Code)
    }

    $url = "http://www.computerperformance.co.uk/Logon/code/code_$Code.htm"

    Start-Process -FilePath $url
}

# You can submit any decimal or hexadecimal error code, 
# and the function opens the appropriate web page where the error is explained (if information is available):

Get-HelpForErrorCode -Code -2146828235

# Link to site: http://www.computerperformance.co.uk/Logon/code/code_800A0035.htm

###########################################################################################################################################################################################

# Tip 93: Using Cmdlets to Manage Virtual Hard Drives


# Both Windows 8.1 and Server 2012 R2 come with a vast number of additional cmdlets, 
# some of which can be used to manage virtual disks. However, before you can find and use these cmdlets,
# you need to activate the “Hyper-V role” 
# (note that Hyper-V support on the client side requires Windows 8.1 Pro or Enterprise. It is not included in the “Home” versions).

# In Windows 8.1, you need to do this manually: go to Control Panel,
# and then to Programs/Programs and Features. You can also enter “appwiz.cpl” in PowerShell to get there.

# Next, click “Turn Windows features on or off”. This opens up a dialog with all available features. 
# Identify the node “Hyper-V”, and enable it. Then click OK. If the node “Hyper-V” is missing, 
# then your version of Windows does not support Hyper-V on the client side. 
# If the option “Hyper-V Platform” is grayed out, then you need to enable virtualization support in your computer BIOS settings.

# The feature installation takes a couple of seconds. Once it is completed, you have a whole new bunch of cmdlets available:

Get-Command -Module Hyper-V

###########################################################################################################################################################################################

# Tip 94: Join-Path Fails with Nonexistent Drives


# To construct path names from parent folders and files, you may have been using Join-Path. 
# This cmdlet takes care of the correct number of backslashes when you combine path components:

$part1 = "C:\Windows\"
$part2 = "\myfile.txt"
$result = Join-Path -Path $part1 -ChildPath $part2

$result   # Output: C:\Windows\myfile.txt



# However, Join-Path will fail if the path components do not exist. 
# So you cannot create a path for a drive that is not mounted:

$part1 = "L:\Windows\"
$part2 = "\myfile.txt"
$result = Join-Path -Path $part1 -ChildPath $part2

$result   # Exception here: Join-Path : Cannot find drive. A drive with the name 'L' does not exist.



# In essence, what Join-Path does can be done manually as well. 
# This will combine two path segments and take care of backslashes:

$part1 = "L:\Windows\"
$part2 = "\myfile.txt"
$result = $part1.TrimEnd("\") + "\" + $part2.TrimStart("\")

$result   # Output: L:\Windows\myfile.txt

###########################################################################################################################################################################################

# Tip 96: Finding Out Windows Version


# Do you own Windows 8.1 Basic, Pro, or Enterprise? Finding out the Windows version is easy. 
# Finding out the exact subtype is not so trivial.

# At best, you may get the SKU number which tells you exactly the Windows version you have, 
# but it’s again not trivial to translate back the number to a meaningful name:

Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty OperatingSystemSKU   # Output: 4


# A better way may be this line which returns a clear text description of the license type you are using:

Get-WmiObject SoftwareLicensingProduct -Filter 'Name like "Windows%" and LicenseStatus=1' | Select-Object -ExpandProperty Name 
# Output:
#        Windows(R) 7, Enterprise edition


# Another approach could be this which will include the major Windows version as well:

Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty Caption
# Output:
#        Microsoft Windows 7 Enterprise 


###########################################################################################################################################################################################

# Tip 97: Reading Disks and Partitions


# Disk management has been greatly simplified with the many new client and server cmdlets that ship with Windows 8.1 and Server 2012 R2.

# Let’s start with looking at disks and partitions. This would list all disks you have mounted:
Get-Disk


# And this gets you the partitions:
Get-Partition


# Both cmdlets reside in the module “Storage”:
Get-Command -Name Get-Disk | Select-Object -ExpandProperty Module


# This will show you all the other storage management commands found there:
Get-Command -Module Storage

###########################################################################################################################################################################################

# Tip 98: Randomize Lists of Numbers


$data = 1, 2, 3, 5, 8, 13

# This line will take a list of numbers and randomize their order:
Get-Random -InputObject $data -Count ([int]::MaxValue)


# Piping works too, but it's slower:
$data | Sort-Object -Property { Get-Random }

###########################################################################################################################################################################################

# Tip 99:　NULL Values in Arrays


$a = @()

$a += 1
$a += $null
$a += $null
$a += 2

$a
# Output:
#        1
#        2

$a.count    # Output: 4


# Note: Whenever you assign NULL values to array elements, they will count as array elements, 
# but will not be output (after all, they are NULL aka nothing). 
# This can lead to tough debugging situations, so when the size of an array does not seem to match the content, look for NULL values

###########################################################################################################################################################################################

# Tip 100: Steps to Configure PowerShell 


# Steps to Configure PowerShell 
$PSVersionTable.PSVersion.Major         # Output: 3



# To enable script execution
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force

# Note: You will now be able to run script located anywhere. 
# If you are a beginner and want some extra protection, replace “Bypass” with “RemoteSigned”. 
# This will keep you from running PowerShell scripts that were downloaded from the Internet or received as email attachment.
# It will also prevent you from running scripts outside your own domain, though.


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# If you use PowerShell at home or in an unmanaged environment, 
# here are some additional steps you should consider to make PowerShell fully functional.

# To allow Administrators to connect to your machine remotely and run cmdlets like Get-Process or Get-Service, 
# you may want to enable the Remote Administration firewall exception. 
# Open PowerShell with Administrator privileges, and run this:

netsh firewall set service remoteadmin enable

# Output:
#        IMPORTANT: Command executed successfully.
#        However, "netsh firewall" is deprecated;
#        use "netsh advfirewall firewall" instead.
#        For more information on using "netsh advfirewall firewall" commands
#        instead of "netsh firewall", see KB article 947709
#        at http://go.microsoft.com/fwlink/?linkid=121488 .
#        
#        Ok.



# The command returns that there is a newer command and that it is deprecated, 
# but it will still work and enable the firewall exception. 
# The newer command is much harder to use because its parameters are localized, 
# and you would need to know the exact names of the firewall exceptions.

# To really use the remoting capabilities of cmdlets,
# you would also have to start the RemoteRegistry service and set it to auto start:

Start-Service RemoteRegistry

Set-Service -Name RemoteRegistry -StartupType Automatic

# Now you can use Get-Process, Get-Service, or other cmdlets that expose a –ComputerName parameter to connect to your computer remotely, 
# provided the user running these cmdlets has Administrator privileges on your system.

# In a simple peer-to-peer home environment, it would be sufficient to set up Administrator accounts with the same name on each computer.


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# To use the PowerShell remoting feature against your own machine, you need to enable PowerShell remoting on your machine. 
# To do this, start PowerShell with full Administrator privileges, and run this command:

Enable-PSRemoting -SkipNetworkProfileCheck -Force

# Note that the parameter -SkipNetworkProfileCheck was introduced in PowerShell 3.0. 
# If you are still using PowerShell 2.0, omit this parameter. 
# You would then have to manually temporarily disable public network adapters 
# if PowerShell complains about public network connections being present.

# The command enables PowerShell remoting on your machine. Others can now connect to your computer, 
# provided they are members of the Administrators group on your machine. 

# However, you would only be able to connect to others using Kerberos authentication. 
# So at this point, remoting would only work for domain environments. 
# If you operate a simple peer-to-peer network or want to use remoting across different domains, then enable NTLM authentication. 
# Important: this is a setting that needs to be set on the client side: 
# Not on the machine you want to connect to, but on the machine that you start your remote call:

Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value * -Force

# Using "*" allows you to contact any target machine via NTLM authentication. 
# Since NTLM is a non-mutual authentication, NTLM can impose risks when you authenticate against a non-trusted and possibly compromised host. 
# So instead of "*", you could also specify an IP address or the beginning of an IP address such as "10.10.*".


# Once PowerShell remoting is set up, you can start to play.

# This line would run arbitrary PowerShell code on the machine ABC 
# (and requires that you first have enabled remoting on machine ABC and that you have Administrator privileges on ABC):

Invoke-Command -ScriptBlock { "Hello" > C:\IwasHERE.txt } -ComputerName ABC


# This would do the same, but here you would explicitly specify credentials.
# When you specify an account, always make sure you specify domain and username. 
# If it is not a domain account, specify computer name and username:

Invoke-Command -ScriptBlock { "Hello" > C:\IwasHERE.txt } -ComputerName ABC -Credential domain\user

###########################################################################################################################################################################################