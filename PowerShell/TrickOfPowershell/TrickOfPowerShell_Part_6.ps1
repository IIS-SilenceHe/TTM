# Reference site: http://powershell.com/cs/blogs/tips/
###########################################################################################################################################################################################

# Tip 1: Using Functions inside the Pipeline

function Test-Pipeline([Parameter(ValueFromPipeline=$true)]$incoming)
{
    process
    {
        "Working with $incoming"
    }
}

Get-Process | Test-Pipeline

# Code inside PowerShell functions always is placed in the begin, process, or end block. 
 # You may not have noticed this because when you don't do this yourself, PowerShell places your entire function code inside an end block.
 # When you try and run a function inside the pipeline, it becomes important to manually place the code into the appropriate script block. 
 # Only code inside the process block is executed for each incoming pipeline element. 
 # If you leave the code in the end block, only the last pipeline data element will be executed.

###########################################################################################################################################################################################

# Tip 2: Creating Pipeline Filters

# Here is a filter that will only display processes that have a "MainWindowTitle" text, 
 # thus filtering all running processes and showing only those that have an application window. It then displays the three desktop applications with the most CPU usage

function Filter-ApplicationProgram([Parameter(ValueFromPipeline=$true)]$incoming)
{
    process
    {
        if($incoming.MainWindowTitle -ne "")
        {
            $incoming
        }
    }
}

Get-Process | Filter-ApplicationProgram | Sort-Object CPU -Descending | Select-Object CPU, Name, MainWindowTitle -First 3

###########################################################################################################################################################################################

# Tip 3: HTML-Scraping with RegEx

# To scrape valuable information from websites with PowerShell you can download the HTML code and then use regular expressions to extract what you are after

$webClient = New-Object System.Net.WebClient

$html = $webClient.DownloadString("http://www.cnn.com") | Out-String

$headerPattern = '(?i)<h1>(.*?)</h1>'

$header = ([regex]$headerPattern).Matches($html) | ForEach-Object {$_.Groups[1].Value}

$header                 # It downloads the HTML content from www.cnn.com and then extracts all <h1>...</h1> headers. That way, you get a quick headline overview
# Output:
#        <a href="/2014/09/16/living/spanking-cultural-roots-attitudes-parents/index.html?hpt=hp_t1" target="">To spank or not spank? </a>
#        <a href="/2014/09/16/living/spanking-cultural-roots-attitudes-parents/index.html?hpt=hp_t1" target="">Answer depends on who you are, where you live</a>

###########################################################################################################################################################################################

# Tip 4: Finding Code Signing Certificates

# To digitally sign PowerShell scripts, you need a certificate with the purpose "CodeSigning". 

dir Cert:\CurrentUser\My -CodeSigningCert                               # find out which code signing certificates are available to you - if any

# If you have more than one code signing certificate listed, use Where-Object to pick the one you want to use. 
$cert = dir cert:\CurrentUser\my -CodeSigningCert | Where-Object { $_.Subject -like '*Henry*' }

dir Cert:\CurrentUser\My -CodeSigningCert | Select-Object -Property *   # To view all certificate properties, use Select-Object

###########################################################################################################################################################################################

# Tip 5: Digitally Signing PowerShell Scripts

# sign all PowerShell scripts in folder c:\scripts with the first available code signing certificate from your certificate store:

$cert = dir Cert:\CurrentUser\My -CodeSigningCert | Select-Object -First 1

if($cert)
{
    dir C:\myScripts -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue | Set-AuthenticodeSignature -Certificate $cert
}
else
{
    Write-Warning 'You do not have a digital certificate for code signing.'

}

# Note that there are two situations when Set-AuthenticodeSignature cannot sign a script: if a script is smaller than 5 Bytes, 
 # and if the script was saved with "Unicode Big Endian" encoding which occurs when you save a script with the PowerShell ISE editor.

###########################################################################################################################################################################################

# Tip 6: PowerShell Script Security Auditing

function Test-PSScript($path = "C:\", [switch]$unsafeOnly)
{
    Get-ChildItem $path -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue | Get-AuthenticodeSignature | 
        
        Where-Object {($_.Status -ne "Valid") -or ($unsafeOnly -eq $false)} | ForEach-Object {

            $result = $_ | Select-Object Path, Status
        
            switch($_.Status)
            {
                'notsigned' { $result.Status = 'no digital signature present, unsafe script.' }
                'unknownerror' { $result.Status = 'script author is not trusted by your organization.' }
                'hashmismatch' { $result.Status = 'script content has been manipulated.' }
                'valid' { $result.Status = 'trusted script in original condition.' }
            }

            $result
        }
}

###########################################################################################################################################################################################

# Tip 7:　Setting Breakpoints in PowerShell Scripts

# Most sophisticated PowerShell editors have built-in debugging support. PowerShell can handle breakpoints natively, too

# Next time you run this script from within your PowerShell session, it will stop at line 4, and you find yourself in a debugging mode where you can access (and change) all variables
Set-PSBreakpoint -Script c:\scripts\somescript.ps1 -Line 4

# To make the script continue, use the command "continue". 

Get-PSBreakpoint | Remove-PSBreakpoint                               # To remove all breakpoints

###########################################################################################################################################################################################

# Tip 8: Creating Intelligent Variables

# To create variables that calculate their content each time they are accessed, use access-triggered variable breakpoints.
$Global:Now = Set-PSBreakpoint -Variable Now -Mode Read -Action {$Global:Now = Get-Date}

# Try it: each time you output $today, it returns the correct date and time - because each time you access it, the associated script block updates its content.

###########################################################################################################################################################################################

# Tip 9: Controlling Automatic Updates

# To control whether Windows download and/or installs updates silently or prompts for permission, use this script and set the appropriate NotificationLevel via script. 
 # Just make sure you run this code with full Administrator privileges. 
 # This code will set Automatic Downloads Notifications to level 3, so Windows will prompt you before it actually installs updates:

# run with full Admin privileges!
$updateObj = New-Object -ComObject Microsoft.Update.AutoUpdate
# ' 1 =  Never Check for Updates
# ' 2 =  Prompt for Update and Prompt for Installation
# ' 3 =  Prompt for Update and Prompt for Installation
# ' 4 =  Install automatically
$updateObj.Settings.NotificationLevel = 3
$updateObj.Settings.Save()

###########################################################################################################################################################################################

# Tip 10: Get Automatic Updates Installation Time

$updateObj = New-Object -ComObject Microsoft.Update.AutoUpdate
$day = $updateObj.Settings.ScheduledInstallationDay
$hour = $updateObj.Settings.ScheduledInstallationTime
$level = $updateObj.Settings.NotificationLevel

if($level -eq 4)
{
    if($day -eq 0)
    {
        $weekDay = "Every Day"
    }
    else
    {
        $weekDay = [System.DayOfWeek]($day - 1)
    }

    "Automatic updates installed $weekday at $hour o'clock."
}
else
{
    'Updates will not be installed automatically. Check update settings for more info.'
}

# To double-check settings or change them via UI, open the appropriate control like this:
$updateObj = New-Object -ComObject Microsoft.Update.AutoUpdate
$updateObj.ShowSettingsDialog()

###########################################################################################################################################################################################

# Tip 11: Bypassing Execution Policy

# When execution policy prevents execution of PowerShell scripts, you can still execute them. There is a secret parameter called "-" . 
 # When you use it, you can pipe a script into powershell.exe and execute it line by line:

Get-Content 'C:\somescript.ps1' | powershell.exe -noprofile -    

###########################################################################################################################################################################################

# Tip 12: Scanning Registry for ClassIDs

# The Windows Registry is a repository filled with various Windows settings. Get-ItemProperty can read Registry values and accepts wildcards. 
 # So, with Get-ItemProperty, you can create tools to find and extract Registry information. 

function Get-ClassID($fileName = "vbscript.dll")
{
    Write-Warning 'May take some time to produce results...'
    
    Get-ItemProperty 'HKLM:\Software\Classes\CLSID\*\InprocServer32' | Where-Object {$_."(Default)" -like "*$FileName"} | ForEach-Object{
    
        if($_.PSPath -match "{.*}")
        {
            $Matches[0]
        }
    }
}

# It finds all keys called InprocServer32 that are located one level below the CLSID-key, 
 # then checks whether the (Default) value of that key contains the file name you are looking for. A regular expression then extracts the ClassID from the path of the registry key.

###########################################################################################################################################################################################

# Tip 13: Creating PowerShell Menus

$title = "Reboot System Now"
$message = "Do you want to reboot your machine now?"

$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Reboots the system now."
$no  = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Does not reboot the system. You will have to reboot manually later."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

$result = $Host.UI.PromptForChoice($title, $message, $options, 0)

switch($result)
{
    0 {
        "You selected Yes."
        Restart-Computer -whatif  
      }

    1 {"You selected No."}
}

###########################################################################################################################################################################################

# Tip 14: Using Wildcards with Environment Variables

dir env:*user*
# Output:
#         Name                           Value                                                                                                                                                                                 
#         ----                           -----                                                                                                                                                                                 
#         USERDNSDOMAIN                  xxx.xxx.xxx.xxx                                                                                                                                                            
#         USERPROFILE                    C:\Users\v-sihe                                                                                                                                                                       
#         ALLUSERSPROFILE                C:\ProgramData                                                                                                                                                                        
#         USERNAME                       v-sihe                                                                                                                                                                                
#         USERDOMAIN                     xxx                                                                                                                                                                               
#         GIT_USERNAME                   silence                                                                                                                                                                               
#         USERDOMAIN_ROAMINGPROFILE      xxx

"UserName: $env:username"                                        # Output: UserName: v-sihe
                                                                 
                                                                 
dir env:                                                         # use the env: PowerShell drive to list all Windows environment variables 
                                                                 
dir env: | Where-Object {$_.Value -like "*Program*"}             # to list environments with a keyword in its value

###########################################################################################################################################################################################

# Tip 15: Permanently Changing User Environment Variables

[environment]::SetEnvironmentVariable("Test", 12, "User")        # To create or change an environment variable in the user context

# This environment variable will keep the value until you change it or delete it - even across reboots. 
 # So, it can be used for communication between processes or to keep state across reboots.

[environment]::GetEnvironmentVariable("Test", "User")            # to read the variable
                                                                 
[environment]::SetEnvironmentVariable("Test", "", "User")        # to delete the variable

###########################################################################################################################################################################################

# Tip 16: Using Regular Expressions with Dir

dir $home -Recurse | Where-Object {$_.Name -match "\d.*?\."}     #  get you any file with a number in its filename, ignoring numbers in the file extension

# When you use Dir (alias: Get-ChildItem) to list folder contents, you can use simple wildcards but they do not give you much control. 
 # A much more powerful approach is to use regular expressions. Since Get-ChildItem does not support regular expressions, you can use Where-Object to filter the results returned by Dir

###########################################################################################################################################################################################

# tip 17: Formatting Currencies

# Formatting numbers as currencies is straight-forward - as long as it is your own currency format
"{0:C}" -f 12.22                                                   # Output: $12.22

# If you want to output currencies in other cultures, you still can do it. 
$culture = New-Object System.Globalization.CultureInfo("zh-CN")
(12.22).ToString("c", $culture)                                    # Output: ¥12.22


function Get-CultureInfo($keyword = "*")
{
    [System.Globalization.CultureInfo]::GetCultures("AllCultures") | Where-Object {$_.DisplayName -like "*$keyword*"}
}

Get-CultureInfo

Get-CultureInfo Chinese
# Output:
#         LCID             Name             DisplayName                                                                                                                                                                        
#         ----             ----             -----------                                                                                                                                                                        
#         4                zh-Hans          Chinese (Simplified)                                                                                                                                                               
#         1028             zh-TW            Chinese (Traditional, Taiwan)                                                                                                                                                      
#         2052             zh-CN            Chinese (Simplified, PRC)                                                                                                                                                          
#         3076             zh-HK            Chinese (Traditional, Hong Kong S.A.R.)                                                                                                                                            
#         4100             zh-SG            Chinese (Simplified, Singapore)                                                                                                                                                    
#         5124             zh-MO            Chinese (Traditional, Macao S.A.R.)                                                                                                                                                
#         30724            zh               Chinese                                                                                                                                                                            
#         31748            zh-Hant          Chinese (Traditional)                                                                                                                                                              
#         4                zh-CHS           Chinese (Simplified) Legacy                                                                                                                                                        
#         31748            zh-CHT           Chinese (Traditional) Legacy 
  
###########################################################################################################################################################################################

# Tip 18: Closing Excel Gracefully

'Excel processes: {0}' -f @(Get-Process excel -ea 0).Count           # Output: Excel processes: 0

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
Start-Sleep 5
$excel.Quit()
'Excel processes: {0}' -f @(Get-Process excel -ea 0).Count           # Output: Excel processes: 1


Stop-Process -Name excel                                             # try and kill that orphaned process

# However, you now would kill all open Excel instances, not just the one you launched from script. 
 # Also, killing a process with Stop-Process is hostile, and Excel may complain the next time you launch it that it was shut down unexpectedly.

# Here is how you tell the .NET Framework to quit Excel peacefully:
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)   # ReleaseComObject 方法：递减所提供的运行库可调用包装的引用计数
Start-Sleep 1
'Excel processes: {0}' -f @(Get-Process excel -ea 0).Count           # Output: Excel processes: 0

###########################################################################################################################################################################################

# Tip 19: Changing Units

# When you list folder contents, file sizes are in bytes. If you'd rather like to view them in MB or GB, you can use calculated properties, 
 # but by turning numbers into MB or GB, you turn them into text strings. That's bad because then you can no longer sort them or filter them based on size

# A better approach overrides the display function ToString() only. This way, the file size stays numeric but displays any way you like:

dir $env:windir | Select-Object Mode, LastWriteTime, Length, Name | ForEach-Object {

    if($_.Length -ne $null)
    {
        $_.Length = $_.Length | Add-Member ScriptMethod ToString {"{0:0.0} KB" -f ($this / 1KB)} -Force -PassThru
        
        $_
    }

} | Sort-Object Length


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------
# A much more friendly way here: [Reference]http://www.lucd.info/2011/11/06/friendly-units/

dir $env:windir | Select-Object Mode, LastWriteTime, Length, Name | ForEach-Object { 

if ($_.Length -ne $null) 
{
    $_.Length = $_.Length | Add-Member ScriptMethod ToString {
    
        $Units = "B","KB","MB","GB","TB","PB","EB","ZB","YB"
        $modifier = 0

        if($this -gt 0)
        {
            $modifier = [math]::Floor([Math]::Log($this,1KB))
        }

        ('{0,7:f2} {1,2}' -f ($this/[math]::Pow(1KB,$modifier)),(&{if($modifier -lt $units.Count){$units[$modifier]}else{"1KB E{0}" -f $modifier}}))} -Force -pass
    }
    
    $_

} | Sort-Object Length

# Output:
#         Mode               LastWriteTime                       Length                      Name                                                
#         ----               -------------                       ------                      ----                                                
#         d----              2014/08/25 10:31:12                                             Resources                                           
#         d----              2009/07/14 10:35:47                                             SchCache                                                                                                                                     
#         d----              2009/07/14 13:37:46                                             IME                                                 
#         d-r-s              2013/08/15 16:22:55                                             Fonts                                               
#         d----              2011/04/12 15:46:37                                             Globalization                                       
#         -a---              2012/10/11 18:02:09                    0.00  B                  ativpsrm.bin                                        
#         -a---              2009/07/14 12:51:00                    0.00  B                  setuperr.log                                                                                 
#         -a---              2013/06/03 12:01:12                  857.00  B                  pear.ini                                            
#         -a---              2009/06/11 04:36:48                    1.37 KB                  msdfmap.ini                                         
#         -a---              2014/05/16 03:27:49                    1.90 KB                  epplauncher.mif                                                                            
#         -a---              2009/07/14 09:39:10                   15.00 KB                  fveupdate.exe                                       
#         -a---              2009/07/14 09:39:12                   16.50 KB                  hh.exe                                                                                
#         -a---              2011/06/03 01:21:02                    2.74 MB                  explorer.exe                                        
#         -a---              2014/03/24 09:09:27                  522.39 MB                  MEMORY.DMP 

###########################################################################################################################################################################################

# Tip 20: Converting User Names to SIDs

function NametoSID($name, $domain = $env:USERDOMAIN)
{
    $objUser = New-Object System.Security.Principal.NTAccount($domain, $name)
    $strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
    $strSID.Value
}

NametoSID $env:username      # Output: S-1-5-21-2146773085-903363285-719344707-1385424


# Turning SIDs into Real Names

function SIDtoName($sid)
{
    $objSID = New-Object System.Security.Principal.SecurityIdentifier($sid)

    try
    {
        $objUser = $objSID.Translate([System.Security.Principal.NTAccount])             # to turn security identifiers (SIDs) into real names
        $objUser.Value
    }
    catch
    {
        $sid
    }
}

SIDtoName "S-1-5-21-2146773085-903363285-719344707-1385424"                              # Output: domain\v-sihe

# And, here is a show case for the function: to enumerate all profiles on your computer, you can read them from the Registry. 
 # However, all profiles are stored with SIDs only. Thanks to your new function, you can now display the real user names of everyone who has a profile on your machine:

function Get-Profile
{
    $key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    
    dir $key -Name | ForEach-Object {SIDtoName $_}                                       # display the real user names of everyone who has a profile on your machine
}

###########################################################################################################################################################################################

# Tip 21: Outputting Text Reports without Truncating

# If you want to capture PowerShell results in a text file, you can redirect the results or pipe them to Out-File. 
 # In any case, what you capture is the exact representation of what would have been displayed in your PowerShell console. 
 # So, depending on the amount of data, columns may be missing or truncated.

function Out-Report($path = "$home\silencetest", [switch]$open)
{
    $input | Format-Table -AutoSize | Out-File $path -Width 10000

    if($open)
    {
        Invoke-Item $path
    }
}

Get-Process | Out-Report -open

###########################################################################################################################################################################################

# Tip 22: Creating Excel Reports from PowerShell Data

function Out-ExcelReport($Path = "$env:temp\report$(Get-Date -format yyyyMMddHHmmss).csv", [switch]$open)
{
    $input | Export-Csv -NoTypeInformation -UseCulture -Encoding UTF8

    if($open)
    {
        Invoke-Item $Path
    }
}

Out-ExcelReport | Out-ExcelReport -open

###########################################################################################################################################################################################

# Tip 23: Finding Network Adapter Data Based On Connection Name

# Sometimes it would be nice to be able to access network adapter configuration based on the name of that adapter as it appears in your network and sharing center.
Get-WmiObject Win32_NetworkAdapter -Filter 'NetConnectionID like "%LAN%"' |  ForEach-Object { $_.GetRelated('Win32_NetworkAdapterConfiguration') }

# If you want to further narrow this down and cover only NICs that are currently connected to the network, extend the WMI filter
Get-WmiObject Win32_NetworkAdapter -Filter '(NetConnectionID like "%LAN%") and (NetConnectionStatus=2)' |  ForEach-Object { $_.GetRelated('Win32_NetworkAdapterConfiguration') }

###########################################################################################################################################################################################

# Tip 24: Turning Multi-Value WMI Properties into Text

Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled=true" | Select-Object -ExpandProperty IPAddress
# Output:
#        178.18.30.42
#        fe80::70cb:39dc:abc:250c

(Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled=true" | Select-Object -ExpandProperty IPAddress) -join ","
# Output:
#        178.18.30.42,fe80::70cb:39dc:abc:250c

###########################################################################################################################################################################################

# Tip 25: Finding IP and MAC address

# When you query network adapters with WMI, it is not easy to find the active network card. To find the network card(s) that are currently connected to the network, 
 # you can filter based on NetConnectionStatus which needs to be "2" for connected cards. Then you can take the MAC information
 # from the Win32_NetworkAdapter class and the IP address from the Win32_NetworkAdapterConfiguration class and combine both into one custom return object

Get-WmiObject Win32_NetworkAdapter -Filter "NetConnectionStatus=2"
# Output:
#         ServiceName      : b37od60c
#         MACAddress       : 18:A8:05:C8:2B:D4
#         AdapterType      : Ethernet 805.3
#         DeviceID         : 7
#         Name             : Broadcom NetXtreme Gigabit Ethernet
#         NetworkAddresses : 
#         Speed            : 100000



function Get-NetworkConfig
{
    Get-WmiObject Win32_NetworkAdapter -Filter "NetConnectionStatus=2" | ForEach-Object {
    
        $result = 1 | Select-Object Name, IP, MAC
        $result.Name = $_.Name
        $result.MAc = $_.MacAddress

        $config = $_.GetRelated("Win32_NetworkAdapterConfiguration")
        $result.IP = $config | Select-Object -ExpandProperty IPAddress

        $result
    }
}

# Output:
#        Name                                            IP                                                   MAC                                                                   
#        ----                                            --                                                   ---                                                                   
#        Broadcom NetXtreme Gigabit Ethernet             {175.18.52.42, fe50::70cb:32dc:abc:223c}             16:A9:08:B8:3B:D6

###########################################################################################################################################################################################

# Tip 26: Print All PDF Files in Folders

dir C:\myfolder\*.pdf | ForEach-Object {Start-Process -FilePath $_.fullname -Verb print}        # to print out all PDF documents you have stored in one folder

###########################################################################################################################################################################################

# Tip 27: Determining Service Start Modes

Get-WmiObject Win32_Service | Select-Object Name, StartMode                # By using WMI, you can enumerate the start mode that you want your services to use

([wmi]"Win32_Service.Name='Spooler'").StartMode                            # to find out the start mode of one specific service

# Change Service Startmode
([wmi]'Win32_Service.Name="Spooler"').ChangeStartMode('Automatic').ReturnValue
([wmi]'Win32_Service.Name="Spooler"').ChangeStartMode('Manual').ReturnValue 

Get-Service spooler | Set-Service -StartupType Automatic

###########################################################################################################################################################################################

# Tip 28: Asking for Credentials

# When you write functions that accept credentials as parameters, add a transformation attribute! 
 # This way, the user can either submit a credential object (for example, a credential supplied from Get-Credential), or simply a user name as string. 
 # The transformation attribute will then convert this string automagically into a credential and ask for a password.

function Do-Something([System.Management.Automation.Credential()]$Credential)
{
    '$Credential now holds a valid credential object'
    $Credential
}

Do-Something                                                               # When you run do-something without parameters, it will automatically invoke the credentials dialog

###########################################################################################################################################################################################

# Tip 29: Enumerating Registry Keys

Dir HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | Select-Object -expand PSPath                                       # To enumerate all subkeys in a Registry key
# Output: Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AddressBook

Dir HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall -Name
# Output: AddressBook

Resolve-Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*
# Output: HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AddressBook

Resolve-Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object -ExpandProperty ProviderPath
# Output: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AddressBook

###########################################################################################################################################################################################

# Tip 30: Creating Multiline Strings

"Hello" * 12                              # Output: HelloHelloHelloHelloHelloHelloHelloHelloHelloHelloHelloHello 
, "Hello" * 12                            # Return an array with 12 elements

# The comma puts the string into an array, and when you multiply arrays, you get additional array elements. Simply use Out-String to turn that into a single multi-line string
$text = , "Hello" * 12
$text.GetType().FullName                  # Output: System.Object[]
$text.Count                               # Output: 12

$text = , "Hello" * 12 | Out-String       # use Out-String to turn that into a single multi-line string
$text.GetType().FullName                  # Output: System.String

$text
# Output:
#        Hello
#        Hello
#        Hello
#        (....)

###########################################################################################################################################################################################

# Tip 31: Adding New Lines to Strings

$text = "Hello"
$text += "World"
$text                                      # Output: HelloWorld


$text = @()
$text += "Hello"
$text += "World"
$text                                      # return array
# Output:
#        Hello
#        World

$text | Out-String                         # return string
# Output:
#        Hello
#        World

# So, to construct multiline text throughout your script, start with @() to create an empty array, then add all the lines to this array using +=. 
 # When you are done, pipe the array to Out-String to get one multiline string. 

###########################################################################################################################################################################################

# Tip 32: Best Practice for PowerShell Functions

<#

This is a best-practice message: when you create your own function, here are some things you should consider:


- Function name: use cmdlet naming syntax (Verb-Noun), and for verbs, stick to the list of approved verbs. For the noun part, 
use a meaningful English term, and use singular, not plural. So, don't call a function 'ListNetworkCards' but rather 'Get-NetworkCard'


- Company Prefix: To avoid name collisions, all public functions should use your very own noun-prefix. So don't call your 
function "Get-NetworkCard" because this generic name might be used elsewhere, too. Instead, pick a prefix for your company. 
If you work for, let's say, 'Global International', a prefix could be 'GI', and your function name would be "Get-GINetworkCard".


- Standard Parameter Names: Stick to meaningful standard parameter names. Don't call a parameter -PC. Instead, call it -ComputerName. 
Don't call it -File. Instead, call it -Path. While there is no official list of approved parameter names, 
you should get familiar with the parameter names used by the built-in cmdlets to get a feeling for it.

#>
Get-Command -CommandType Cmdlet | Where-Object {$_.Parameters} |Select-Object -ExpandProperty Parameters | 
    ForEach-Object { $_.Keys } | Group-Object -NoElement | Sort-Object Count, Name -Descending             # To get a feeling for what the parameter names are that built-in cmdlets use

###########################################################################################################################################################################################

# Tip 33: Finding Driver Information

# driverquery.exe returns all kinds of information about installed drivers, but the information seems a bit useless at first:
driverquery.exe /v                          # shows info in a not friendly way


# This console application does support a parameter called /FO CSV. This formats the information as a comma-separated list:
driverquery.exe /v /FO CSV


# Now here's the scoop: Powershell can not only read the results from console application. 
 # When the output is CSV, it can even convert the raw text automagically into objects. So, with just one line, you get tremendous information about drivers 
driverquery.exe /v /FO CSV | ConvertFrom-Csv | Select-Object "Display Name", "Start Mode", "Paged Pool(bytes)", Path
driverquery.exe /v /FO CSV | ConvertFrom-Csv | Select-Object "Display Name", "Start Mode", "Paged Pool(bytes)", Path | Out-GridView

$col = @{Name = "File Name"; Expression = {Split-Path $_.Path -Leaf}}              # this one can show path in a much more suitable way with being truncated
driverquery.exe /v /FO CSV | ConvertFrom-Csv | Select-Object "Display Name", "Start Mode", "Paged Pool(bytes)", $col | Sort-Object "Display Name"



# Home-Made Driver Query Tool
function Show-DriverDialog($computerName = $env:COMPUTERNAME)
{
    driverquery.exe /S $computerName /FO CSV | ConvertFrom-Csv | Out-GridView -Title "Driver on \\$computerName"
}

Show-DriverDialog


# Creating Your Own Get-Driver Tool
function Get-Driver($keyword = "*")
{
    $col = @{Name = "File Name"; Expression = {Split-Path $_.Path -Leaf}} 

    driverquery.exe /v /FO CSV | ConvertFrom-Csv | Select-Object "Display Name", "Start Mode", "Paged Pool(bytes)", $col | Where-Object {
    
        ($_ | Out-String -Width 300) -like "*$keyword*"
    }
}

Get-Driver -keyword "Disk Driver"
# Output:
#         Display Name                    Start Mode               Paged Pool(bytes)                File Name                                           
#         ------------                    ----------               -----------------                ---------                                           
#         Disk Driver                     Boot                     36,864                           disk.sys                                            
#         Floppy Disk Driver              Manual                   16,384                           flpydisk.sys   

###########################################################################################################################################################################################

# Tip 34: "More" Can Be Dangerous - Use Better Alternative

Get-EventLog -LogName System | more
# "more" can be dangerous as you see here. You will not get any results for a long time, 
 # and your CPU load increases. more.com first collects all results before it starts paginating it. This takes a long time and a lot of resources.


Get-EventLog -LogName System | Out-Host -Paging
# You immediately see the benefit: results appear momentarily, and no additional CPU load is created



# Creating a "Better" More
function more 
{
    param(
        [Parameter(ValueFromPipeline=$true)]
        [System.Management.Automation.PSObject]
        $InputObject
    )

 begin
 {
    $type = [System.Management.Automation.CommandTypes]::Cmdlet
    $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Out-Host', $type)
    $scriptCmd = {& $wrappedCmd @PSBoundParameters -Paging }
    $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
    $steppablePipeline.Begin($PSCmdlet)
 }

 process 
 { 
    $steppablePipeline.Process($_) 
 }

 end 
 {
    $steppablePipeline.End() 
 }

 #.ForwardHelpTargetName Out-Host
 #.ForwardHelpCategory Cmdlet
}

# Once you run it, whenever you use "more", behind the scenes PowerShell will now call "Out-Host -Paging" instead. 
 # That's why now, with the new "more" in place, you can safely use lines like this:
Get-EventLog -LogName System | more

###########################################################################################################################################################################################

# Tip 35: Opening MsgBoxes

$msg = New-Object -ComObject WScript.Shell
$msg.Popup("Hello", 5, "Title", 48)                 # creates a MsgBox for 5 seconds

# Note: If the user does not make a choice within that time, it returns -1: a perfect solution for scripts that need to run unattended if no one is around

# [Reference site:] http://msdn.microsoft.com/en-us/library/x83z1d9f(v=VS.84).aspx

###########################################################################################################################################################################################

# Tip 36: About Shares Remotely

# One: Creating Shares Remotely

# Let's assume you need to access another machine's file system but there is no network share available. 
 # Provided you have local administrator privileges and WMI remoting is allowed in your Firewall, here is a one-liner that adds another share remotely:

([wmiclass]"\\IIS-CTI5052\root\cimv2:Win32_Share").Create("C:\", "Hidden", 0, 12, "secret share").ReturnValue     # Create on remote machine

([wmiclass]"root\cimv2:Win32_Share").Create("C:\Users\v-sihe\silencetest", "Hidden", 0, 12, "secret share").ReturnValue                   # Create on local machine

# Note that this method does not allow for separate authentication, so your current user must have local administrator privileges on the target machine.
 # A return value of 0 indicates success. If you receive a "2", you do not have proper permissions. 

# [Reference site for Create()] http://msdn.microsoft.com/zh-cn/library/aa389393(v=vs.85).aspx



# Two: Removing Shares (Remotely, Too)

([wmi]"IIS-CTI5052\root\cimv2:Win32_Share='Hidden'").Delete()        # Removes a share locally or remote

([wmi]"root\cimv2:Win32_Share='Hidden'").Delete()                    # Removes a share locally

###########################################################################################################################################################################################

# Tip 37: Spying on Parameters

# Your own PowerShell functions can have the same sophisticated parameters, parameter types and parameter sets that you know from cmdlets. 
 # However, it is not always obvious how to construct the param() block appropriately. 
 # A clever way is to spy on cmdlets and look how they did it.

$cmd = Get-Command -CommandType Cmdlet Get-Event

[System.Management.Automation.ProxyCommand]::GetParamBlock($cmd)

# Output:
#        [Parameter(ParameterSetName='BySource', Position=0, ValueFromPipelineByPropertyName=$true)]
#        [string]
#        ${SourceIdentifier},
#    
#        [Parameter(ParameterSetName='ById', Mandatory=$true, Position=0, ValueFromPipelineByPropertyName=$true)]
#        [Alias('Id')]
#        [int]
#        ${EventIdentifier}

# Not only will you discover that the keyword Mandatory=$true made -LogName mandatory. You also see all the hidden parameter aliases as well as validator attributes. 

###########################################################################################################################################################################################

# Tip 38: About Clipboard

# Sending Text to the Clipboard

Get-Process | clip.exe                                 # send results directly to the clipboard
                                                                                                             
Set-Alias Out-Clipboard clip.exe                       # To make this line a bit more PowerShellish, add an alias
Get-Process | out-Clipboard




# Sending Text to Clipboard Everywhere if you didn't have clip.exe

function Out-Clipboard($text)                          # It solely uses .NET Framework functionality that is available in all versions and modes of PowerShell
{
    Add-Type -AssemblyName System.Windows.Forms

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true

    if($input -ne $null)                               # 如果你在脚本中使用管道，脚本收集上一个语句的执行结果，默认保存在$input自动变量中
    {
        $input.Reset()
        $tb.Text = $input | Out-String
    }
    else
    {
        $tb.Text = $text
    }

    $tb.SelectAll()
    $tb.Copy()
}

"Hello Silence." | Out-Clipboard
Get-Process | Out-Clipboard




# Reading the Clipboard

function Get-Clipboard
{
    Add-Type -AssemblyName System.Windows.Forms
    
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true
    $tb.Paste()
    $tb.Text
}

Get-Clipboard

###########################################################################################################################################################################################

# Tip 39: Bulk-Creating PDF Files from Word

function Export-WordToPDF                                                   # To convert a MS Word documents to PDF
{
    param(
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [Alias("FullName")]
    $path, 
    $pdfpath = $null
    )
    
    process 
    {
        if (!$pdfpath) 
        {
            $pdfpath = [System.IO.Path]::ChangeExtension($path, '.pdf')
        }
    
        $word = New-Object -ComObject Word.Application
        $word.displayAlerts = $false
        
        $word.Visible = $true
        $doc = $word.Documents.Open($path)
        #$doc.TrackRevisions = $false
        $null = $word.ActiveDocument.ExportAsFixedFormat($pdfpath, 17, $false, 1)
        
        $word.ActiveDocument.Close()
        $word.Quit()
    }
} 

Export-WordToPDF -path 'C:\Users\v-sihe\silencetest\New Microsoft Word Document.docx'

Dir c:\folder -Filter *.doc | Export-WordToPDF

###########################################################################################################################################################################################

# Tip 40: Retrieve Exchange Rates

$url = 'http://www.ecb.int/rss/fxref-usd.html'

$xml = New-Object XML
$xml.Load($url)

$xml.RDF.Item | ForEach-Object {

    $rv = 1 | Select-Object Date, Currency, Rate, Description

    $rv.Date = [DateTime]$_.Date
    $rv.Description = $_.Description."#text"
    $rv.Currency = $_.Statistics.ExchangeRate.TargetCurrency
    $rv.rate = $_.Statistics.ExchangeRate.Value.'#text'
    
    $rv
}

# Ouput:
#        Date                         Currency           Rate              Description                                         
#        ----                         --------           ----              -----------                                         
#        2014/09/17 21:15:00          USD                1.2956            1 EUR buys 1.2956 US dollar (USD) - The reference...
#        2014/09/16 21:15:00          USD                1.2949            1 EUR buys 1.2949 US dollar (USD) - The reference...
#        2014/09/15 21:15:00          USD                1.2911            1 EUR buys 1.2911 US dollar (USD) - The reference...
#        2014/09/12 21:15:00          USD                1.2931            1 EUR buys 1.2931 US dollar (USD) - The reference...
#        2014/09/11 21:15:00          USD                1.2928            1 EUR buys 1.2928 US dollar (USD) - The reference...

###########################################################################################################################################################################################

# Tip 41: Creating System Footprints

# WMI can retrieve a lot more object instances than you might think. If you submit a parent class, Get-WmiObject returns all instances of all derived classes.

Get-WmiObject CIM_PhysicalElement | Group-Object -Property __Class
# Output:
#         Count Name                      Group                                                                                                                                                                                
#         ----- ----                      -----                                                                                                                                                                                
#             2 Win32_PhysicalMemoryArray {\\IIS-V-SIHE-01\root\cimv2:Win32_PhysicalMemoryArray.Tag="Physical Memory Array 0", \\IIS-V-SIHE-01...
#             1 Win32_BaseBoard           {\\IIS-V-SIHE-01\root\cimv2:Win32_BaseBoard.Tag="Base Board"}                                       ...
#             4 Win32_SystemSlot          {\\IIS-V-SIHE-01\root\cimv2:Win32_SystemSlot.Tag="System Slot 0", \\IIS-V-SIHE-01\root\cimv2:Win32_S...
#            36 Win32_PortConnector       {\\IIS-V-SIHE-01\root\cimv2:Win32_PortConnector.Tag="Port Connector 0", \\IIS-V-SIHE-01\root\cimv2:W...
#             1 Win32_SystemEnclosure     {\\IIS-V-SIHE-01\root\cimv2:Win32_SystemEnclosure.Tag="System Enclosure 0"}                         ...
#             3 Win32_PhysicalMemory      {\\IIS-V-SIHE-01\root\cimv2:Win32_PhysicalMemory.Tag="Physical Memory 0", \\IIS-V-SIHE-01\root\cimv2...
#             2 Win32_PhysicalMedia       {\\IIS-V-SIHE-01\root\cimv2:Win32_PhysicalMedia.Tag="\\\\.\\PHYSICALDRIVE0", \\IIS-V-SIHE-01\root\ci...




$col =  @{Name = "Key"; Expression = {($_.__Class.ToString() -split "_")[-1]}}
$info = Get-WmiObject CIM_PhysicalElement | Select-Object *, $col | Group-Object -Property Key -AsHashTable -AsString
# Output:
#        Name                  Value                                                                                                                                                                       
#        ----                  -----                                                                                                                                                                       
#        PortConnector         {@{PSComputerName=IIS-V-SIHE-01; Status=; Name=Port Connector; ExternalReferenceDesignator=; __GENU...
#        PhysicalMedia         {@{PSComputerName=IIS-V-SIHE-01; __GENUS=2; __CLASS=Win32_PhysicalMedia; __SUPERCLASS=CIM_PhysicalM...
#        BaseBoard             {@{PSComputerName=IIS-V-SIHE-01; Status=OK; Name=Base Board; PoweredOn=True; __GENUS=2; __CLASS=Win...
#        SystemSlot            {@{PSComputerName=IIS-V-SIHE-01; Status=OK; SlotDesignation=PCI1; __GENUS=2; __CLASS=Win32_SystemSl...
#        SystemEnclosure       {@{PSComputerName=IIS-V-SIHE-01; Tag=System Enclosure 0; Status=; Name=System Enclosure; SecuritySt...
#        PhysicalMemoryArray   {@{PSComputerName=IIS-V-SIHE-01; Status=; Name=Physical Memory Array; Replaceable=; Location=3; __G...
#        PhysicalMemory        {@{PSComputerName=IIS-V-SIHE-01; __GENUS=2; __CLASS=Win32_PhysicalMemory; __SUPERCLASS=CIM_Physical...

$info.PortConnector
$info.PhysicalMedia
$info.BaseBoard



# HTML Reporting: Create a System Report
function Get-SystemReport($ComputerName = $env:ComputerName) 
{
    $htmlStart = "
    <HTML><HEAD><TITLE>Server Report</TITLE>
    <style>  
    body { background-color:#EEEEEE; }  
    body,table,td,th { font-family:Tahoma; color:Black; Font-Size:10pt }  
    th { font-weight:bold; background-color:#AAAAAA; }  
    td { background-color:white; }  
    </style></HEAD><BODY>
    <h2>Report listing for System $Computername</h2>
    <p>Generated $(get-date -Format 'yyyy-MM-dd hh:mm') </p>
    "
       
    $htmlEnd = '</body></html>'
    
    $htmlStart
     
    Get-WmiObject -Class CIM_PhysicalElement | Group-Object -Property __Class | ForEach-Object { 
       
        $_.Group |Select-Object -Property * | ConvertTo-HTML -Fragment -PreContent ('<h3>{0}</h3>' -f $_.Name ) 
    }

    $htmlEnd 
}

$path = "$home\silencetest\test.hta"
Get-SystemReport | Out-File -FilePath $path
Invoke-Item $path

###########################################################################################################################################################################################

# Tip 42: Removing Characters at the Beginning of Text

# To remove text at the beginning of a sentence rather than somewhere inside the sentence, use the operator -replace and the anchor '^'. 

'PC101 is the PC we are overhauling' -replace 'PC', ''            # Output: 101 is the  we are overhauling

'PC101 is the PC we are overhauling' -replace '^PC', ''           # Output: 101 is the PC we are overhauling

###########################################################################################################################################################################################

# Tip 43: Ignoring Empty Lines

# To read in a text file and skip blank lines:

$file = "c:\sometextfile.txt"
Get-Content $file | Where-Object {$_.Trim() -ne ""}               # It will omit empty lines, lines with only blanks and lines with only tabs

###########################################################################################################################################################################################

# Tip 44: Managing Internet Cookies

dir ([System.Environment]::GetFolderPath("Cookies"))                                                                        # Get all cookies

dir ([system.environment]::GetFolderPath('Cookies')) | Where-Object { $_.Name -like '*microsoft*' }                         # To list all cookies which contains microsoft

dir ([system.environment]::GetFolderPath('Cookies')) | Where-Object { $_.Name -like '*microsoft*' } | Get-Content           # Read the cookie content

dir ([system.environment]::GetFolderPath('Cookies')) | Where-Object { $_.Name -like '*microsoft*' } | Remove-Item -WhatIf   # Remove the related cookies
 
###########################################################################################################################################################################################

# Tip 44: Read/Delete/Move Every X. File

dir $env:windir | ForEach-Object {$x = 0} {$x++; if($x % 5 -eq 0) {$_}}                                          # list every 5th file in the Windows folder 

Get-Content $env:windir\windowsupdate.log | ForEach-Object { $x=0 } { $x++; if($x % 50 -eq 0) { $_ } }           # read every 50th line from windowsupdate.log

###########################################################################################################################################################################################

# Tip 45: Use WMI and WQL

Get-WmiObject -List *                                  # list all classes

Get-WmiObject -List Win32_*network*                    # list classes that deal with network

Get-WmiObject Win32_NetworkAdapterConfiguration        # pick one of the classes and enumerate its instances


# With WQL, a SQL-type query language for WMI, you can even create more sophisticated queries
Get-WmiObject -Query 'select * from Win32_NetworkAdapterConfiguration where IPEnabled = true'

# [Reference site:]http://msdn.microsoft.com/en-us/library/windows/desktop/aa394606(v=vs.85).aspx

###########################################################################################################################################################################################

# Tip 46: Check Active Internet Connection

# If your machine is connected to the Internet more than once, let's say cabled and wireless at the same time, which connection is used?
function Test-IPMetric
{
    Get-WmiObject Win32_NetworkAdapter -Filter "AdapterType = 'Ethernet 802.3'" | ForEach-Object {$_.GetRelated("Win32_NetworkAdapterConfiguration")} | 
        Select-Object Description, Index, IPEnabled, IPConnectionMetric
}

Test-IPMetric
# Output:
#         Description                                Index          IPEnabled         IPConnectionMetric
#         -----------                                -----          ---------         ------------------
#         Broadcom NetXtreme Gigabit Ethernet            7               True                         2

###########################################################################################################################################################################################

# Tip 47: Sending Emails with Special Characters

# PowerShell has built-in support for sending emails: Send-MailMessage! All you need is an SMTP server. 
 # However, with standard encoding you may run into issues where special characters are mangled. Use the -Encoding parameter and specify UTF8 to preserve such characters.

Send-MailMessage -Body 'I can send special characters: äöüß' -From hhbstar@hotmail.com -to hhbstar@hotmail.com `
    -Credential (Get-Credential hhbstar@hotmail.com) -SmtpServer smtp.live.com -Subject 'Sending Mail From PowerShell' -Encoding ([System.Text.Encoding]::UTF8)
# Exception here on powershell v3: Send-MailMessage : The SMTP server requires a secure connection or the client was not authenticated.

# when you receive above error message, make sure you add the switch parameter -UseSsl. 
 # This only works right though if you use PowerShell v3 (the public CTP2 is readily available). In PowerShell v2, Send-MailMessage does not use the correct port for SSL connections



# Solution:
Send-MailMessage -Body 'My mail message can contain special characters: äöüß' -From hhbstar@hotmail.com -to hhbstar@hotmail.com `
    -Credential hhbstar@hotmail.com -Port 587 -SmtpServer smtp.live.com -Subject 'Sending Mail from PowerShell' -Encoding UTF8 -UseSsl

# Note that PowerShell v3 Send-MailMessage accepts "UTF8" directly so you no longer need to submit the awkward ([System.Text.Encoding]::UTF8).


# [Reference site:] http://www.e-eeasy.com/SMTPServerList.aspx

###########################################################################################################################################################################################

# Tip 48: Analyzing System Restarts

# Get the reason why the remote machine restarted
Get-EventLog -LogName System -ComputerName localhost | Where-Object {$_.EventID -eq 1074} | ForEach-Object {

    $rv = New-Object PSObject | Select-Object Date, User, Action, process, Reason, ReasonCode, Comment, Message

    if ($_.ReplacementStrings[4])
    {
        $rv.Date = $_.TimeGenerated
        $rv.User = $_.ReplacementStrings[6]
        $rv.Process = $_.ReplacementStrings[0]
        $rv.Action = $_.ReplacementStrings[4]
        $rv.Reason = $_.ReplacementStrings[2]
        $rv.ReasonCode = $_.ReplacementStrings[3]
        $rv.Comment = $_.ReplacementStrings[5]
        $rv.Message = $_.Message

        $rv
    }

} | Select-Object Date, Action, Reason, User

# Output:
#        Date                        Action             Reason                                         User                                              
#        ----                        ------             ------                                         ----                                              
#        2014/09/18 11:33:26         restart            No title for this reason could be found        domain\v-sihe                                    
#        2014/09/18 11:32:36         restart            Other (Unplanned)                              domain\v-sihe                                    
#        2014/09/12 18:02:06         power off          No title for this reason could be found        domain\v-sihe                                    
#        2014/09/12 18:01:25         power off          Other (Unplanned)                              domain\v-sihe                                    
#        2014/09/12 03:34:08         restart            Operating System: Recovery (Planned)           NT AUTHORITY\SYSTEM 



<#    
    Event ID 1074 represents a restart event. Rather than extracting the relevant information from the event message text, 
    this code uses the ReplacementStrings property which is an array and holds the significant information bits.
    Accessing the event entries' replacement strings is much easier than parsing the message text.
    
    The code returns information only if the particular event entry has content in ReplacementStrings[4] (the 5th element of the array) 
    because only then does the event entry represent a shutdown or reboot event.
    
    Note that Get-EventLog supports the -ComputerName parameter, 
    so if a remote system is set up for remote access and you own the appropriate privileges, you can also analyze remote systems.
#>




# Analyzing System Restarts (Alternative)

# In PowerShell v2, a new cmdlet called Get-WinEvent was added. 
 # With it, you can not only access and read the "classic" event logs but also the application event logs introduced in Windows Vista.

Get-WinEvent -FilterHashtable @{logname = "System"; id = 1074} | ForEach-Object{

    $rv = New-Object PSObject | Select-Object Date, User, Action, process, Reason, ReasonCode, Comment
    $rv.Date = $_.TimeCreated
    $rv.User = $_.Properties[6].Value
    $rv.Process = $_.Properties[0].Value
    $rv.Action = $_.Properties[4].Value
    $rv.Reason = $_.Properties[2].Value
    $rv.ReasonCode = $_.Properties[3].Value
    $rv.Comment = $_.Properties[5].Value

    $rv

} | Select-Object Date, Action, Reason, User                                   # Note that Get-WinEvent will not work with Windows XP.

###########################################################################################################################################################################################

# Tip 49: Using Shared Variables

<#
    By default, all variables created in functions are local, so they only exist within that function and all functions that are called from within this function.
    
    Sometimes, you'd like to examine variables defined in a function after that function executed. Or you'd like to persist a variable, 
    so next time the function is called it can increment some counter or continue to work with that variable rather than creating a brand new variable each time.
    
    To achieve this, use shared variables by prepending 'Script:' to a variable name.
#>

function Call-Me 
{
    $Script:counter++
    
    "You called me $script:counter times!" 
}

Call-Me                                                               # Output: You called me 1 times!
Call-Me                                                               # Output: You called me 2 times!
Call-Me                                                               # Output: You called me 3 times!
                                                                      
$counter                                                              # Output: 3

# Note that the variable $counter now does not exist inside the function anymore. Instead, it is created in the context of the caller (the place where you called Call-Me). 
 # If you prepended a variable name with 'global:', the variable would be created in the topmost context and thus available everywhere. 

###########################################################################################################################################################################################

# Tip 50: About Regex Expression



# One: Remove Options from Command String

$str = 'xcopy "C:\Some Folder" "C:\Some New Folder Name" /y /r /Q'

# Method One:

$str -replace '/.', $null                                             # To remove all options from a raw text command                                     
# Output: xcopy "C:\Some Folder" "C:\Some New Folder Name" 

# All options are removed. -replace was looking for a "/" and then any other character (".") and replacing every occurrence with nothing ($null)


# Method Two: Removing Options from Command String (Enhancement)

($str -replace '/\w*', $null).Trim()
# Output: xcopy "C:\Some Folder" "C:\Some New Folder Name"

# It is a more versatile approach. It will remove any character or word that starts with "/" and then remove all trailing spaces as well


#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Two: Removing Multiple White Spaces

$str = '[     Man,     it works!     ]'

$str -replace '\s+', " "                   # Use -replace operator and look for whitespaces ("\s") that occur one or more time ("+"), then replace them all with just one whitespace       
# Output: [ Man, it works! ]


#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Three: Converting TABs to Spaces

$str = "This is a`tTAB-delimited`tText" 
$str                                               # Output: This is a	TAB-delimited	Text

$str -replace '\t', " "                            # Output: This is a TAB-delimited Text

###########################################################################################################################################################################################

# Tip 51: Keeping Remote Programs Running

<#    
    When you use PowerShell Remoting (like the Enter-PSSession cmdlet) to connect to another machine and then start a program using Start-Process, 
    the program is automatically associated to your remote session. Once you leave and discard your remote session, 
    all programs you previously started with Start-Process will also be killed.
    
    While this is a reasonable cleanup strategy for most scenarios, you sometimes want programs to continue to run, 
    even after you discarded your remoting session. To do that, use WMI rather than Start-Process to run the program:
#>

(Invoke-WmiMethod Win32_Process Create calc.exe).ReturnValue -eq 0                                 # Start process on local machine, it's(UI) visable

(Invoke-WmiMethod Win32_Process Create calc.exe -ComputerName IIS-CTI5052).ReturnValue -eq 0       # Start process on remote machine, it's(UI) not visable

(Invoke-WmiMethod Win32_Process Create calc.exe -ComputerName IIS-CTI5052 -Credential Administrator).ReturnValue -eq 0

# Note that calc.exe will run but is not visible to anyone. So in real life, you'd use this technique to launch command line tools or applications that are designed to run unattended.

###########################################################################################################################################################################################

# Tip 52: Finding Email of Logged On User

# In an Active Directory environment, PowerShell can easily find the currently logged on user and retrieve AD information about that user, for example, his or her email address:

$searcher = [ADSISearcher]"(SAMAccountName=$env:username)"
$searcher.FindOne().Properties.mail                                          # Output: v-alias@microsoft.com

###########################################################################################################################################################################################

# Tip 53: The Two Faces of -match

# The -match operator can be extremely useful because it accepts regular expression patterns and extracts the information that matches the pattern

'PC678 had a problem' -match 'PC(\d{3})'                                     # Output: True

$Matches
# Output:
#         Name                           Value                                                                                                                                                                       
#         ----                           -----                                                                                                                                                                       
#         1                              678                                                                                                                                                                         
#         0                              PC678 

$Matches[0]                              # Output: PC678
$Matches[1]                              # Output: 678



Remove-Variable matches
$Matches

'PC678 had a problem', 'PC112 was ok', 'SERVER12 was ok', 'PC612 not checked' -match 'PC(\d{3})'
$Matches

# As it turns out, -match works differently when applied to a collection (a comma-separated list of multiple items). 
# Here, it filters out those items that match the pattern. That's why in the example above, the entry 'SERVER12 was ok' was filtered out. 
# When applied to collections, -match does not populate $matches. If $matches still does contain information, then it is a left-over from a previous call to -match.

###########################################################################################################################################################################################

# Tip 54: Stripping Decimals Without Rounding

18 / 5                                                                                            # Output: 3.6
                                                                                                  
[int](18 / 5)                                                                                     # Output: 4

# To strip off all decimals behind the decimal point, use Truncate() from the .NET Math library
[Math]::Truncate(18 / 5)                                                                          # Output: 3

# Likewise, to manually round, use Floor() or Ceiling().
[Math]::Floor(18 / 5)                                                                             # Output: 3
[Math]::Ceiling(18 / 5)                                                                           # Output: 4

<#

    使用floor函数。floor(x)返回的是小于或等于x的最大整数。如：
    floor(10.5) == 10    floor(-10.5) == -11
    
    
    使用ceil函数。ceil(x)返回的是大于x的最小整数。如：
    ceil(10.5) == 11    ceil(-10.5) ==-10

#>

###########################################################################################################################################################################################

# Tip 55: Check for 64-bit Environment

[System.Environment]::Is64BitOperatingSystem            # Output: True
[System.Environment]::Is64BitProcess                    # Output: True


function Get-Platform
{
    if([System.IntPtr]::Size -eq 4)
    {
        "X86"
    }
    else
    {
        "AMD64"
    }
}

Get-Platform                                            # Output: AMD64

###########################################################################################################################################################################################

# Tip 56: Shutdown, Restart, Logoff, Hibernation, Standby-Mode


# Stop
Stop-Computer -Confirm



# Restart
Restart-Computer -Confirm




# LogOFF

# this is no exist powershell commlet to logoff computer, but you can use cmd command instead.
function Logoff-Computer
{
    shutdown.exe /L
}

function Logoff-Computer($computerName)
{
    (Get-WmiObject Win32_OperatingSystem -ComputerName $computerName).Win32Shutdown(0)
}

<#
    We can use the Win32Shutdown method with the following Flags to perform additional actions.

    Logoff:    
                0            Log Off     
                4            Forced Log Off (0+4) 
    
    
    Shutdown:
                1            Shutdown    
                5            Forced Shutdown (1+4) 
    
    
    Reboot:   
                2            Reboot    
                6            Forced Reboot (2+4) 
    
    
    Poweroff:
                8            Power Off    
                12           Forced Power Off (8+4)
    
    
    [Reference site:]http://msdn.microsoft.com/en-us/library/windows/desktop/aa394058  
#>



# Hibernate
function Invoke-Hibernate
{
    shutdown.exe /H
}

function Invoke-Hibernate
{
    Add-Type -AssemblyName System.Windows.Forms

    [System.Windows.Forms.Application]::SetSuspendState(1,0,0) | Out-Null
}

# Remember though that the machine must allow hibernation which is typically turned off for servers. To turn hibernation on, use this:
powercfg.exe /hibernate on

<# 小知识：

        睡眠(Sleep)：切断除内存的其它设备的供电，数据都还在内存中。需要少量电池来维持内存供电，一旦断电，则内存中的数据丢失，下次开机就是重新启动。

　　      休眠(hibernate）：把内存数据写入到硬盘中，然后切断所有设备的供电。不再需要电源。

　　      混合休眠(Hybird Sleep)：把内存数据写到硬盘中，然后切断除内存外的所有设备的供电。避免了一旦断电则内存中的数据丢失的情况，下次开机则是从休眠中恢复过来。
#>




# Standby-Mode
function Invoke-StandbyMode                                              # To programmatically enter standby mode
{
    Add-Type -AssemblyName System.Windows.Forms

    [System.Windows.Forms.Application]::SetSuspendState(0,0,0) | Out-Null
}

###########################################################################################################################################################################################

# Tip 57: Locking Workstation

# Lock-Workstation <=> Windows + L

# Method One:

function Lock-WorkStation
{
    $signature = '

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool LockWorkStation();
        
    '

    $lockWorkStation = Add-Type -MemberDefinition $signature -Name Win32LockWorkStation -Namespace Win32Functions -PassThru
    $lockWorkStation::LockWorkStation() | Out-Null
}




# Method Two:

function Lock-WorkStation
{
    # rundll32.exe location: -> $env:windir\System32\rundll32.exe
    rundll32.exe  user32.dll, LockWorkStation    
}

###########################################################################################################################################################################################

# Tip 58: Create Group Policy Reports

# Windows Server 2008 R2 comes with the GroupPolicy PowerShell module. 
# You might have to install that feature first before you can use it – run these lines with full Administrator privileges

Import-Module ServerManager
Add-WindowsFeature GPMC
# Output:
#        Success Restart Needed Exit Code Feature Result
#        ------- -------------- --------- --------------
#        True    No             Success   {Group Policy Management}


# Once installed, the GroupPolicy module provides you with a lot of new cmdlets to manage group policy objects. 
# Here’s a function that creates a nice HTML report for one or all group policies available in your domain:

function Show-GPOReport($GPOName = $null, $fileName = "$home\silencetest\test.hta")
{
    Import-Module GroupPolicy

    if($GPOName -eq $null)
    {
        Get-GPO -All | Select-Object -ExpandProperty DisplayName
    }
    else
    {
        Get-GPOReport -Name $GPOName -ReportType Html | Out-File $fileName
    }

    Invoke-Item $fileName
}

(Get-GPO -All).DisplayName | % {Get-GPOReport -Name $_ -ReportType Html} | Out-File $home\silencetest\test.hta          # Get all related detail report

###########################################################################################################################################################################################

# Tip 59: Make PowerShell Speak!

# By adding a new system type called System.Speech, you can make PowerShell speak out loudly
Add-Type -AssemblyName System.Speech
$synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
$synthesizer.Speak("Hi Silence, love you so much!")


# This .NET approach can replace the old COM-based approach occasionally used:
$oldstuff = New-Object -ComObject SAPI.SpVoice
$oldstuff.Speak("Dear Silence, love you forever!") | Out-Null


# The .NET code can do much more, supports lexicons and provides more information.

$synthesizer.GetInstalledVoices() | ForEach-Object {$_.VoiceInfo}                    # list the voices installed on your system
# Output:
#         Gender                : Female
#         Age                   : Adult
#         Name                  : Microsoft Anna
#         Culture               : en-US
#         Id                    : MS-Anna-1033-20-DSK
#         Description           : Microsoft Anna - English (United States)
#         SupportedAudioFormats : {System.Speech.AudioFormat.SpeechAudioFormatInfo}
#         AdditionalInfo        : {[Age, Adult], [AudioFormats, 18], [Gender, Female], [Language, 409]...}

###########################################################################################################################################################################################

# Tip 60: Recording Audio Text Files

$path = "$home\silencetest\voice.wav"

$text = "Did you know that PowerShell can record audio messages?
 All you need is some text. You can then turn the text into spoken language, convert it to a WAV file and play it back or send it to someone"

Add-Type -AssemblyName System.Speech

$synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
$synthesizer.SetOutputToWaveFile($path)
$synthesizer.Speak($text)
$synthesizer.SetOutputToDefaultAudioDevice()

Invoke-Item $path                                                             # Play back the recorded file

###########################################################################################################################################################################################

# Tip 61: Adding PowerShell Goodies to Server 2008 R2

# Windows Server 2008 R2 comes with a PowerShell module called ServerManager which in turn allows you to add additional features to the server.

Import-Module ServerManager                         # load the ServerManager cmdlets

Add-WindowsFeature PowerShell-ISE                   # add the PowerShell ISE editor (which brings the Out-GridView cmdlet along):

Add-WindowsFeature RSAT-AD-PowerShell               # add the ActiveDirectory module that provides you with many new Active Directory cmdlets

Add-WindowsFeature GPMC                             # to add the GPMC feature which installs a GroupPolicy PowerShell module

Import-Module ActiveDirectory, GroupPolicy          # To view all the new cmdlets, first load the two new modules

Get-Command -Module ActiveDirectory, GroupPolicy    # create a list of all the new cmdlets


# When you’re logged on to a domain, you will also get a new AD provider and the new AD: PowerShell drive
dir AD:

###########################################################################################################################################################################################

# Tip 62: Safely Running PowerShell Scripts

# If you want to run a PowerShell script from outside PowerShell, for example from within a batch file, 
# you probably know that you need to prepend powershell.exe to the script path. But that is not enough. Always add these three parameters to launch your script safely:

powershell.exe -noprofile -executionpolicy bypass -file "pathtoscript.ps1"

# -noprofile makes sure that your script runs in a default PowerShell environment and does not load any profile scripts. 
# That does not only speed up script launch, it also prevents profile scripts from changing the environment. 
# After all, you don’t want anyone to change “dir” to “del” before your script runs.

###########################################################################################################################################################################################

# Tip 63: Correctly Returning Exit Codes

'
param($code = 99)
exit $code
' | Out-File "$home\silencetest\test.ps1"

cmd
powershell.exe -noprofile -file "C:\Users\v-sihe\silencetest\test.ps1" 1234
echo %ERRORLEVEL%                                                                                             # Output: 1234

# This returns 1234 (or whatever code you submitted to your script when you called it). 
# The reason why so many people have trouble with returning exit codes is that they miss the –file parameter

powershell.exe -noprofile "C:\Users\v-sihe\silencetest\test.ps1" 1234
echo %ERRORLEVEL%                                                                                             # Output: 1

# Without –file, powershell.exe returns either 0 or 1. That’s because now your code is interpreted as a command, 
# and when the command returns an exit code other than 0, PowerShell assumes it failed. 
# It then returns its own exit code, which is always 0 (command ran fine) or 1 (command failed).

###########################################################################################################################################################################################

# Tip 64: Listing Domains in Forest

function Get-Domain
{
    $root = [ADSI]"LDAP://RootDSE"

    try
    {
        $oForestConfig = $root.Get("configurationNamingContext")
    }
    catch
    {
        Write-Warning 'You are currently not logged on to a domain' 
        break  
    }

    $oSearchRoot = [ADSI]("LDAP://CN=Partitions," + $oForestConfig)
    $AdSearcher  = [ADSISearcher]"(&(objectcategory=crossref)(netbiosname=*))"
    $AdSearcher.SearchRoot = $oSearchRoot

    $AdSearcher.FindAll() | ForEach-Object {
    
        if($_.Path -match "LDAP://CN=(.*?),")        # uses regular expressions to only return the last CN part of the domain DN
        {
            $Matches[1]
        }
    }
}

Get-Domain
# Output:
#         AFRICA
#         CORP
#         EUROPE

###########################################################################################################################################################################################

# Tip 65: Adding More Fonts to PowerShell Console

$key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Console\TrueTypeFont'

Set-ItemProperty -Path $key -Name '0' -Value 'Lucida Console'
Set-ItemProperty -Path $key -Name '00' -Value 'Courier New'
Set-ItemProperty -Path $key -Name '000' -Value 'Consolas'
Set-ItemProperty -Path $key -Name '0000' -Value 'Lucida Sans Typewriter'
Set-ItemProperty -Path $key -Name '00000' -Value 'OCR A Extended'

# Run this with Administrator privileges to add additional fonts to the console. Every new console you open now – including cmd.exe – can now select from these fonts. 
# To select a font, right-click the icon in the console title bar and choose “Properties”, then click on the “Font” tab.

###########################################################################################################################################################################################

# Tip 66: Finding Disk Controller Errors

# analyze your system event log for disk controller errors
Get-EventLog -LogName System -InstanceId 3221487627 -ErrorAction SilentlyContinue | ForEach-Object {$_.ReplacementStrings[0]} | Group-Object -NoElement | Sort-Object Count -Descending

# Such errors can indicate disk failure but most often they result from USB sticks that you removed unexpectedly. 
# If you do not get any results, then there are no disk controller errors – good for you!

###########################################################################################################################################################################################

# Tip 67: Adding Support For –WhatIf and -Confirm


function Check-RiskMitigation                    # to support the risk mitigation parameters –WhatIf and –Confirm in your functions
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()
    
    "This always works."

    if($PSCmdlet.ShouldProcess("localhost", "Trying to do something real bad"))
    {
        "I am doing serious stuff here."
    }
    else
    {
        "I am just kidding..."
    }

    "This always works."
}

Check-RiskMitigation
# Output:
#         This always works.
#         I am doing serious stuff here.
#         This always works.

Check-RiskMitigation -WhatIf
# Output:
#         This always works.
#         What if: Performing operation "Trying to do something real bad" on Target "localhost".
#         I am just kidding...
#         This always works.

Check-RiskMitigation -Confirm
#         This always works.

#         Confirm
#         Are you sure you want to perform this action?
#         Performing operation "Trying to do something real bad" on Target "localhost".
#         [Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"): y

#         I am doing serious stuff here.
#         This always works.


Check-RiskMitigation -Confirm
#         This always works.

#         Confirm
#         Are you sure you want to perform this action?
#         Performing operation "Trying to do something real bad" on Target "localhost".
#         [Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"): n

#         I am just kidding...
#         This always works.

###########################################################################################################################################################################################

# Tip 68: Converting UNIX time

$key = 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion'

Get-ItemProperty $key | Select-Object -ExpandProperty InstallDate            # Output: 1349935294

function ConvertFrom-UnixTime
{
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [int32]
        $unixTime
    )

    begin
    {
        $startDate = Get-Date -Date '01/01/1970' 
    }

    process
    {
        $timeSpan = New-TimeSpan -Seconds $unixTime
        $startDate + $timeSpan
    }
}

Get-ItemProperty $key | Select-Object -ExpandProperty InstallDate | ConvertFrom-UnixTime                                        # Output: Thursday, October 11, 2012 06:01:34

Get-ItemProperty $key | Select-Object -ExpandProperty InstallDate | ConvertFrom-UnixTime | New-TimeSpan | Select-Object -ExpandProperty Days   # Output: windows 708 days old
[int](((get-date -u %s) - (Get-ItemProperty 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion' | Select-Object -ExpandProperty InstallDate)) / 86400)         # Output: 708

###########################################################################################################################################################################################

# Tip 69: Validate IP Addresses

$pattern = '^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'

do
{
    $ip = Read-Host "Enter IP"
    $ok = $ip -match $pattern

    if($ok -eq $false)
    {
        Write-Warning ("'{0}' is not an IP address." -f $ip)
        Write-Host -fore Red -back White 'TRY AGAIN!'
    }

}until($ok)

###########################################################################################################################################################################################

# Tip 70: Using Background Jobs to Speed Up Things

# PowerShell is single-threaded and can only do one thing at a time, but by using background jobs, 
# you can spawn multiple PowerShell instances and work simultaneously. Then, you can synchronize them to continue when they all are done

# starting different jobs (parallel processing)
$job1 = Start-Job { Dir $env:windir *.log -Recurse -ea 0 }
$job2 = Start-Job { Start-Sleep -Seconds 10 }
$job3 = Start-Job { Get-WmiObject Win32_Service }

# synchronizing all jobs, waiting for all to be done
Wait-Job $job1, $job2, $job3

# receiving all results
Receive-Job $job1, $job2, $job3

# cleanup
Remove-Job $job1, $job2, $job3

# This can speed up logon scripts and other tasks as long as the jobs you send to the background jobs are independent of each other – and take considerable processing time. 
# If your tasks only take a couple of seconds, then sending them to background jobs causes more delay than benefit.

###########################################################################################################################################################################################

# Tip 71: Pipeline Used Or Not

function Test
{
    [CmdletBinding(DefaultParameterSetName="NonPipeline")]
    param(
        [Parameter(ValueFromPipeline=$true)]
        $data
    )

    begin
    {
        $direct = $PSBoundParameters.ContainsKey("Data")
    }

    process{}

    end
    {
        $pipeline = $PSBoundParameters.ContainsKey("Data") -and -not $direct

        "Direct? $direct"
        "Pipeline? $pipeline"
    }
}

1..3 | Test
# Output:
#        Direct? False
#        Pipeline? True


Test 123
# Output:
#        Direct? True
#        Pipeline? False


Test
# Output:
#        Direct? False
#        Pipeline? False

###########################################################################################################################################################################################

# Tip 72: ASCII Table

32..255 | ForEach-Object {

    "{0} -> {1}" -f $_, [char]$_
}

-join [char[]](32..255)
# Output:
#          !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~#          ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéê
#         ëìíîïðñòóôõö÷øùúûüýþÿ


-join [char[]](65..90)            # Output: ABCDEFGHIJKLMNOPQRSTUVWXYZ

[char[]](65..90) -join ","        # Output: A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z

###########################################################################################################################################################################################

# Tip 73: Sending Variables through Environment Variables

# Let's assume you want to launch a new PowerShell session from an existing one. To carry over information from the initial session, you can use environment variables

$env:payload = 'Open a new powershell console and this is message for initial'

Start-Process powershell -ArgumentList '-noexit -command Write-Host "I received: $env:payload"'

Start-Process powershell -ArgumentList '-command Write-Host "I received: $env:payload"'          # Note: if missed -noexit here, the new console will exit once the command finished

###########################################################################################################################################################################################

# Tip 74: Converting to Hex

# decimal -> hex
(-21441564124).ToString("X")   # Output: FFFFFFFB01FBB224
(-21441564124).ToString("x")   # Output: fffffffb01fbb224


"{0:X}" -f -21441564124        # Output: FFFFFFFB01FBB224
"{0:x}" -f -21441564124        # Output: fffffffb01fbb224


# hex -> decimal
0x8004101                      # Output: 134234369 (Automatically)

###########################################################################################################################################################################################

# Tip 75: Counting Work Days

function Get-WeekDay
{
    param(
        $month = $(Get-Date -Format "MM"),
        $year  = $(Get-Date -Format "yyyy"),
        $days  = 1..5
    )

    $maxDays = [System.DateTime]::DaysInMonth($year, $month)                                      # days in month

    1..$maxDays | ForEach-Object {
    
        Get-Date -Day $_ -Month $month -Year $year | Where-Object {$days -contains $_.DayOfWeek}
    }
}

Get-WeekDay

Get-WeekDay | Measure-Object | Select-Object -ExpandProperty Count                                # working days (monday to friday) in month

Get-WeekDay -days "Saturday", "Sunday" | Measure-Object | Select-Object -ExpandProperty Count     # how many Saturday and Sunday in month?        Output: 8

Get-WeekDay -days "Monday"
# Output:
#         Monday, September 01, 2014 18:02:40
#         Monday, September 08, 2014 18:02:40
#         Monday, September 15, 2014 18:02:40
#         Monday, September 22, 2014 18:02:40
#         Monday, September 29, 2014 18:02:40

Get-Date -Day 19 -Month 10 -Year 2014 -Format "yyyy-MM-dd"                                        # Output: 2014-10-19

###########################################################################################################################################################################################

# Tip 76: Fun with Date and Time

# Get-Date can extract valuable information from dates. All you need to know is the placeholder for the date part you are after. 

Get-Date -Format "M"                        # Output: September 22
                                            
Get-Date -Format "MM"                       # Output: 09
                                            
Get-Date -Format "MMM"                      # Output: Sep
                                            
Get-Date -Format "MMMM"                     # Output: September

###########################################################################################################################################################################################

# Tip 77: Mapping Printers

# Method One: CMD command

rundll32 printui.dll,PrintUIEntry /in /n "\\pntsrv1\HP552"                                       # To map a network printer to a user profile [low level command]
# This will map the printer share and also install drivers if required. A dialog window outputs progress information (unless you also add the switch /q for "quiet operation").

rundll32 printui.dll,PrintUIEntry /?                                                             # To see all options and also example code

# Two pitfalls you should know: this command is case-sensitive, so "PrintUIEntry" must be spelled exactly like this. And mapping a printer affects the user profile. 
# So if you map a printer while running an elevated shell, the mapping will not be visible to your regular user shell (and vice versa).


# Method Two: COM object

(New-Object -ComObject Wscript.Network).AddWindowsPrinterConnection("\\pntsrv1\HP552")

###########################################################################################################################################################################################

# Tip 78: Getting Timezones

[System.TimeZoneInfo]::GetSystemTimeZones()                      # Returns all time zones:


[System.TimeZoneInfo]::Local
# Output:
#        Id                         : China Standard Time
#        DisplayName                : (UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi
#        StandardName               : China Standard Time
#        DaylightName               : China Daylight Time
#        BaseUtcOffset              : 08:00:00
#        SupportsDaylightSavingTime : False




function Get-TimeZone([string]$city)                               # to get the time zone for a specific city
{
    $timeZones = [System.TimeZoneInfo]::GetSystemTimeZones()

    foreach($timeZone in $timeZones)
    {
        if($city -eq $null -or ($timeZone -like ("*"+$city+"*")))
        {
            Write-Host $timeZone
        }
    }
}

Get-TimeZone Beijing                                               # Output: (UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi


# Convert the current time to different time zone
$now = Get-Date
[System.TimeZoneInfo]::GetSystemTimeZones() | ForEach-Object{

    Write-Host ("{0,-35} ==>> {1}" -f $_.Id,[System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($now, $_.Id))
}

###########################################################################################################################################################################################

# Tip 79: Converting to Signed

# If you convert a hex number to a decimal, the result may not be what you want:
0xFFFF                                                                              # Output: 65535

# PowerShell converts it to an unsigned number (unless its value is too large for an unsigned integer). If you need the signed number, 
# you would have to use the Bitconverter type and first make the hex number a byte array, then convert this back to a signed integer like this:

[BitConverter]::ToInt16([BitConverter]::GetBytes(0xFFFF), 0)                        # OUtput: -1



# Converting to Signed Using Casting
0xffff                                                                              # Output: 65535
0xfffe                                                                              # Output: 65534
                                                                                   
[int16]("0x{0:x4}" -f ([UInt32]0xffff))                                             # Output: -1
[int16]("0x{0:x4}" -f ([UInt32]0xfffe))                                             # Output: -2

###########################################################################################################################################################################################

# Tip 80: Shrinking Paths

# Many file-related .NET Framework methods fail when the overall path length exceeds a certain length. 
# Use low-level methods to convert lengthy paths to the old 8.3 notation which is a lot shorter in many cases:

function Get-ShortPath($Path) 
{
    $code = @'

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError=true)]
        public static extern uint GetShortPathName(string longPath,StringBuilder shortPath,uint bufferSize);
'@

    $API = Add-Type -MemberDefinition $code -Name Path -UsingNamespace System.Text -PassThru

    $shortBuffer = New-Object Text.StringBuilder ($Path.Length * 2)
    $rv = $API::GetShortPathName( $Path, $shortBuffer, $shortBuffer.Capacity )

    if ($rv -ne 0) 
    {
        $shortBuffer.ToString()
    } 
    else 
    {
        Write-Warning "Path '$path' not found."
    }
}


$longFileName = -join (0..200 | ForEach-Object {([char[]](97..122) + [char[]](65..90) | Get-Random -Count 1)})

$null = md $home\silencetest\$longFileName                          # windows下完全限定文件名必须少于260个字符，目录名必须小于248个字符
Test-Path "$home\silencetest\$longFileName"                         # Output: True


Get-ShortPath -Path "$home\silencetest\$longFileName"               # Output: C:\Users\v-sihe\SILENC~4\JFZWIR~1
Test-Path C:\Users\v-sihe\SILENC~4\JFZWIR~1                         # Output: True

###########################################################################################################################################################################################

# Tip 81: Counting Log Activity

# Did you know that Group-Object can analyze text-based log files for you? Here's sample code that tells you how many log entries on a given day a log file contains:

Get-Content $env:windir\windowsUpdate.log | Group-Object {$_.SubString(0,10)} -NoElement | Sort-Object Count -Descending | Select-Object Count,Name
# Output:
#           Count Name                                                                                                 
#           ----- ----                                                                                                 
#            2045 2014-09-10                                                                                           
#            1972 2014-09-12                                                                                           
#            1787 2014-09-18                                                                                           
#            1777 2014-09-09                                                                                           
#            1724 2014-09-15 

# The trick is to submit a script block to Group-Object that extracts the piece of information you want to use for grouping. 
# In the file windowsupdate.log, the first ten characters represent the date on which the line was written to the file.


# Counting Log Activity Based On Product Install

# Here's a refined snippet. It will count on which days your Windows box received the most updates:
Get-Content $env:windir\WindowsUpdate.log | Where-Object {$_ -like '*successfully installed*'} | 
    Group-Object {$_.SubString(0,10)} -NoElement | Sort-Object Count -Descending | Select-Object Count,Name
# Output:
#          Count Name                                                                                                 
#          ----- ----                                                                                                 
#              7 2014-09-10                                                                                           
#              5 2014-09-12                                                                                           
#              3 2014-09-15                                                                                           
#              3 2014-09-11                                                                                           
#              3 2014-09-05                                                                                           


###########################################################################################################################################################################################

# Tip 82: Checking if a Text Ends with Certain Characters

"somefile.pdf".ToLower().EndsWith(".pdf")                                  # Output: True     # Check whether the file name ends with '.pdf'

# You can always use the String method EndsWith(). Just make sure you convert the text to lower-case first to avoid case-sensitive comparison. This method does not accept wildcards.

###########################################################################################################################################################################################

# Tip 83: Checking Text Ending with Wildcards

# This uses a regular expression to check for text that ends with three ("{3}") digits ("\d") at the end of a line ("$").

"Account123" -match '\d{3}$'                                               # Output: True

dir $env:windir\System32 | Where-Object {$_.BaseName -match '\d{3}$'}      # listing all files in a folder that end with three digits

###########################################################################################################################################################################################

# Tip 84: Making Names Unique

# To make a list of items or names unique, you could use grouping and then, when a group has more than one item, append a numbered suffix to these items. 
function Convert-Unique($list)
{
    $list | Group-Object | ForEach-Object {
    
        $_.Group | ForEach-Object {$i = 0}{
        
            $i++

            if($i -gt 1)
            {
                $_ += $i
            }

            $_
        }    
    }
}

$names = "Peter", "Mary", "Peter", "Fred", "Tom", "Peter", "Mary"

Convert-Unique $names
# Output:
#         Peter
#         Peter2
#         Peter3
#         Mary
#         Mary2
#         Fred
#         Tom

# Of course this is just a very simple algorithm. It assumes that there are no names in the initial list that end with numbers.

###########################################################################################################################################################################################

# Tip 85: Verbose Driver Information

# driverquery.exe to list driver information. This tool sports a /V switch for even more verbose information. 
# However, due to localization errors, when you specify /V, column names may no longer be unique.

function Show-DriverDialogVerbose($ComputerName = $env:computername)
{   
    function Convert-Unique($list) 
    {
        $list | Group-Object | ForEach-Object {

            $_.Group | ForEach-Object { $i=0 }{
             
                $i++

                if ($i -gt 1) 
                {
                   $_ += $i
                }
                
                $_ 
             }
         }
     }

     driverquery.exe /V /S $ComputerName /FO CSV | ForEach-Object {$i=0} { 
    
        $i++

        if ($i -eq 1) 
        {
            """$((Convert-Unique ($_.Replace('"','').Split(',') )) -join '","')"""
        } 
        
        $_

     } | ConvertFrom-Csv | Out-GridView -Title "Driver on \\$ComputerName"
}

###########################################################################################################################################################################################

# Tip 86: Finding Numbers in Text

# Regular Expressions are a great help in identifying and extracting data from text. Here's an example that finds and extracts a number that ends with a comma:

$text = "I am looking for a number like this 67868683468932689223479, that is delimited by a comma."

$pattern = '(\d*),'

if($text -match $pattern)
{
    $Matches[1]
}
else
{
    Write-Warning "Not found."
}

# Output: 67868683468932689223479

###########################################################################################################################################################################################

# Tip 87: Converting Bitmaps to Icons

function ConvertTo-Icon 
{
    param(
      [Parameter(Mandatory=$true)]
      $bitmapPath,
      $iconPath = "$env:temp\newicon.ico"
    )
    
    Add-Type -AssemblyName System.Drawing
    
    if (Test-Path $bitmapPath) 
    {
        $b = [System.Drawing.Bitmap]::FromFile($bitmapPath)
        $icon = [System.Drawing.Icon]::FromHandle($b.GetHicon())

        $file = New-Object System.IO.FileStream($iconPath, 'OpenOrCreate') 
        $icon.Save($file)
        $file.Close()
        $icon.Dispose()
    
        explorer "/SELECT,$iconpath"               # Open an Explorer window and select a file inside of it
    } 
    else 
    { 
        Write-Warning "$BitmapPath does not exist" 
    }
}

ConvertTo-Icon -bitmapPath C:\Users\v-sihe\silencetest\Untitled.bmp   

###########################################################################################################################################################################################

# Tip 88: Use Internet Connection with Default Proxy

# To use the same proxy settings that are set in your Internet Explorer browser
$proxy = [System.Net.WebRequest]::GetSystemWebProxy()
$proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

$web = New-Object System.Net.WebClient
$web.Proxy = $proxy
$web.UseDefaultCredentials = $true

# Once set up, you can then use the $web object to download web site content or even files:

$url = "http://www.powershell.com"
$web.DownloadString($url)




# IE -> Internet Options -> Connections -> LAN settings -> Proxy server

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings" -name ProxyEnable -Value 1           # 设置代理生效
set-itemproperty -path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings" -name ProxyServer -Value $ip         # 设置代理ip

###########################################################################################################################################################################################

# Tip 89: Output Scheduled Tasks to XML

function Export-ScheduledTask([Parameter(Mandatory=$true)]$taskName, [Parameter(Mandatory=$true)]$xmlFileName)
{
    schtasks /QUERY /TN $taskName /XML | Out-File $xmlFileName
}

Export-ScheduledTask "Book lunch on ele.me" -xmlFileName $home\silencetest\booklunch.xml

# Specify the name of a scheduled task and a path to some XML file. If you are not sure what the names of your scheduled tasks are, 
# this is how you can list the names of all scheduled tasks you can access:

schtasks /QUERY                 # list the names of all scheduled tasks you can access
schtasks /QUERY /FO List

schtasks /QUERY /?

###########################################################################################################################################################################################

# Tip 90: Creating Scheduled Tasks From XML

function Import-ScheduledTask               # re-import that XML file to re-create the scheduled task
{
    param
    (
        [Parameter(Mandatory = $true)]
        $jobName,

        [Parameter(Mandatory = $true)]
        $path, 

        $computerName = $null
    )

    if($computerName -ne $null)
    {
        $iption = "/S $computerName"
    }
    else
    {
        $option = ""
    }

    schtasks /CREATE /TN $jobName /XML $path $option
}

Import-ScheduledTask -jobName testbooklauch -path $home\silencetest\booklunch.xml

# Note: the xml was exported from local machine with account v-sihe, it can be imported success on machine IIS-CTI5052 with the same account, but got error on machine IIS-CSB2021
# because this machine was login with administrator account. Error message: ERROR: No mapping between account names and security IDs was done.


# You can use this technique to clone a scheduled task to multiple machines, or you can first export a scheduled task to XML, 
# then adjust all the advanced settings inside the XML file, and finally reimport the scheduled task from the adjusted XML file.

###########################################################################################################################################################################################

# Tip 91: Killing Long-Running Scripts

# You can use a background thread to monitor how long a script is running, and then kill that script if it takes too long. 
# You can even write to an event log or send off a mail before the script is killed.


function Start-TimeBomb                              # This function can be used to exec a command in a limited time
{
    param
    (
        [int32]$seconds,
        [ScriptBlock]$Action = {Stop-Process -id $pid}
    )

    $wait = "Start-Sleep -Seconds $seconds"
    $Script:newPowershell = [PowerShell]::Create().AddScript($wait).AddScript($Action)
    $handle = $Script:newPowershell.BeginInvoke()

    Write-Warning "Timebomb is active and will go off in $Seconds seconds unless you call Stop-Timebomb before."
}


function Stop-TimeBomb
{
    if($Script:newPowerShell -ne $null)
    {
        Write-Host "Trying to stop timebomb ..." -NoNewline

        $Script:newPowerShell.Stop()
        $Script:newPowerShell.Runspace.Close()
        $Script:newPowerShell.Dispose()

        Remove-Variable newPowerShell -Scope Script

        Write-Host "Done!"
    }
    else
    {
        Write-Warning "No timebomb found."
    }
}

Start-Timebomb -Seconds 5                            # Kill the powershell.exe or powershell_ise.exe after 5 seconds if you don't call stop-timebomb function

# After 30 seconds, the script will be killed unless you call Stop-Timebomb in time. Start-Timebomb supports the parameter -Action that accepts a script block. 
# This is the code that gets executed after time is up, so here you could also write to event logs or simply output a benign warning:

Start-Timebomb -Seconds 5 -Action { [System.Console]::Write('Your breakfast egg is done, get it now!') }

# Note though that because the timebomb is running in another thread, you cannot output data to the foreground thread. 
# That's why the sample used the type System.Console to write directly into the console window.

###########################################################################################################################################################################################

# Tip 92: Adding Progress to Long-Running Cmdlets

# Sometimes cmdlets take some time, and unless they emit data, the user gets no feedback. 
# Here are three examples for calls that take a long time without providing user feedback:
$hotfix = Get-HotFix
$products = Get-WmiObject Win32_Product
$scripts = Get-ChildItem $env:windir *.ps1 -Recurse -ErrorAction SilentlyContinue


# To provide your scripts with better user feedback, here's a function called Start-Progress. It takes a command and then executes it in a background thread. 
# The foreground thread will output a simple progress indicator. Once the command completes, the results are returned to the foreground thread.

function Start-Process([ScriptBlock]$code)
{
    $newPowerShell = [PowerShell]::Create().AddScript($code)
    $handle = $newPowerShell.BeginInvoke()

    while($handle.IsCompleted -eq $false)
    {
        Write-Host "." -NoNewline
        Start-Sleep -Milliseconds 500
    }

    Write-Host

    $newPowerShell.EndInvoke($handle)
    $newPowerShell.Runspace.Close()
    $newPowerShell.Dispose()
}


$hotfix = Start-Process {Get-HotFix}
$products = Start-Process {Get-WmiObject Win32_Product}
$scripts = Start-Process {Get-ChildItem $env:windir *.ps1 -Recurse -ErrorAction SilentlyContinue}


# Much more friendly process message:
function Start-Progress 
{
    param
    (
      [parameter(mandatory=$true)]
      [ScriptBlock] $Code,
    
      [parameter(mandatory=$false)] 
      [string] $Statusmessage
    )
      
    $newPowerShell = [PowerShell]::Create().AddScript($Code)
    $handle = $newPowerShell.BeginInvoke() 
    
    if (-not [string]::IsNullOrEmpty($Statusmessage))
    { 
        Write-Host $Statusmessage -NoNewline -ForegroundColor Yellow
    }
    
    while ($handle.IsCompleted -eq $false) 
    {   
        Write-Host '.' -NoNewline -ForegroundColor Yellow  
        Start-Sleep -Milliseconds 500
    }
    
    Write-Host ' Done!' -ForegroundColor Green
    Write-Host ''  

    $newPowerShell.EndInvoke($handle)  
    $newPowerShell.Runspace.Close()
    $newPowerShell.Dispose()
}

$hotfix = Start-Progress {Get-Hotfix}
$hotfix = Start-Progress {Get-Hotfix} -Statusmessage "Inventoring hotfixes "

$products = Start-Progress {Get-WmiObject Win32_Product} -Statusmessage "Inventoring installed products "

$lookHere = $env:windir
$scripts = Start-Progress {Get-ChildItem $lookHere *.ps1 -Recurse -ea 0} -Statusmessage "Inventoring Powershellfiles found in '$lookHere' "


###########################################################################################################################################################################################

# Tip 93: Listing All WMI Namespaces

Get-WmiObject -Query "select * from __Namespace" -Namespace Root | Select-Object -ExpandProperty Name     # lists all namespaces you got
# Output:
#        subscription
#        DEFAULT
#        cimv2
#        Cli
#        Nap
#        SECURITY
#        SmsDm
#        CCMVDI
#        RSOP
#        WebAdministration
#        HP
#        ccm
#        WMI
#        directory
#        Policy

(Get-WmiObject -Namespace root\SecurityCenter2 -List).Name                  # Get all classes for namespace root\SecurityCenter2
Get-WmiObject -Namespace root\SecurityCenter2 -Class AntivirusProduct       # once you know the classes, you could retrieve information

###########################################################################################################################################################################################

# Tip 94: Executing Commands in Groups


# In traditional batch files, you can use "&&" to execute a second command only if the first one worked. 
# In PowerShell, the same can be achieved by using the try/catch construct. 


# Now, if you want to execute a group of command and abort everything once an error occurs, simply place the commands inside the try block.
# If the commands are native console commands, add a "2>&1" to each command.

try
{
    $ErrorActionPreference = "Stop"

    net user nonexistent 2>&1             # this raises an error 
           
    ipconfig 2>&1                         # this will not execute due to the previous error
}
catch{}


# Try and replace "nonexistent " with an existing local user account such as "Administrator", and you'll see that ipconfig will execute.

try
{
    $ErrorActionPreference = "Stop"

    net user administrator 2>&1           # this raises an error 
           
    ipconfig 2>&1                         # this will not execute due to the previous error
}
catch{}

###########################################################################################################################################################################################

# Tip 95: Finding Domain Controllers

# If your computer is logged on to an Active Directory, here is some code to get to your domain controllers. 
# Note that this will raise errors if you are currently not logged on to a domain.

$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$domain.DomainControllers                                                        # to get to your domain controllers

$domain.FindDomainController()                                                   # lists all domain controllers

$domain.Forest.Domains                                                           # to find all the domain controllers in your forest

###########################################################################################################################################################################################

# Tip 96: Creating Symmetric Array

# By default, PowerShell uses jagged arrays. To create conventional symmetric arrays, here's how:

$array = New-Object "Int32[,]" 2,2
$array
# Output:
#        0
#        0
#        0
#        0

# This creates a two-dimensional array of Int32 numbers. Note that each dimension starts with 0, so this array goes from $array[0,0] to $array[1,1].

$array[1,1] = 100
$array[1,0] = 100
$array
# Output:
#        0
#        0
#        100
#        100

###########################################################################################################################################################################################

# Tip 97: Get Localized Month Names

[System.Enum]::GetNames([System.DayOfWeek])                                             # To get a list of week names

# Output: (en-us)
#        Sunday
#        Monday
#        Tuesday
#        Wednesday
#        Thursday
#        Friday
#        Saturday

# Output: (fr-fa)
#        Sonntag
#        Montag
#        Dienstag
#        Mittwoch
#        Donnerstag
#        Freitag
#        Samstag

0..11 | ForEach-Object {[Globalization.DatetimeFormatInfo]::CurrentInfo.MonthNames[$_]}  # To get a list of month names

# Output: (en-us)
#        January
#        February
#        March
#        April
#        May
#        June
#        July
#        August
#        September
#        October
#        November
#        December

###########################################################################################################################################################################################

# Tip 98: Optimizing PowerShell Performance

# PowerShell is loading .NET assemblies. These assemblies can be precompiled using the tool ngen.exe 
# which improves loading times (because the DLLs no longer have to be compiled each time they are loaded).

# Before you think about optimizing the DLLs PowerShell uses, you should do some reading on ngen.exe and its benefits. 
# Then, you could use the following code to optimize all DLLs loaded by PowerShell. You do need Administrator privileges for this.


$frameWorkDir = [Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()   # Output: C:\Windows\Microsoft.NET\Framework64\v4.0.30319\
$ngenPath = Join-Path $frameWorkDir "ngen.exe"

[AppDomain]::CurrentDomain.GetAssemblies() | Select-Object -ExpandProperty Location | ForEach-Object {& $ngenPath """$_"""}

###########################################################################################################################################################################################

# Tip 99: Who am I?

$env:USERDOMAIN
$env:username

# You get a lot more information including your security identifier (SID) by using the appropriate .NET methods:

[System.Security.Principal.WindowsIdentity]::GetCurrent()

# Output:
#        AuthenticationType : Kerberos                                      
#        ImpersonationLevel : Impersonation
#        IsAuthenticated    : True
#        IsGuest            : False
#        IsSystem           : False
#        IsAnonymous        : False
#        Name               : domain\sihe
#        Owner              : S-1-5-32-544
#        User               : S-1-5-21-2146773085-903363285-719344707-138...
#        Groups             : {S-1-5-21-2146773085-903363285-719344707-51...
#        Token              : 2004                                       
#        UserClaims         : {http://schemas.xmlsoap.org/ws/2005/05/iden...
#                             S-1-5-21-2146773085-903363285-719344707-138... 
#                             http://schemas.microsoft.com/ws/2008/06/ide...
#        DeviceClaims       : {}                                         ...
#        Claims             : {http://schemas.xmlsoap.org/ws/2005/05/iden...
#                             S-1-5-21-2146773085-903363285-719344707-138... 
#                             http://schemas.microsoft.com/ws/2008/06/ide...
#        Actor              :                                            
#        BootstrapContext   :                                            
#        Label              :                                            
#        NameClaimType      : http://schemas.xmlsoap.org/ws/2005/05/ident...
#        RoleClaimType      : http://schemas.microsoft.com/ws/2008/06/ide...


# Getting Group Memberships

$User=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$User.Groups | ForEach-Object { $_.Translate([System.Security.Principal.NTAccount]).Value } | Sort-Object



# Integrating WhoAmI Into PowerShell

whoami /groups /fo csv | ConvertFrom-Csv

###########################################################################################################################################################################################

# Tip 100: Clearing WinEvent Logs

# With Get-WinEvent you can access the various Windows log files such as this one:
Get-WinEvent Microsoft-Windows-WinRM/Operational

# There is no cmdlet to actually clear such event log, though. With this line, you can:
[System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.ClearLog(' Microsoft-Windows-WinRM/Operational') 

# Note: you need Admin privileges to clear most event logs.

###########################################################################################################################################################################################