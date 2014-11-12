# Reference site: http://powershell.com/cs/blogs/tips/
###########################################################################################################################################################################################

# Tip 1: Creating Dynamic Breakpoints

# This one would halt whenever ForEach-Object is executed
Set-PSBreakpoint -Command ForEach-Object                                     


# This one would halt for each object that has the noun "Object", but only if the command is called from the script that is in the current ISE tab
Set-PSBreakpoint -Command *-Object -Script { $psISE.CurrentFile.FullPath }


# Even more advanced debugging is doable. This line halts whenever the variable $cpu gets assigned a $null value:
Set-PSBreakpoint -Variable cpu -Mode Write -Script ($psISE.CurrentFile.FullPath) -Action { if($cpu -eq $null) { break } }


Get-PSBreakpoint | Remove-PSBreakpoint     # remove all breakpoints once done

###########################################################################################################################################################################################

# Tip 2: Errors Travel in the Opposite Direction

1..10 | ForEach-Object { Trap { Write-Host "Phew: $_"; continue } $_ } | ForEach-Object {

    if($_ -gt 4)
    {
        Throw "Too big!"
    }
    else
    {
        $_
    }
}

# Ten numbers are fed into the pipeline. They reach the downstream cmdlet ForEach-Object which basically only implements an error handler (trap) 
# and then passes the number on to the next downstream cmdlet which happens to be another ForEach-Object.
# The second ForEach-Object checks the number, and if it is greater than 4, it throws an error.

# Pipelines usually work strictly downstream, so if something happens in a downstream cmdlet, upstream cmdlets won't notice. 
# Errors (exceptions), however, travel the opposite direction, so when you run the code, this is the result:

# Output:
#         1
#         2
#         3
#         4
#         Phew: Too big!
#         Phew: Too big!
#         Phew: Too big!
#         Phew: Too big!
#         Phew: Too big!
#         Phew: Too big!

# Although the second ForEach-Object encountered some special situation, it is the first (upstream) cmdlet that responds to it. 

# This basically enables you to establish an inner-pipeline-communication-system where a downstream cmdlet can tell an upstream cmdlet: "I've had enough"! 
# Coincidentally, this is how PowerShell 3.0 implemented Select-Object with its parameter -First. Once the first x elements are received, upstream cmdlets stop.

###########################################################################################################################################################################################

# Tip 3: Writing Back WMI Property Changes

# Only a few properties in WMI objects are actually writeable although Get-Member insists they are all "Get/Set":

Get-WmiObject -Class Win32_OperatingSystem | Get-Member -MemberType Properties

# That's because WMI only provides "copies" of information to you, and you can do whatever you want. WMI won't care. 
# To make WMI care and put property changes into reality, you need to send back your copy to WMI by calling Put(). 

# This will change the description of your operating system: (if you have Admin rights):
$os = Get-WmiObject -Class Win32_OperatingSystem
$os.Description = "I changed this!"
$result = $os.PSBase.Put()

# If you change a property that in reality is not writeable, you'll get an error message only once you try and write it back to WMI. 
# The truly changeable WMI properties need to be retrieved from WMI:
$class = [Wmiclass]"Win32_OperatingSystem"
$class.Properties | Where-Object { $_.Qualifiers.Name -contains "Write" } | Select-Object -Property Name, Type

###########################################################################################################################################################################################

# Tip 4: Examine "Extended" Object Members

# PowerShell is based on .NET objects but often refines them by adding more. If you'd like to see just what PowerShell has added, use Get-Member with its parameter -View Extended. 
# You'll be surprised how many useful properties are invented by PowerShell and not present in pure .NET objects of that type:

Get-Process | Get-Member -View Extended

###########################################################################################################################################################################################

# Tip 5: Renaming Object Properties in Powershell

# Let's say you want to output just your top-level processes like this:
Get-Process | Where-Object { $_.MainWindowTitle } | Select-Object name, product, id, mainwindowtitle

# This works like a charm, but you'd like to rename the column "MainWindowTitle" to "Title" only. That's what AliasProperties are for:
Get-Process | Where-Object { $_.MainWindowTitle } | Add-Member -MemberType AliasProperty -Name Title -Value MainWindowTitle -PassThru | Select-Object Name, Product, ID, Title

###########################################################################################################################################################################################

# Tip 6: Listing Currently Loaded Format Files

# The internal PowerShell formatting system (called ETS) relies on XML-based formatting data that comes from .ps1xml files. 
# To see all files currently loaded by PowerShell and 3rd party extensions, use this line:

$Host.Runspace.InitialSessionState.Formats | Select-Object -ExpandProperty FileName

# Output(part of them):
#        C:\Windows\System32\WindowsPowerShell\v1.0\Certificate.format.ps1xml
#        C:\Windows\System32\WindowsPowerShell\v1.0\DotNetTypes.format.ps1xml
#        C:\Windows\System32\WindowsPowerShell\v1.0\FileSystem.format.ps1xml
#        C:\Windows\System32\WindowsPowerShell\v1.0\Help.format.ps1xml

###########################################################################################################################################################################################

# Tip 7: Getting Help for Objects - Online

# In PowerShell 3.0, you finally can extend object types dynamically without having to write and import ps1xml-files. Here is an especially useful example:

$code = {

    $url = 'http://msdn.microsoft.com/en-US/library/{0}(v=vs.80).aspx' -f $this.GetType().FullName
    
    Start-Process $url
}

Update-TypeData -MemberType ScriptMethod -MemberName GetHelp -Value $code -TypeName System.Object

# Once you execute this code, every single object has a new method called GetHelp(), and when you call it, your browser will open and show the MSDN documentation page for it
#  - provided the object you examined was created by Microsoft, of course. There are many ways how you can call GetHelp(), for example:

$thedate = Get-Date
$thedate.GetHelp()
(Get-Date).GetHelp()        # Open link: http://msdn.microsoft.com/en-US/library/System.DateTime(v=vs.80).aspx

###########################################################################################################################################################################################

# Tip 8: Testing Numbers and Date

# With a bit of creativity (and the help from the -as operator), you can create powerful test functions. These two test for valid numbers and valid DateTime information:

function Test-Numeric($value)
{
    ($value -as [int64]) -ne $null
}
Test-Numeric 2.5     # True
Test-Numeric "2.3"   # True
Test-Numeric "2.3a"  # False



function Test-Date($value)
{
    ($value -as [datetime]) -ne $null
}
Test-Date "2014/10/29"  # True
Test-Date "29/10/2014"  # False



# If things get more complex, you can combine tests. This one tests for valid IP addresses:

function Test-IPAddress($value)
{
    ($value -as [System.Net.IPAddress]) -ne $null 
}
Test-IPAddress "localhost"   # False
Test-IPAddress "127.0.0.1"   # True

# Note: the limit here is the type converter. Sometimes, you may be surprised what kind of values are convertible into a specofic type. 
# The date check, on the other hand, illustrates why this can be very useful. Most cmdlets that require dates support the same type conversion rules. 
# So this boils down to the question who is going to process the validated data and whether this target supports the same conversion range.

###########################################################################################################################################################################################

# Tip 9: Get List of Type Accelerators

# Ever wondered what the difference between [Int], [Int32], and [System.Int32] is? They all are data types, and the first two are type accelerators, so they are really all the same. 

# To list all the type accelerators PowerShell provides, use this undocumented (and unsupported) call:

[PSObject].Assembly.GetType("System.Management.Automation.TypeAccelerators")::Get

# Output(part of them):
#        Key                        Value                                                                                                     
#        ---                        -----                                                                                                                                                                                  
#        int                        System.Int32                                                                                              
#        int32                      System.Int32                                                                                              
#        int16                      System.Int16                                                                                              
#        long                       System.Int64                                                                                              
#        int64                      System.Int64                                                                                              
#        wmiclass                   System.Management.ManagementClass                                                                         
#        wmi                        System.Management.ManagementObject                                                                        
#        wmisearcher                System.Management.ManagementObjectSearcher                                                                
#        ciminstance                Microsoft.Management.Infrastructure.CimInstance                                                           
#        NullString                 System.Management.Automation.Language.NullString                                                                                                                                      
#        timespan                   System.TimeSpan                                                                                           
#        uint16                     System.UInt16                                                                                             
#        uint32                     System.UInt32                                                                                             
#        uint64                     System.UInt64                                                                                             
#        uri                        System.Uri                                                                                                
       
###########################################################################################################################################################################################

# Tip 10: Identifying .NET Framework 4.5

# PowerShell 3.0 can run both on .NET Framework 4.0 and 4.5. .NET Framework 4.5 adds additional objects and members, 
# so for example this line will list the members of an enumeration such as System.ConsoleColor in .NET 4.5, but not in .NET 4.0:


[System.ConsoleColor].DeclaredMembers.Name        # # requires .NET 4.5

# Output:
#         value__
#         Black
#         DarkBlue
#         DarkGreen
#         DarkCyan
#         DarkRed
#         DarkMagenta
#         DarkYellow
#         Gray
#         DarkGray
#         Blue
#         Green
#         Cyan
#         Red
#         Magenta
#         Yellow
#         White

# Of course there are ways to work around this but that's not the point here. The point is: the version of .NET on your machine can make a difference to your code. 
# How can PowerShell code determine the .NET version present?

# Interestingly, $PSVersionTable returns a version 4.0 both for .NET 4.0 and .NET 4.5. 
# To test for .NET 4.5, you need to look at the revision number and check whether it is greater than 17.000:

($PSVersionTable.CLRVersion.Major -eq 4 -and $PSVersionTable.CLRVersion.Revision -gt 17000)                # Output: True

###########################################################################################################################################################################################

# Tip 11: Why Using Here-Strings?

$text = 'First Line
Second Line
Third Line'

$text
# Output:
#        First Line
#        Second Line
#        Third Line

# Why should you ever bother using the so-called here-strings? 

$text = @'
First Line
Second Line
Third Line
'@

$text
# Output:
#        First Line
#        Second Line
#        Third Line

# Here-strings have a strict format design: immediately after @' there has to be a new line, and the closing term '@ must start at the beginning of a new line. 
# Because of this, anything in between is treated as content, no matter what characters you use. 

# Here-strings mask special characters like quotes or hashes. Simple quotes would interpret them, for example as end-of-string or comment, respectively. 
# That's why here-strings are a safe way of enclosing source code or XML data.

$text = @'
<html>
<head><title>I'm Silence!</title></head>
<body>
 <h1>"Keep Silence!"</h1>
</body>
</html>
'@

$text
# Output:
#        <html>
#        <head><title>I'm Silence!</title></head>
#        <body>
#         <h1>"Keep Silence!"</h1>
#        </body>
#        </html>

###########################################################################################################################################################################################

# Tip 12: Finding Popular Historic First Names

# To find popular first names for given decades, check out the function Get-PopularName. It accepts a decade between 1880 and 2000 and then uses the new and awesome Invoke-WebRequest
# in PowerShell 3.0 to visit a statistical website and retrieve popular first names using a regular expression.

# Note: if your Internet connection requires a proxy server and/or authentication, please add the appropriate parameters to Invoke-WebRequest.

function Get-PopularName
{
    param
    (
        [ValidateSet('1880','1890','1900','1910','1920','1930','1940','1950','1960','1970','1980','1990','2000')]
        $decade = "1950"  
    )

    $regex = [regex]'(?si)<td>(\d{1,3})</td>\s*?<td align="center">(.*?)</td>\s*?<td>((?:\d{0,3}\,)*\d{1,3})</td>\s*?<td align="center">(.*?)</td>\s*?<td>((?:\d{0,3}\,)*\d{1,3})</td></tr>'
    
    $web = Invoke-WebRequest -UseBasicParsing -Uri "http://www.ssa.gov/OACT/babynames/decades/names$($decade)s.html"

    $html = $web.Content
    $Matches = $regex.Matches($html)

    $Matches | ForEach-Object {
    
        $rv = New-Object PSObject | Select-Object -Property Name, Rank, Number, Gender
        $rv.Rank = [int]$_.Groups[1].Value
        $rv.Gender = "m"
        $rv.Name = $_.Groups[2].Value
        $rv.Number =[int]$_.Groups[3].Value
        $rv

        $rv = New-Object PSObject | Select-Object -Property Name, Rank, Number, Gender
        $rv.Rank = [int]$_.Groups[1].Value
        $rv.Gender = "f"
        $rv.Name = $_.Groups[4].Value
        $rv.Number = [int]$_.Groups[5].Value
        $rv

    } | Sort-Object Name, Rank
}

Get-PopularName -decade 1900

###########################################################################################################################################################################################

# Tip 13: Using Specific Error Handlers

$oldValue = $ErrorActionPreference
$ErrorActionPreference = "Stop"

trap [System.Management.Automation.ItemNotFoundException]
{
    "Element not found: $_"
    continue
}

trap [System.DivideByZeroException]
{
    "Divided by zero"
    continue
}

trap [Microsoft.PowerShell.Commands.ProcessCommandException]
{
    "No such process running."
    continue
}

trap [System.Management.Automation.RemoteException]
{
    "Console command did not succeed: $_"
    continue
}

trap 
{
    "Console command did not succeed: $_"
    continue
}

Get-ChildItem -Path c:\notpresent
1/$null
Get-Process -Name notpresent
net.exe user noptresent 2>$1

$ErrorActionPreference = $oldValue

###########################################################################################################################################################################################

# Tip 14: Accessing PowerShell Host Process


# With this line, you always get back the process object representing the current PowerShell host:

[System.Diagnostics.Process]::GetCurrentProcess()

# or
Get-Process -Id $pid

# Either way, the process object returned can provide useful information like the host that runs your script, or the total time the process has been running:

[System.Diagnostics.Process]::GetCurrentProcess().Description            # Output: Windows PowerShell ISE
(Get-Process -Id $pid).Description                                       # Output: Windows PowerShell ISE

###########################################################################################################################################################################################

# Tip 15: Converting Date to WMI Date

$date = Get-Date
$wmiDate = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($date)

$wmiDate   # Output: 20141103104851.369841+480



$os = Get-WmiObject -Class Win32_OperatingSystem
$os.LastBootUpTime                                                                             # Output: 20141103085744.234799+480

$bootupTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootupTime)  
$bootupTime                                                                                    # Output: Monday, November 03, 2014 08:57:44

# Checking System Uptime

$timeSpan = New-TimeSpan -Start $bootupTime
$days = $timeSpan.TotalDays

"Your system is running for {0:0.0} days." -f $days

###########################################################################################################################################################################################

# Tip 16: Getting Weather Forecast from an Airfield near You

$weather = New-WebServiceProxy -Uri "http://www.webservicex.com/globalweather.asmx?WSDL"
$city = ([xml]$weather.GetCitiesByCountry("United States")).NewDataSet.Table | Select-Object -ExpandProperty City

# Next, you can specify one of the airfields to get current weather information:
$data = ([xml]$weather.GetWeather('Seattle, Seattle Boeing Field','United States')).CurrentWeather
$data 

# Output:
#        Location         : SEATTLE BOEING FIELD, WA, United States (KBFI) 47-33N 122-19W 4M
#        Time             : Nov 02, 2014 - 11:53 PM EST / 2014.11.03 0453 UTC
#        Wind             :  from the SE (140 degrees) at 8 MPH (7 KT):0
#        Visibility       :  10 mile(s):0
#        SkyConditions    :  overcast
#        Temperature      :  55.0 F (12.8 C)
#        DewPoint         :  50.0 F (10.0 C)
#        RelativeHumidity :  83%
#        Pressure         :  30.08 in. Hg (1018 hPa)
#        Status           : Success

# To get the current temperature, for example, you could access the appropriate object property:
$data.Temperature

($data.Temperature -split "[\(\)]")[1]     # Output: 12.8 C
($data.Temperature -split "\(")[0]         # Output:  55.0 F 

###########################################################################################################################################################################################

# Tip 17: Using PropertySets

# PropertySets are lists of properties, and PowerShell sometimes adds PropertySets to result objects to make picking the right information easier.

Get-Process | Get-Member -MemberType PropertySet
# Output:
#    TypeName: System.Diagnostics.Process
# 
# Name            MemberType  Definition                                                                                                                                                   
# ----            ----------  ----------                                                                                                                                                   
# PSConfiguration PropertySet PSConfiguration {Name, Id, PriorityClass, FileVersion}                                                                                                       
# PSResources     PropertySet PSResources {Name, Id, Handlecount, WorkingSet, NonPagedMemorySize, PagedMemorySize, PrivateMemorySize, VirtualMemorySize, Threads.Count, TotalProcessorTime}


Get-Process | Select-Object -Property PSConfiguration
# Output(Part of them):
#       Name                                                  Id PriorityClass                    FileVersion                                         
#       ----                                                  -- -------------                    -----------                                                                               
#       atiesrxx                                             944                                                                                      
#       BingDict                                            1756 Normal                           3.5.0.4311                                          
#       CcmExec                                             4268                                                                                      
#       conhost                                             7732 Normal                           6.1.7600.16385 (win7_rtm.090713-1255)                                                                 
#       devenv                                              8460 Normal                           11.0.60610.1 built by: Q11REL                                                                    
#       dwm                                                 2132 High                             6.1.7600.16385 (win7_rtm.090713-1255)               
#       explorer                                            2152 Normal                           6.1.7600.16385 (win7_rtm.090713-1255)               
#       FlashUtil64_15_0_0_189_ActiveX                      5692 Normal                           15,0,0,189                                                                                         
#       iexplore                                             540 Normal                           11.00.9600.16428 (winblue_gdr.131013-1700)                   
#       IMECMNT                                             2948 Normal                           14.0.5800.1000 

Get-Process | Select-Object -Property PSResources
# Output(Part of them):
#       Name               : BingDict
#       Id                 : 1756
#       HandleCount        : 3395
#       WorkingSet         : 140701696
#       PagedMemorySize    : 127700992
#       PrivateMemorySize  : 127700992
#       VirtualMemorySize  : 562143232
#       TotalProcessorTime : 00:02:06.1892089

# Note that in PowerShell 3.0, the parameter -Property supports auto-completion, so just enter "PS" and then press TAB to see all properties that start with "PS".

###########################################################################################################################################################################################

# Tip 18: Using MemberSets

# Objects can contain a MemberSet called PSStandardMembers which normally is hidden (along with a number of other MemberSets):

Get-Process | Get-Member -MemberType MemberSet             # return nothing
Get-Process | Get-Member -MemberType MemberSet -Force
# Output:
#           TypeName: System.Diagnostics.Process
#        
#        Name              MemberType Definition                                                                                                                                                                              
#        ----              ---------- ----------                                                                                                                                                                              
#        psadapted         MemberSet  psadapted {BasePriority, ExitCode, HasExited, ExitTime, Handle, HandleCount, Id, MachineName, MainWindowHandle, ...
#        psbase            MemberSet  psbase {BasePriority, ExitCode, HasExited, ExitTime, Handle, HandleCount, Id, MachineName, MainWindowHandle, Mai...
#        psextended        MemberSet  psextended {__NounName, Name, Handles, VM, WS, PM, NPM, Path, Company, CPU, FileVersion, ProductVersion, Descrip...
#        psobject          MemberSet  psobject {BaseObject, Members, Properties, Methods, ImmediateBaseObject, TypeNames, get_Members, get_Properties,...
#        PSStandardMembers MemberSet  PSStandardMembers {DefaultDisplayPropertySet} 

# PSStandardMembers controls the standard fallback properties that PowerShell displays. This line lists the default properties for service objects returned by Get-Service:

Get-Process | Select-Object -First 1 | ForEach-Object { $_.PSStandardMembers.DefaultDisplayPropertySet.ReferencedPropertyNames }
# Output:
#        Id
#        Handles
#        CPU
#        Name

###########################################################################################################################################################################################

# Tip 19: Auto-Discovering Online Help for WMI

Get-WmiObject -List                    # returns all available WMI class names

Get-WmiObject -Class *Share* -List     # search for WMI class name:
Get-WmiObject -Class Win32_Share       # retrieve all instances of WMI class

# To really dive into WMI, you will want to get full documentation for all the WMI classes. While documentation is available in the Internet, the URLs for those websites are cryptic. 

# Here is a clever way how you can translate WMI class names into the cryptic URLs: 
# use PowerShell 3.0 Invoke-WebRequest to search for the WMI class using a major search engine, and then retrieve the URL: 

function Get-WmiHelpLocation
{
    param($wmiClassName = "Win32_BIOS")

    $uri = 'http://www.bing.com/search?q={0}+site:msdn.microsoft.com' -f $wmiClassName         # Search this class name on site: http://www.bing.com

    $url = (Invoke-WebRequest -uri $uri -UseBasicParsing).Links | Where-Object { $_.href -like "http://msdn.microsoft.com*" } | Select-Object -ExpandProperty href -First 1
    $url

    Start-Process $url
}

Get-WmiHelpLocation      # Output: http://msdn.microsoft.com/en-us/library/aa394077(v=vs.85).aspx

# Note: If your Internet access requires a proxy server or authentication, add the appropriate parameters to Invoke-WebRequest inside the function Get-WmiHelpLocation.

###########################################################################################################################################################################################

# Tip 20: Displaying WMI Inheritance

# In PowerShell 3.0, the (hidden) object property PSTypeNames shows you the complete inheritance tree for WMI objects:

$os = Get-WmiObject -Class Win32_OperatingSystem
$os.pstypenames

# Output:
#        System.Management.ManagementObject#root\cimv2\Win32_OperatingSystem
#        System.Management.ManagementObject#root\cimv2\CIM_OperatingSystem
#        System.Management.ManagementObject#root\cimv2\CIM_LogicalElement
#        System.Management.ManagementObject#root\cimv2\CIM_ManagedSystemElement
#        System.Management.ManagementObject#Win32_OperatingSystem
#        System.Management.ManagementObject#CIM_OperatingSystem
#        System.Management.ManagementObject#CIM_LogicalElement
#        System.Management.ManagementObject#CIM_ManagedSystemElement
#        System.Management.ManagementObject
#        System.Management.ManagementBaseObject
#        System.ComponentModel.Component
#        System.MarshalByRefObject
#        System.Object

# As you can see, the instance of Win32_OperatingSystem is derived from CIM_LogicalElement, and you can use this WMI class instead to widen your query and get back all logical elements:

Get-WmiObject -Class CIM_LogicalElement | Select-Object -Property Caption, __Class

# In PowerShell 2.0, the property PSTypeNames does not include WMI inheritance information yet.

###########################################################################################################################################################################################

# Tip 21: Converting Date from French to Taiwanese

# Date and Time formats are highly culture-specific, so often you need to convert date and time from one cultural format to another. 
# That's pretty straight-forward. All you need to know are the culture IDs to convert from and convert to.

[System.Globalization.CultureInfo]::GetCultures("AllCultures") | Sort-Object DisplayName       # Get all curltures: LCID, Name, DisplayName

$dateFrench = 'vendredi 23 novembre 2012 11:19:13'

[System.Globalization.CultureInfo]$freach = "fr-FR"
[System.Globalization.CultureInfo]$taiwan = "zh-TW"

$dateTime = [DateTime]::Parse($dateFrench, $freach)
$dateTaiwan = $dateTime.ToString($taiwan)
$dateTaiwan

###########################################################################################################################################################################################

# Tip 22: Finding WMI Class Static Methods

# Note that WMI methods can be very powerful but are a low-level interface. This tip can only serve as a starter. 
# Once you find an interesting class and method, you will probably have to do additional research until all works as expected. 
# Also note that some WMI classes may not be implemented in all operating system versions.

$exclude = "SetPowerState", "Reset", "Invoke"

# The code lists all WMI classes that start with "Win32_" and have at least one method that is not listed in $exclude

Get-WmiObject -List -Class Win32_* | Where-Object { $_.Methods } | ForEach-Object {

    $result = $_ | Select-Object -Property Name, Methods

    [Object[]]$result.Methods = $_.Methods | Where-Object { ($_.Qualifiers | Select-Object -ExpandProperty Name) -contains "Static" } | ForEach-Object {
    
        if($exclude -notcontains $_.Name)
        {
            $_.Name
        }
    }

    $result

} | Where-Object { $_.Methods }

# Output:
#        Name                                                    Methods                                                                                                   
#        ----                                                    -------                                                                                                   
#        Win32_Process                                           {Create}                                                                                                  
#        Win32_BaseService                                       {Create}                                                                                                  
#        Win32_Service                                           {Create}                                                                                                  
#        Win32_TerminalService                                   {Create}                                                                                                  
#        Win32_SystemDriver                                      {Create}                                                                                                  
#        Win32_PrinterDriver                                     {AddPrinterDriver}                                                                                        
#        Win32_LogicalDisk                                       {ScheduleAutoChk, ExcludeFromAutochk}                                                                     
#        Win32_Volume                                            {ScheduleAutoChk, ExcludeFromAutoChk}                                                                     
#        Win32_Printer                                           {AddPrinterConnection}                                                                                    
#        Win32_Share                                             {Create}                                                                                                  
#        Win32_ClusterShare                                      {Create}                                                                                                  
#        Win32_ScheduledJob                                      {Create}                                                                                                  
#        Win32_DfsNode                                           {Create}                                                                                                  
#        Win32_ShadowCopy                                        {Create}                                                                                                  
#        Win32_NetworkAdapterConfiguration                       {RenewDHCPLeaseAll, ReleaseDHCPLeaseAll, EnableDNS, SetDNSSuffixSearchOrder...}                           
#        Win32_Product                                           {Install, Admin, Advertise}                                                                               
#        Win32_SecurityDescriptorHelper                          {Win32SDToSDDL, Win32SDToBinarySD, SDDLToWin32SD, SDDLToBinarySD...}                                      
#        Win32_ShadowStorage                                     {Create}                                                                                                  
#        Win32_ReliabilityStabilityMetrics                       {GetRecordCount}                                                                                          
#        Win32_ReliabilityRecords                                {GetRecordCount}                                                                                          
#        Win32_OfflineFilesCache                                 {Enable, RenameItem, RenameItemEx, Synchronize...}  

###########################################################################################################################################################################################

# Tip 23: Create share folder locally

# If you wanted to try Win32_Share with its method Create(), you then need to know the correct order of arguments for the method:

$class = [wmiclass]"Win32_Share"
$methodName = "Create"

$class.psbase.GetMethodParameters($methodName).Properties | Select-Object -Property Name, Type | Format-Table -AutoSize
# Output:
#        Name             Type
#        ----             ----
#        Access         Object
#        Description    String
#        MaximumAllowed UInt32
#        Name           String
#        Password       String
#        Path           String
#        Type           UInt32

# create new shares locally and remotely like this
$class = 'Win32_Share'
$methodname = 'Create'
$Access = $null
$Description = 'A new share'
$MaximumAllowed = 10
$Name = 'myNewShare'
$Password = $null
$Path = 'C:\Users\v-sihe\Desktop\PS'
$Type = 0
Invoke-WmiMethod -Path $class -Name $methodname -ArgumentList $Access, $Description, $MaximumAllowed, $Name, $Password, $Path, $Type

# To create the share remotely, add the parameter(s) -ComputerName and -Credential. Note that you will need local Administrator privileges to create a new share. 
# Invoke-WmiMethod returns numeric return values. To decipher these and find out more about WMI methods, 
# navigate to a search engine and enter the WMI class name (such as "Win32_Share"). Most WMI classes and their methods are well documented on MSDN.

###########################################################################################################################################################################################

# Tip 24: Checking Network Adapter Speed

# Sometimes, just one line of PowerShell code gets you all the information you may have needed. 
# There is a .NET type called NetworkInterface, for example, that lists all of your network adapters, their speed and status:

[System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | Select-Object Description, Speed, OperationalStatus

# Output:
#        Description                                                    Speed                            OperationalStatus
#        -----------                                                    -----                            -----------------
#        Broadcom NetXtreme Gigabit Ethernet                       1000000000                                           Up
#        Software Loopback Interface 1                             1073741824                                           Up
#        Microsoft ISATAP Adapter                                      100000                                         Down

###########################################################################################################################################################################################

# Tip 25: Finding Constructors (and submitting Credentials unattended)

$cred = Get-Credential
$cred.GetType().FullName            # Output: System.Management.Automation.PSCredential

# Here is how you can list the available constructor methods and view the additional information they require:

([Type]"System.Management.Automation.PSCredential").GetConstructors() | ForEach-Object { $_.ToString() }             # Output: Void .ctor(System.String, System.Security.SecureString)

# As you see, the constructor for credential objects needs two pieces of information: a string , and an encrypted string - the user name and the user password. 
# So here is how you create logon credentials that you then can use whenever a cmdlet provides a parameter -Credential:

$username = "domain\username"
$Password = "password" | ConvertTo-SecureString -Force -AsPlainText
$cred = New-Object -TypeName System.Management.Automation.PSCredential $username,$Password

###########################################################################################################################################################################################

# Tip 26: Discovering Network Access

$cat = "Public", "Private", "Domain"
$GUID = [GUID]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}'
$network = [Activator]::CreateInstance([type]::GetTypeFromCLSID($GUID))

$network.GetNetWorkConnections() | ForEach-Object {
    
    $result = $_ | Select-Object -Property Name, Description, *, Category

    $result.Name = $_.GetNetwork().GetName()
    $result.Description = $_.GetNetwork().GetDescription()
    $result.Category = $cat[$_.GetNetwork().GetCategory()]

    $result
}

# Output:
#        Name                  : corp.company.com
#        Description           : corp.company.com
#        IsConnectedToInternet : True
#        IsConnected           : True
#        Category              : Domain

###########################################################################################################################################################################################

# Tip 27: Parsing Custom Date and Time Formats

# Sometimes, date and time information is not formatted according to the standards PowerShell understands by default. 
# When this happens, you can provide a hint and tell PowerShell how to correctly interpret date and time.

$information = '12Nov(2012)18h30m17s'
$pattern = 'ddMMM\(yyyy\)HH\hmm\mss\s'
[datetime]::ParseExact($information, $pattern, $null)      # Output: Monday, November 12, 2012 18:30:17

# By providing the pattern in $pattern, PowerShell can still correctly interpret it. Note that the placeholders in $pattern are case-sensitive. 
# "MMM" represents a short month name, whereas "mm" stands for minutes. 
# The backslash escapes literals (text information that does not belong to the date and time format, for example braces or descriptive text).

###########################################################################################################################################################################################

# Tip 28: Installing Local Printer

# WMI represents all locally installed printers with its class Win32_Printer, so you can easily look what's installed:

Get-WmiObject -Class Win32_Printer | Select-Object -Property *

# To add a new local printer, just add a new instance of Win32_Printer. 
# The example adds a new local printer and shares it over the network (provided you have sufficient privileges and the appropriate printer drivers):

$printerclass = [wmiclass]'Win32_Printer'
$printer = $printerclass.CreateInstance()
$printer.Name = $printer.DeviceID = 'NewPrinter'
$printer.PortName = 'LPT1:'
$printer.Network = $false
$printer.Shared = $true
$printer.ShareName = 'NewPrintServer'
$printer.Location = 'Office 12'
$printer.DriverName = 'HP LaserJet 3050 PCL5'
$printer.Put()

# To find out what the properties are that you must set for a given printer, simply install the printer manually on a test system, then query the installed printer with the line above. 

# This will dump all the properties like driver name etc. that you need to set to install the printer via script.

###########################################################################################################################################################################################

# Tip 29: Discovering Date and Time Culture Information

[System.Globalization.CultureInfo]::CurrentUICulture.DateTimeFormat

# Output:
#        AMDesignator                     : AM
#        Calendar                         : System.Globalization.GregorianCalendar
#        DateSeparator                    : /
#        FirstDayOfWeek                   : Sunday
#        CalendarWeekRule                 : FirstDay
#        FullDateTimePattern              : dddd, MMMM dd, yyyy HH:mm:ss
#        LongDatePattern                  : dddd, MMMM dd, yyyy
#        LongTimePattern                  : HH:mm:ss
#        MonthDayPattern                  : MMMM dd
#        PMDesignator                     : PM
#        RFC1123Pattern                   : ddd, dd MMM yyyy HH':'mm':'ss 'GMT'
#        ShortDatePattern                 : yyyy/MM/dd
#        ShortTimePattern                 : hh:mm tt
#        SortableDateTimePattern          : yyyy'-'MM'-'dd'T'HH':'mm':'ss
#        TimeSeparator                    : :
#        UniversalSortableDateTimePattern : yyyy'-'MM'-'dd HH':'mm':'ss'Z'
#        YearMonthPattern                 : MMMM, yyyy
#        AbbreviatedDayNames              : {Sun, Mon, Tue, Wed...}
#        ShortestDayNames                 : {Su, Mo, Tu, We...}
#        DayNames                         : {Sunday, Monday, Tuesday, Wednesday...}
#        AbbreviatedMonthNames            : {Jan, Feb, Mar, Apr...}
#        MonthNames                       : {January, February, March, April...}
#        IsReadOnly                       : False
#        NativeCalendarName               : Gregorian Calendar
#        AbbreviatedMonthGenitiveNames    : {Jan, Feb, Mar, Apr...}
#        MonthGenitiveNames               : {January, February, March, April...}

# Any format defined here is a legal date and time format. 

# Note: When you cast a string to a DateTime type, PowerShell always uses the culture-invariant format. 
# Your own culture is honored when you use the -as operator. That's why these two lines may produce different results if you're not on the en-US culture:

[DateTime] '10/1/2013'             # Output: Dienstag, 1. Oktober 2013 00:00:00

'10/1/2013' -as [DateTime]         # Output: Donnerstag, 10. Januar 2013 00:00:00

###########################################################################################################################################################################################

# Tip 30: Playing WAV files

$player = New-Object System.Media.SoundPlayer "$env:windir\Media\notify.wav"
$player.Play()

# You can also use this as sort of an acoustic progress bar because the sound plays in a separate thread and won't block PowerShell. 
# So a script could repeat some sound until the task is done, then stop the sound:

$player.PlayLooping()
# do some lengthy job
$player.Stop()

# Note: The SoundPlayer object can play sounds in WAV format only, so if you'd like to play your favorite song or record some voice message, make sure you use this format.

###########################################################################################################################################################################################

# Tip 31: Using Open File Dialogs

# To spice up your scripts, PowerShell can use the system open file dialog, so users could easily select files to open or to parse. 

$dialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog

$dialog.AddExtension = $true
$dialog.Filter = "PowerShell-Script (*.ps1)|*.ps1|All Files|*.*"
$dialog.Multiselect = $false
$dialog.FilterIndex = 0
$dialog.InitialDirectory = "$home\documents"
$dialog.RestoreDirectory = $true
$dialog.ShowReadOnly = $true
$dialog.ReadOnlyChecked = $false
$dialog.Title = "Select a PS-Script"

$result = $dialog.ShowDialog()

if($result -eq "OK")
{
    $fileName = $dialog.FileName
    $readOnly = $dialog.ReadOnlyChecked

    if($readOnly)
    {
        $mode = "read-only"
    }
    else
    {
        "read-write"
    }

    "I could new open '$fileName' as $mode and do something ..."
}

# Note that the dialog requires the STA mode. The code runs fine in PowerShell 3.0 and in PowerShell 2.0 ISE. 
# It will crash PowerShell 2.0 consoles unless you start them with the parameter -STA.

# Note also that the dialog may occasionally appear behind the ISE editor, so if the editor seems to not respond, 
# check to see whether the dialog is waiting for your input in the background.

###########################################################################################################################################################################################

# Tip 32: New WMI Help Topics in PowerShell 3.0 Released

# PowerShell 3.0 comes with two new help topics (among many others) that may be especially useful for those who work a lot with WMI (or would like to dive into it):

help about_WMI -ShowWindow
help about_WQL -ShowWindow

# "about_WMI" is a great introduction to WMI in general, and "about_WQL" has all the secrets you need to know to query WMI for information.

# Should both files be missing, then you either do not run PowerShell 3.0, or you did not update the help files yet. 
# Should only "about_WQL" be missing, then we recommend you update your help again. 
# One of the benefits of PowerShell 3.0 updatable help is - that it is updatable! So every now and then, new help files like "about_WQL" are added.


# update your PowerShell 3.0 help: Open PowerShell with full administrator privileges, then enter:
Update-Help -Force

###########################################################################################################################################################################################

# Tip 33: Loading Additional Assemblies

# When you want to load additional .NET assemblies to extend the types of object you can use, there are two ways of loading them: the direct .NET approach and the Add-Type cmdlet.

# Method 1: 
$null = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
$date = [Microsoft.VisualBasic.Interaction]::InputBox("What is your birthday?")

$days = (New-TimeSpan -Start $date).Days

"You are $days days old."



# Method 2:
Add-Type -AssemblyName Microsoft.VisualBasic
$date = [Microsoft.VisualBasic.Interaction]::InputBox("What's your birthday?")

$days = (New-TimeSpan -Start $date).Days

"Your are $days days old."

# There is a crucial difference between these two methods, though. When you load an assembly using reflection, you load whatever version of that assembly exists on your machine. 
# With Add-Type, the assembly version is hard-coded to the latest assembly.

###########################################################################################################################################################################################

# Tip 34: Displaying MsgBox TopMost

# make the MsgBox stay on top of all other windows so it never gets covered and hidden in the background:

Add-Type -AssemblyName Microsoft.VisualBasic
$result = [Microsoft.VisualBasic.Interaction]::MsgBox('My Message', 'OKOnly,SystemModal,Information', 'Title')
$result

# Note the flag "SystemModal": it keeps the MsgBox dialog topmost so it cannot be covered by any other window. 

# To display different buttons and/or icons, take a look at all of the flags you can combine as a comma-separated list:

[System.Enum]::GetNames([Microsoft.VisualBasic.MsgBoxStyle])

# Output:
#        ApplicationModal
#        DefaultButton1
#        OkOnly
#        OkCancel
#        AbortRetryIgnore
#        YesNoCancel
#        YesNo
#        RetryCancel
#        Critical
#        Question
#        Exclamation
#        Information
#        DefaultButton2
#        DefaultButton3
#        SystemModal
#        MsgBoxHelp
#        MsgBoxSetForeground
#        MsgBoxRight
#        MsgBoxRtlReading

###########################################################################################################################################################################################

# Tip 35: Finding Built-In Variables

# Finding Built-In Variables

[PSObject].Assembly.GetType("System.Management.Automation.SpecialVariables").GetFields('NonPublic,Static') | Where-Object FieldType -eq ([string]) | ForEach-Object GetValue $null

# Note: This line with SpecialVariables doesn't work in PS2.


# Once you know the built-in variables, it is easy to sort them out and create a function that displays just your own variables. 
# As it turns out, there are still a couple of built-in variables left that need to be hand-coded, but then Get-MyVariable is a useful function to list all user variables:



# Get yourself variables but have no build-in varables included:

function Get-MyVariable
{
    $buildinVar = [PSObject].Assembly.GetType("System.Management.Automation.SpecialVariables").GetFields("NonPublic,Static") | 
        Where-Object FieldType -EQ ([String]) | ForEach-Object GetValue $null

    $buildinVar += 'MaximumAliasCount','MaximumDriveCount','MaximumErrorCount', 'MaximumFunctionCount', 'MaximumFormatCount', 'MaximumVariableCount', 
                    'FormatEnumerationLimit', 'PSSessionOption', 'psUnsupportedConsoleApplications'

    Get-Variable | Where-Object { $buildinVar -notcontains $_.Name } | Select-Object -Property Name,Value,Description
}

Get-MyVariable


$s1 = [PSObject].Assembly.GetType("System.Management.Automation.SpecialVariables").GetFields("NonPublic,Static") | 
        Where-Object {$_.FieldType -EQ ([String])} | ForEach-Object GetValue $null

$s2 = [PSObject].Assembly.GetType("System.Management.Automation.SpecialVariables").GetFields("NonPublic,Static") | 
        Where-Object {$_.FieldType -EQ ([String])} | ForEach-Object {$_.GetValue($null)}


$s1.ToString() -eq $s2.ToString()                    # Output: True


# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Finding Built-In Variables Part 2

[Powershell]::Create().AddCommand("Get-Variable").Invoke() | Select-Object -ExpandProperty Name

# we can turn this into a function called Get-MyVariable which lists only your own variables. 
# Like with the other approach, it is still necessary to hard-code a small number of built-in variables.

function Get-MyVariable
{
    $buildinVar = [Powershell]::Create().AddCommand("Get-Variable").Invoke() | Select-Object -ExpandProperty Name

    $buildinVar += 'args','MyInvocation','profile', 'PSBoundParameters', 'PSCommandPath', 'psISE', 'PSScriptRoot', 'psUnsupportedConsoleApplications'

    Get-Variable | Where-Object { $buildinVar -notcontains $_.Name } | Select-Object -Property Name,Value,Description
}

Get-MyVariable

###########################################################################################################################################################################################

# Tip 36: Adjust Text to Specific Length

$text1 = "some short text"
$text2 = "some very very very very very very very very very very very very long text"
$desiredLength = 20

$text1 = $text1.PadRight($desiredLength).Substring(0, $desiredLength)
$text2 = $text2.PadRight($desiredLength).Substring(0, $desiredLength)

$text1 += "<- ends here"
$text2 += "<- ends here"

# By combining PadRight() and SubString(), the text is adjusted to the desired length, no matter how long or short the text was before:

$text1         # Output: some short text     <- ends here
$text2         # Output: some very very very <- ends here

###########################################################################################################################################################################################

# Tip 37: Finding Keyboard and Mouse


# use WMI to quickly find all details about your mouse and keyboard

Get-WmiObject Win32_PointingDevice | Where-Object { $_.Description -match "hid" }
Get-WmiObject Win32_Keyboard | Where-Object { $_.Description -match "hid" }

# You can also use this code to detect if a mouse is connected at all:

function Test-Mouse
{
    @(Get-WmiObject Win32_PointingDevice | Where-Object {$Description -match "hid"}).Count -gt 0 
}


# Note: it display nothing on my computer, the description filtered with "hid" looks not suit on my machine, it should be removed here?

###########################################################################################################################################################################################

# Tip 38: New WMI Cmdlets with DateTime Support

# In PowerShell v3, to work with WMI you can still use the old WMI cmdlets like Get-WmiObject. There is a new set of CIM cmdlets, though, that pretty much does the same - but better.

# For example, CIM cmdlets return true DateTime objects. WMI cmdlets returned the raw WMI DateTime format. Compare the different results:

Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty LastBoot*        # Output: Tuesday, November 04, 2014 22:10:56
Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty LastBoot*              # Output: 20141104221056.103718+480


#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# CIM-Cmdlets Work Against Old Windows Boxes

# The new CIM cmdlets require PowerShell v3, but you can still remotely target older boxes without PowerShell v3 or PowerShell at all. 
# By default, CIM cmdlets use WSMan for remote connections. If you want to use the old DCOM technique, create a CIMSession and request DCOM.

# This piece of code targets the server "storage1" using DCOM, so the remote server can be any Windows box and won't need PowerShell. 
# Just the machine running the PowerShell code needs to have PowerShell v3 in place:

$opt = New-CimSessionOption -Protocol Dcom
$sd  = New-CimSession -ComputerName IIS-CTI5052 -SessionOption $opt
Get-CimInstance -CimSession $sd -ClassName Win32_BIOS

###########################################################################################################################################################################################

# Tip 39: Removing Leading "0" in IP Addresses

# Leading "0" in IP addresses can cause confusion because many network commands interpret octets with leading "0" as octal numbers:

# no leading "0":
ping 10.10.5.12                # Pinging 10.10.5.12 with 32 bytes of data:    

# leading "0":
ping 010.10.005.012            # Pinging 8.10.5.10 with 32 bytes of data:

# To remove leading "0" from IP addresses, use this piece of code:

"010.10.005.012" -replace "\b0*",""            # Output: 10.10.5.12
"010.10.005.012" -replace "\b0*\B",""          # Output: 10.10.5.12

###########################################################################################################################################################################################

# Tip 40: New Operator -In

# In PowerShell v3, you can use a new simplified syntax for Where-Object. Both lines below list all files in your Windows folder that are larger than 1MB:

dir $env:windir | Where-Object { $_.Length -gt 1MB }              # old syntax:
dir $env:windir | Where-Object Length -GT 1MB                     # # new alternate simplified syntax:

# The simplified syntax can produce strange results, too. Here is a more complex example. It uses a whitelist to list only certain services:

$whiteList = @("Spooler", "WinRM", "WinDefend")

Get-Service | Where-Object { $whiteList -contains $_.Name }       # old syntax:

Get-Service | Where-Object $whiteList -Contains Name              # new alternate simplified syntax: it fails this time

# As it turns out, when you do use the simplified syntax, the property you are referring to always must come first. So you cannot use operators like -contains and -notcontains. 
# That's why Microsoft has added the two new operators -in and -notin. They work exactly like -contains and -notcontains, just with reversed arguments:

Get-Service | Where-Object Name -In $whiteList

###########################################################################################################################################################################################

# Tip 41: Mixing DCOM and WSMan in WMI Queries

# Using the new CIM cmdlets in PowerShell v3, you can run remote WMI queries against multiple computers using multiple remoting protocols.

# The sample code below gets WMI BIOS information from five remote machines. It uses DCOM for the old machines in $OldMachines and uses WSMan for new machines in $NewMachines. 
# Get-CimInstance gets the BIOS information from all five machines 
# (make sure you adjust the computer names in both lists so that they match real computers in your environment, and that you have admin privileges on all of these machines):

# list of machines with no WSMan capabilities/no PSv3
$oldMachines = 'pc_winxp', 'pc_win7', 'server_win2003'

# list of new machines with PSv3 in place:
$newMachines = 'pc_win8', 'server_win2012'

$useDCOM  = New-CimSessionOption -Protocol DCOM 
$useWSMan = New-CimSessionOption -Protocol WSMan

$session  = New-CimSession -ComputerName $oldMachines -SessionOption $useDCOM
$session += New-CimSession -ComputerName $newMachines -SessionOption $useWSMan

# get WMI info from all machines, using appropriate protocol:
Get-CimInstance -CimSession $session -ClassName Win32_BIOS

###########################################################################################################################################################################################

# Tip 42: Listing Power Plans

# There is a somewhat hidden WMI namespace that holds WMI classes you can use to manage power plans.

# The code below lists all power plans on your local machine, and using -ComputerName, you can easily retrieve this information remotely as well:

Get-CimInstance -Namespace root\cimv2\power -ClassName Win32_PowerPlan
Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan

Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan | Select-Object ElementName, Description
# Output:
#        ElementName                             Description                                                                                               
#        -----------                             -----------                                                                                               
#        Balanced                                Automatically balances performance with energy consumption on capable hardware.                           
#        High performance                        Favors performance, but may use more energy.                                                              
#        Power saver                             Saves energy by reducing your computer’s performance where possible.

###########################################################################################################################################################################################

# Tip 43: Temporarily Activate High Performance Power Plan

# It may be useful to automatically and temporarily switch to a "high performance" power plan from inside a script. 
# Maybe you know that a script has to do CPU-intensive tasks, and you would like to speed it up a bit.


# save current power plan:
$powerPlan = (Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan -Filter 'isActive=True').ElementName
"Current Power Plan: $PowerPlan"


# turn on high performance power plan:
(Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan -Filter 'ElementName="High Performance"').Activate()


# so something here
"Power plan now is High Performance!"
Start-Sleep -Seconds 3


# turn power plan back to what is was before:
(Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan -Filter "ElementName='$PowerPlan'").Activate()
"Power Plan is back to $PowerPlan"

###########################################################################################################################################################################################

# Tip 44: Calling WMI Methods with CIM Cmdlets


# It can be very useful to call WMI methods, for example to create new shares, but in PowerShell v2 you had to know the names and exact order of arguments to submit:
$rv = Invoke-WmiMethod -Path "Win32_Share" -ComputerName $computerName -Name Create -ArgumentList $null, $Description, $MaximumAllowed, $Name, $null, $Path, $Type

# In PowerShell v3, you can use Get-CimClass to discover methods and arguments, and use Invoke-CimMethod instead. 
# It takes the arguments as a hash table, so order is no longer important:


$class = Get-CimClass -ClassName Win32_Share
$class.CimClassMethods
# Output:
#        Name                     ReturnType         Parameters                                           Qualifiers                                          
#        ----                     ----------         ----------                                           ----------                                          
#        Create                       UInt32         {Access, Description, MaximumAllowed, Name...}       {Constructor, Implemented, MappingStrings, Static}  
#        SetShareInfo                 UInt32         {Access, Description, MaximumAllowed}                {Implemented, MappingStrings}                       
#        GetAccessMask                UInt32         {}                                                   {Implemented, MappingStrings}                       
#        Delete                       UInt32         {}                                                   {Destructor, Implemented, MappingStrings}  


$class.CimClassMethods["Create"]
# Output:
#        Name                     ReturnType         Parameters                                           Qualifiers                                          
#        ----                     ----------         ----------                                           ----------                                          
#        Create                       UInt32         {Access, Description, MaximumAllowed, Name...}       {Constructor, Implemented, MappingStrings, Static} 


$class.CimClassMethods["Create"].Parameters
# Output:
#        Name                        CimType          Qualifiers                                           ReferenceClassName                                  
#        ----                        -------          ----------                                           ------------------                                  
#        Access                     Instance          {EmbeddedInstance, ID, In, MappingStrings...}                                                            
#        Description                  String          {ID, In, MappingStrings, Optional}                                                                       
#        MaximumAllowed               UInt32          {ID, In, MappingStrings, Optional}                                                                       
#        Name                         String          {ID, In, MappingStrings}                                                                                 
#        Password                     String          {ID, In, MappingStrings, Optional}                                                                       
#        Path                         String          {ID, In, MappingStrings}                                                                                 
#        Type                         UInt32          {ID, In, MappingStrings}   


# create new share
Invoke-CimMethod -ClassName Win32_Share -MethodName Create -Arguments @{ Name = "TestShare"; Path = "C:\Users\v-sihe\Desktop\test"; MaximumAllowed = [UInt32]4; Type = [UInt32]0 }

###########################################################################################################################################################################################

# Tip 45: Get CPU Load


# To get the average total CPU load for your local system or a remote system, use Get-Counter.

(Get-Counter "\processor(_total)\% processor time" -SampleInterval 10).CounterSamples.CookedValue            # returns the average total CPU load for a 10 second interval:

# Unfortunately, the name of the performance counter is a localized (!) value, so it is translated into the language of your culture, 
# and unless you are running an English copy of Windows, you'd have to look up the localized name. 

Get-Counter -ListSet * | Select-Object -ExpandProperty Counter                                               # To dump the names of all performance counters, use this line:


# To query a remote system, prepend the computer name. This will get the CPU load from a server called 'storage1':
(Get-Counter '\\storage1\processor(_total)\% processor time' -SampleInterval 10).CounterSamples.CookedValue

###########################################################################################################################################################################################

# Tip 46: Calculating Time Differences Using Custom Formats


# Calculating time differences is easy - provided you can convert the date and time information into a DateTime type. 
# If the date is in a custom format, you can use the method ParseExact() and submit your own pattern.

$text = 'January16'
[System.Globalization.CultureInfo]$US = 'en-us'
$date = [System.DateTime]::ParseExact($text, 'MMMMdd', $US)
((Get-Date) - $date).Days


# Likewise, if the custom format uses localized month names (for example German names such as 'Januar16'), adjust the culture information:

$text = "Januar16"
[System.Globalization.CultureInfo]$DE = "de-DE"
$date = [System.DateTime]::ParseExact($text, "MMMMdd", $DE)
((Get-Date) - $date).Days

###########################################################################################################################################################################################

# Tip 47: Check Installed Server Roles and Features

# Beginning with Server 2008 R2, there is a PowerShell module called ServerManager that you can use to manage server features and optional components. 

Import-Module ServerManager       # Simply import the module in PowerShell v2:

# In PowerShell v3, the module is imported automatically on demand.

# Next, you can obtain a list of installed components, like this:


Get-WindowsFeature *file*
# Output: [Test on 2k8R2]
#        Display Name                                            Name
#        ------------                                            ----
#        [ ] File Services                                       File-Services
#            [ ] File Server                                     FS-FileServer
#                [ ] File Services Tools                         RSAT-File-Services


Get-WindowsFeature *powershell*
# Output: [Test on 2k8R2]

#        Display Name                                            Name
#        ------------                                            ----
#                    [ ] Active Directory module for Windows ... RSAT-AD-PowerShell
#        [X] Windows PowerShell Integrated Scripting Environm... PowerShell-ISE


Get-WindowsFeature *
# Output: [Test on 2k8R2]

#         Display Name                                            Name
#         ------------                                            ----
#         [ ] Active Directory Certificate Services               AD-Certificate
#             [ ] Certification Authority                         ADCS-Cert-Authority
#             [ ] Certification Authority Web Enrollment          ADCS-Web-Enrollment
#             [ ] Online Responder                                ADCS-Online-Cert
#             [ ] Network Device Enrollment Service               ADCS-Device-Enrollment
#             [ ] Certificate Enrollment Web Service              ADCS-Enroll-Web-Svc
#             [ ] Certificate Enrollment Policy Web Service       ADCS-Enroll-Web-Pol
#         [ ] Active Directory Domain Services                    AD-Domain-Services
#             [ ] Active Directory Domain Controller              ADDS-Domain-Controller
#             [ ] Identity Management for UNIX                    ADDS-Identity-Mgmt
#                 [ ] Server for Network Information Services     ADDS-NIS
#                 [ ] Password Synchronization                    ADDS-Password-Sync
#                 [ ] Administration Tools                        ADDS-IDMU-Tools
#         [ ] Active Directory Federation Services                AD-Federation-Services
#             [ ] Federation Service                              ADFS-Federation
#             [ ] Federation Service Proxy                        ADFS-Proxy
#             [ ] AD FS Web Agents                                ADFS-Web-Agents
#                 [ ] Claims-aware Agent                          ADFS-Claims
#                 [ ] Windows Token-based Agent                   ADFS-Windows-Token
      
###########################################################################################################################################################################################

# Tip 48: Using Safe Cmdlets Only

# Let's assume you want to set up a restricted PowerShell v3 console that just provides access to Microsoft cmdlets with the verb Get. 
# One way to do this is to create a custom module that publishes the cmdlets you want to keep, then to remove all other modules:

$PSModuleAutoLoadingPreference = "none"

Get-Module | Remove-Module

New-Module -Name SafeSubSet
{
    Get-Module Microsoft* -ListAvailable | Import-Module
    Export-ModuleMember -Cmdlet Get-*, Import-Module

} | Import-Module


# This isn't enough, though, because the PowerShell core snap-in is still there and cannot be removed. 
# It provides cmdlets like Import-Module, so a user could go ahead and re-import modules. That's why you should mark all unwanted core cmdlets as "private", effectively hiding them:

Get-Command -Noun Module*,Job,PSSnapin,PSSessionConfiguration*,PSRemoting | ForEach-Object { $_.Visibility = "Private" }
Get-Command -Verb New | ForEach-Object { $_.Visibility = "Private" }

###########################################################################################################################################################################################

# Tip 49: Get Quotes From the Webservices

# There are plenty of free webservices around, and provided you have direct Internet access (no proxy), you can use New-WebServiceProxy to access them. 

$service = New-WebServiceProxy -Uri 'http://ryanrusson.com/ws/WS.asmx?WSDL'
$service.ChimpOmatic().Tables[0] | Select-Object -ExpandProperty Quote

# Output: ou've heard Al Gore say he invented the Internet.  Well, if he was so smart, why do all the addresses begin with W?

###########################################################################################################################################################################################

# Tip 50: Create Strongly Typed Hash Table

# A hash table can store any data type. If you want more control, you can create a typed dictionary (which behaves pretty much like a hash table):

$hashTable = @{}
$hashTable.ID = 12
$hashTable.Persion = "silence"

$hashTable
# Output:
#        Name                           Value                                                                                                                                                                                 
#        ----                           -----                                                                                                                                                                                 
#        ID                             12                                                                                                                                                                                    
#        Persion                        silence 


$hashTableTyped = New-Object "System.Collections.Generic.Dictionary[string,int]"
$hashTableTyped.ID = 12
$hashTableTyped.Persion = "silence"     # Exception here: Cannot convert value "silence" to type "System.Int32". Error: "Input string was not in a correct format."


# Note: $hashTable is a classic hash table, and you can store anything in it. 
# $hashTableTyped is a typed dictionary. It accepts only strings as key and integers as value. That's why you get an exception when you try to store a name in it.

###########################################################################################################################################################################################

# Tip 51: Dumping Scheduled Tasks

schtasks.exe /query /fo csv | ConvertFrom-Csv | Where-Object { $_.TaskName -ne "TaskName" } | Sort-Object TaskName | Out-GridView -Title "All Scheduled Tasks"
schtasks.exe /query /v /fo csv | ConvertFrom-Csv | Where-Object { $_.TaskName -ne "TaskName" } | Sort-Object TaskName | Out-GridView -Title "All Scheduled Tasks"

# You can even query remote computers when you add the /S Servername switch.

###########################################################################################################################################################################################

# Tip 52: Launching Applications with Alternate Credentials

# If you must run an application with a different identity, Start-Process offers the parameter -Credential. 

Start-Process -FilePath notepad -Credential mydomain\myuser                          # This would launch the Notepad editor using the context of user mydomain\myuser:

# However, you may run into errors. When you switch user context, the application still uses your current path. 
# If the new identity has no access permission to your current path, or your current path is a mapped network drive that is not available for the new user context, then the call fails.


# That's why you should always make sure you also define a working directory that is accessible:

Start-Process -FilePath notepad -Credential mydomain\myuser -WorkingDirectory C:\

# And if you want to also load the user profile, add -LoadUserProfile.

###########################################################################################################################################################################################

# Tip 53: Finding Next or Last Sunday

[int]$index = Get-Date | Select-Object -ExpandProperty DayofWeek
$daysTillNextSunday = 7 - $index
$daysSinceLastSunday = $index

$nextSunday = (Get-Date) + (New-TimeSpan -Days $daysTillNextSunday)
$lastSunday = (Get-Date) - (New-TimeSpan -Days $daysSinceLastSunday)

$nextSunday             # Output: Sunday, November 09, 2014 17:45:20
$lastSunday             # Output: Sunday, November 02, 2014 17:45:20

###########################################################################################################################################################################################

# Tip 54: Get Fully Qualified Domain Name

# Method 1:
ping -a localhost


# Method 2:
[System.Net.Dns]::GetHostByName("").HostName
Invoke-Command $server1,$server2 { [System.Net.Dns]::GetHostByName("").HostName }     # use this on multiple remote hosts


# Method 3:
$env:COMPUTERNAME + "." + $env:USERDNSDOMAIN

###########################################################################################################################################################################################

# Tip 55: Pausing Console Output

# When a command in the PowerShell console outputs lots of data, you can press CTRL+S to pause the output. To continue, press any key.

# Ctrl + S

###########################################################################################################################################################################################

# Tip 56: Get-Member Receives Array Contents

# If you need to know the object nature of command results, you probably know that you can pipe them to Get-Member like this:

Get-Process | Get-Member

# If you want to examine an array (or make sure that the object you want to examine is not altered by the pipeline), submit it to Get-Member directly:

Get-Member -InputObject (Get-Process)

###########################################################################################################################################################################################

# Tip 57: Secret Script Block Parameters

# If you think you understand PowerShell parameter binding, then have a look at this simple function which exposes a little-known PowerShell behavior:

function Test-Function
{
    param
    (
        [Parameter(ValueFromPipeline = $true)][int]
        $number,

        [Parameter(ValueFromPipelineByPropertyName = $true)][int]
        $number2
    )

    process
    {
        "Doing something with $number and $number2"
    }
}

1..5 | Test-Function -number2 3     # Note: -number not visable here.

# The function Test-Function accepts two parameters. Both can come from the pipeline. 
# The first is accepting the entire pipeline input ("ByValue"), the second expects pipeline objects with a property "Number2" ("ByPropertyName"). 
# So what do you think would happen when you call it like this:

# Output:
#        Doing something with 1 and 3
#        Doing something with 2 and 3
#        Doing something with 3 and 3
#        Doing something with 4 and 3
#        Doing something with 5 and 3

# Right: The numbers sent via pipeline go to parameter -Number, and the parameter -Number2 is always 6. 

# However, when you submit a script block to -Number2, everything changes: PowerShell realizes that a script block is not an Integer. 
# The call would have to fail with a type mismatch, but it doesn't. Instead, PowerShell happily accepts the script block and evaluates it, then takes its result as parameter value:

1..5 | Test-Function -number2 { $_ * 2 }

# Output:
#        Doing something with 1 and 2
#        Doing something with 2 and 4
#        Doing something with 3 and 6
#        Doing something with 4 and 8
#        Doing something with 5 and 10

# This is unexpected, yet very useful because it allows you to create parameters that dynamically process the data piped into the command. A lot of cmdlets use the very same technique.

###########################################################################################################################################################################################

# Tip 58: Finding Built-In ISE Keyboard Shortcuts

$gps = $psISE.GetType().Assembly
$rm = New-Object System.Resources.ResourceManager GuiStrings,$gps

$rs = $rm.GetResourceSet((Get-Culture),$true,$true)

$rs | where Name -Match "Shortcut\d?$|^F\d+Keyboard" | Sort-Object Value | Format-Table -AutoSize

###########################################################################################################################################################################################

# Tip 59: Auto-Documenting Script Variables

# PowerShell can automatically find and list all variables that you use in a script. 
# This way, you can easily create variable documentation for your scripts (and also find variables that may be misspelled):


function Get-ISEVariable                
{
    $text = $psISE.CurrentFile.Editor.Text

    [System.Management.Automation.PSParser]::Tokenize($text, [ref]$null) | Where-Object { $_.Type -eq "Variable" } | ForEach-Object {
    
        $rv = 1 | Select-Object -Property Line, Name, Code
        $rv.Name = $text.SubString($_.Start, $_.Length)
        $rv.Line = $_.StartLine

        $psISE.CurrentFile.Editor.SetCaretPosition($_.StartLine, 1)
        $psISE.CurrentFile.Editor.SelectCaretLine()
        
        $rv.Code = $psISE.CurrentFile.Editor.SelectedText.Trim()

        $rv
    }
}

Get-ISEVariable | Out-GridView        # When you run Get-ISEVariable in your ISE editor, you get a dump of all variables found in the currently opened script:

Get-ISEVariable | Select-Object -ExpandProperty Name | Sort-Object -Unique                               # To get a list of all variables used in your script

Get-ISEVariable | Group-Object -Property Name -NoElement | Sort-Object -Property Count -Descending       # You can even document how often a variable was used

###########################################################################################################################################################################################

# Tip 60: Renaming Script Variables

# Often, when you start writing a PowerShell script, you use temporary (dirty) variable names that later you'd like to polish. 
# Renaming script variables by find/replace is not a very good idea, though, 
# because searching for text can find text inside of strings, commands or other variable names, and potentially wrecks your entire script.

function Rename-ISEVariable
{
    param
    (
        [Parameter(Mandatory = $true)]
        $oldName,

        [Parameter(Mandatory = $true)]
        $newName
    )

    $oldName = $oldName.TrimStart("$")
    $newName = $newName.TrimStart("$")

    $text = $psISE.CurrentFile.Editor.Text

    $sb = New-Object System.Text.StringBuilder $text

    $variables = [System.Management.Automation.PSParser]::Tokenize($text, [ref]$null) | Where-Object { $_.Type -eq "Variable" } | Sort-Object -Property Start -Descending | 
        
        Where-Object { $text.SubString($_.Start + 1, $_.Length - 1) -eq $oldName } | ForEach-Object {
            
            $sb.Remove($_.Start + 1, $_.Length - 1)
            $sb.Insert($_.Start + 1, $newName)
        }

    $psISE.CurrentFile.Editor.Text = $sb.ToString()
}

# Rename-ISEVariable lets you easily and safely rename variables in your scripts:

Rename-ISEVariable -oldName beforeRename -newName afterRenamed

# As you can see, you can replace variables with completely new names, or just change casing. 
# Thanks to the parser, Rename-ISEVariable will change only the variable names you were after. 
# It will not touch strings, commands or other variable names that just contain the text you are looking for.


# To run this code from a PowerShell console, you have to load the WPF assemblies first:

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# This step is not required inside of the ISE editor because ISE loads these assemblies automatically.

###########################################################################################################################################################################################

# Tip 61: Temporarily Locking the Screen

# PowerShell 3.0 uses .NET Framework 4.x so it has WPF (Windows Presentation Foundation) capabilities built-in. This way, it only takes a few lines of code to generate GUI elements.

# Here's a sample function called Lock-Screen that places a transparent overlay window on top of your screen. You could use it to temporarily lock out user interaction, for example:

function Lock-Screen
{
    param
    (
        $title = "Go away and come back in 10 seconds ...",

        $delay = 10
    )

    $window = New-Object Windows.Window
    $label = New-Object Windows.Controls.Label

    $label.Content = $title
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
    Start-Sleep -Seconds $delay
    $window.Close()
}

# To lock the screen, call Lock-Screen. By default, it will lock the screen for 10 seconds. Use the -Delay parameter to specify a different time interval.

Lock-Screen

# The function simply sleeps while the screen is locked, but you could of course replace Start-Sleep with something more useful that you'd like to do while the screen is locked.

###########################################################################################################################################################################################

# Tip 62: Documenting CPU Load for Running Processes

# Get-Process can easily return the CPU load for each running process, but this information by itself is not very useful:
Get-Process | Select-Object -Property Name, CPU, Description

# (Note that if you run this code from an elevated PowerShell, you see information for all processes. Else, you can only examine CPU load for your own processes.)

# The CPU property returns the number of seconds the particular process has utilized the CPU. 
# A high number does not necessarily indicate high CPU usage, though, because this all depends on how long a process has been running. 
# CPU is an accumulated number, so the longer a process runs, the higher the CPU load is.


# A much better indicator for CPU-intensive processes is to set CPU load in relation with process runtime and calculate the average CPU load percentage.

$cpuPercent = @{

    Name = "CPU Percent (%)"
    
    Expression = {
    
        $totalSec = (New-TimeSpan -Start $_.StartTime).TotalSeconds
        [Math]::Round( ($_.CPU * 100 / $totalSec), 2 )
    }
}

# This piece of code returns the top 4 CPU-intensive running processes:
Get-Process | Select-Object -Property Name, CPU, $cpuPercent, Description | Sort-Object -Property "CPU Percent (%)" -Descending | Select-Object -First 4

# Output:
#        Name                                   CPU                   CPU Percent (%)                   Description                                         
#        ----                                   ---                   ---------------                   -----------                                         
#        powershell_ise                 638.2780915                              2.38                   Windows PowerShell ISE                              
#        lync                           475.1478458                              1.74                   Microsoft Lync                                      
#        dwm                            391.7029109                              1.43                   Desktop Window Manager                              
#        iexplore                         49.452317                              1.23                   Internet Explorer   

# The code uses a hash table to create a new calculated property called CPUPercent. To calculate the value, New-Timespan determines the total seconds a process is running. 
# Then, the accumulated CPU load is divided by that number and returns the CPU load percentage per second.


# Note also how the result is rounded. By rounding a numeric result, 
# you can control the digits after the decimal without changing the numeric data type, so the result can still be sorted correctly.
  
###########################################################################################################################################################################################

# Tip 63: Identifying Origin of IP Address

$ip = "180.76.3.151"

$geoIP = New-WebServiceProxy 'http://www.webservicex.net/geoipservice.asmx?WSDL' 
$geoIP.GetGeoIP($ip)

# Output:
#        ReturnCode        : 1
#        IP                : 180.76.3.151
#        ReturnCodeDetails : Success
#        CountryName       : China
#        CountryCode       : CHN

# Note that for New-WebServiceProxy to work, you need direct Internet access or else may have to use -Credential to submit logon details.

# If you'd instead like to know the IP address that your ISP has assigned to you, and find out where you are currently located, use GetGeoIPContext():
$geoIP.GetGeoIPContext()


# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Finding Public IP Address


# Whenever a machine is connected to the Internet, it gets a public IP address (typically assigned by the ISP), and this public IP address is not identical to your machine's IP address.

# In PowerShell 3.0, there is a new and extremely useful cmdlet called Invoke-WebRequest that you can use to contact web servers and download data. 
# Here is a simple one-liner that tells you your current public IP address:

(Invoke-WebRequest "http://myip.dnsomatic.com" -UseBasicParsing).Content

# Note that you may need to add the -Proxy and –ProxyCredential parameters if you have a proxy server in place.

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Finding IP GeoLocation Data

# Often, you can get the very same information by contacting websites that provide XML results. 
# Once you know such sites, you can use the new Invoke-WebRequest cmdlet introduced in PowerShell 3.0 to simply download the information you are after.

# For example, if you'd like to nail down an IP address to a specific location, here's a quick and easy solution:

$myPublicIP = (Invoke-WebRequest "http://myip.dnsomatic.com" -UseBasicParsing).Content

$html = Invoke-WebRequest -Uri "http://freegeoip.net/xml/$myPublicIP" -UseBasicParsing
$content = [xml]$html.Content
$content.Response

# In this sample, the first call to Invoke-WebRequest uses a website that provides your own public IP address. The website returns plain text.

# Then, another website is contacted to translate the IP address into geo information. 
# This website returns XML information which is why the script translates the received results to XML and returns the relevant data.

###########################################################################################################################################################################################

# Tip 64: Replacing Aliases with Command Names

# Aliases are shortcuts for commands and useful in interactive PowerShell. 
# Once you write scripts, however, you should no longer use aliases and instead use the underlying commands directly.

# You can always use Get-Command to find out the underlying command for an alias:
Get-Command dir

# Output:
#        CommandType            Name                           ModuleName                                                                                                                                        
#        -----------            ----                           ----------                                                                                                                                        
#        Alias                  dir -> Get-ChildItem                                                                                                                                                             



# The replacement process can be automated, too. Here's an experimental function called Remove-ISEAlias. 
# When used from inside the ISE 3.0 editor, the function automatically replaces all aliases with their underlying commands:

function Remove-ISEAlias
{
    $text = $psISE.CurrentFile.Editor.Text
    $sb = New-Object System.Text.StringBuilder $text

    $commands = [System.Management.Automation.PSParser]::Tokenize($text, [ref]$null) | 
        
        Where-Object { $_.Type -eq "Command" } | Sort-Object -Property Start -Descending | ForEach-Object {
        
            $command = $text.SubString($_.Start, $_.Length)
            $commandType = @(try { Get-Command $command -ErrorAction 0 } catch{} )[0]

            if($commandType -is [System.Management.Automation.AliasInfo])
            {
                $sb.Remove($_.Start, $_.Length)
                $sb.Insert($_.Start, $commandType.ResolvedCommandName)
            }

            $psISE.CurrentFile.Editor.Text = $sb.ToString()
        }
}

# Note that this is just a simple proof-of-concept. Do not use it as-is on large production scripts. This is not a well-tested commercial solution. 

###########################################################################################################################################################################################

# Tip 65: Getting Holiday Dates

$uri = "http://www.holidaywebservice.com/HolidayService_v2/HolidayService2.asmx"
$proxy = New-WebServiceProxy -Uri $uri -Class holiday -Namespace webservice        # connect to the holiday web service

$proxy.GetCountriesAvailable()

# Output:
#         Code                                        Description                                                                                               
#         ----                                        -----------                                                                                               
#         Canada                                      Canada                                                                                                    
#         GreatBritain                                Great Britain and Wales                                                                                   
#         IrelandNorthern                             Northern Ireland                                                                                          
#         IrelandRepublicOf                           Republic of Ireland                                                                                       
#         Scotland                                    Scotland                                                                                                  
#         UnitedStates                                United States  

$proxy.GetHolidaysAvailable("UnitedStates")           # To list all holidays in the United States

# And if you'd like to know just what date New Year's Eve is next year, 
# use GetHolidayDate() and submit the holiday code you got from GetHolidaysAvailable(), plus the year you are interested in:

$proxy.GetHolidayDate("UnitedStates", "NEW-YEARS-EVE", 2014)                 # Output: Wednesday, December 31, 2014 00:00:00

###########################################################################################################################################################################################

# Tip 66： Changing Files without Changing Modification Date

# Whenever you change a file, the file system automatically updates the LastWriteTime property. If you'd like to change a file without leaving such traces, try this:


# make sure this file and folder exist!!
$path = "$home\test.txt"                 

$file = Get-Item $Path
$lastModified = $file.LastWriteTime

# change file content:
"More stuff" >> $path

# replace modification date
$file.LastWriteTime = $lastModified

# As you can see, Get-Item returns a file object that gives you easy read/write access to file information such as LastWriteTime, CreationTime, or LastAccessTime. 
# After the file is changed, the information is simply restored, and the updated file receives the old modification timestamp.

###########################################################################################################################################################################################

# Tip 67: Creating a Drawing Panel

# In PowerShell 3.0, WPF is a great (and easy) way of creating GUIs. If you have a touch-enabled machine, check out how easily you can open a drawing window from PowerShell:

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$window = New-Object Windows.Window
$inkCanvas = New-Object Windows.Controls.InkCanvas

$window.Title = "Scribble Pad"
$window.Content = $inkCanvas
$window.Width = 800
$window.Height = 600
$window.WindowStartupLocation = "CenterScreen"

$inkCanvas.MinWidth = $inkCanvas.Height = 600

$null = $window.ShowDialog()

# If you do not have a touch screen, use the mouse to draw on the canvas.

###########################################################################################################################################################################################

# Tip 68: Accessing Latest Log File Entries

# Sometimes you may just be interested in the last couple of entries in a log file. Here’s a simple yet fast way of outputting only the last x lines from a text log:

$logs = Get-Content -Path $env:windir\windowsupdate.log -ReadCount 0
$logs[-5..-1]                                                             # return only the last (newest) five lines from the windowsupdate.log file.

# It’s pretty fast because Get-Content uses –ReadCount 0, reading in even large text files very fast. The result is a text array with the text lines read in. 
# In the example, only the last 5 lines are output (index -5 through -1). However, $logs will hold the complete log file content which may take considerable memory. 
# So in PowerShell 3.0, there is a more efficient approach:

Get-Content -Path $env:windir\windowsupdate.log -ReadCount 0 -Tail 5      # # Show last 5 lines of windowsupdate.log file

# Here, only the number of lines specified with –Tail are returned, so there is no need for a text array to store all the other text lines found in the log file.

###########################################################################################################################################################################################

# Tip 69: Using Regions in ISE Editor

# PowerShell 3.0 ISE editor creates collapsible regions automatically, so you can collapse a function body or an IF statement. 
# In addition, you can add your own custom collapsible regions like this:


#region
Get-Process
#endregion

# Make sure you type the region comment statements exactly as shown. There may not be spaces between the comment character and the keyword, and both keywords are case-sensitive.

###########################################################################################################################################################################################

# Tip 70: Using Here-String Correctly

# Whenever you need to assign multi-line text to a variable, you can use so-called here-string. 

$myCodeSnippet = @"

function Verb-Noun
{
    param()
}

"@

# Anything in between the quotes is untouched by the PowerShell parser and can contain any character (including special characters).

# The price you pay for this is that here-string is picky. The opening tag must be at the end of a line (so add a line break right after it), 
# and the closing tag must be at the very beginning of a line. If you accidentally indent the closing tag, the here-string will break.

# This requirement makes sure that your here-string can contain almost anything without confusing the parser. 
# The only thing you may not use in a here-string is the closing tag at line position 1.

###########################################################################################################################################################################################

# Tip 71: About Powershell Snippets


# Adding Custom Snippets to ISE editor

# The new PowerShell 3.0 ISE editor features a snippet menu that lets you easily insert predefined code snippets. 
# Simply press CTRL+J to open the menu, select the snippet, and insert it into your code.

$myCodeSnippet = @"

function Verb-Noun
{
    param()
}

"@

New-IseSnippet -Title "Function body (simple)" -Description "My simple function snippet" -Author $env:USERNAME -Text $myCodeSnippet -Force

# This adds a new snippet called "Function body (simple)" to your snippet menu. 
# It is stored permanently, so next time you launch the ISE editor, it is still there. Use CTRL+J to open the snippet menu, and you will find your new custom snippet. 
# From now on, whenever you need to insert the same types of code blocks, just define a code snippet.

# These custom snippets stay permanently because PowerShell creates XML files for them. You can view these files with Get-IseSnippet:
Get-IseSnippet

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Removing ISE Snippets

Get-IseSnippet | Where-Object Name -Like "Function body*" | Remove-Item         # The snippet will stay in memory until you close the ISE editor.

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Hiding Default ISE Snippets

# The PowerShell 3.0 ISE editor ships with a number of default code snippets that you can see (and insert) by pressing CTRL+J. 
# Once you start to refine your snippets and create your own (New-IseSnippet), you may want to hide the default snippets.

$psISE.Options.ShowDefaultSnippets = $false         # This line will remove the default snippets from the snippet menu:

$psISE.Options.ShowDefaultSnippets = $true          # this line will re-enable default snippets

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Sharing and Exchanging ISE Code Snippets

# By default, the PowerShell 3.0 ISE editor loads code snippets automatically, and you can then select and insert any of these by pressing CTRL+J.

# Custom code snippets are stored in a special folder that you can open like this:
$snippetPath = Join-Path (Split-Path $profile.CurrentUserCurrentHost) "Snippets"
explorer $snippetPath

# As you see, custom snippets are stored in a folder, and each is a file with the extension ps1xml. 
# By opening this folder in your File Explorer, it is very easy to delete unwanted snippets or share some of them with friends and colleagues.

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


# Using Central ISE Snippet Repository

# Imagine you and your colleagues set up a network drive to share useful snippets. 
# Instead of manually checking this folder and manually copying the snippets from this folder onto your local machine, 
# you can also tell the ISE editor to load the snippets directly from the central repository. Here is a sample:

$snippetPath = Join-Path (Split-Path $profile.CurrentUserCurrentHost) "Snippets"
$snippetPath

Get-ChildItem -Path $snippetPath -Filter *.ps1xml | ForEach-Object {

    $psISE.CurrentPowerShellTab.Snippets.Load($_.FullName)
}

# This sample loads the snippets from your private local folder, so it shows what ISE does by default. 
# Simply assign another path to $snippetPath to load the snippets from another folder. This could be your central team repository, 
# and it could also be a USB stick you are carrying around with you that holds your favorite snippets.

# As it turns out, this is the code executed behind the scenes by Import-IseSnippet. In fact, all *-IseSnippet commands are functions, and you can view the source code like this:

$path = "$pshome\modules\ise\ise.psm1"
ise $path

###########################################################################################################################################################################################

# Tip 72: Using Default Parameter Values

# If you find yourself submitting the same value for a cmdlet parameter over and over again, then PowerShell 3.0 lets you set a default value.

# Once you defined a default value for a parameter, you no longer need to specify it. This can be very convenient in production scenarios. 
# It can be dangerous, too, because scripts that rely on this may not run on other machines anymore.

$PSDefaultParameterValues.Add("Get-ChildItem:Path", $env:TEMP)
# Next time you run Get-ChildItem (or one of its aliases such as dir) without specifying a path, you get the content of your temp folder.

# To make a default that is not limited to Get-ChildItem but applies to any cmdlet that has a Path parameter, replace the cmdlet name with a wildcard:

$PSDefaultParameterValues.Add("*:Path", $env:TEMP)

# Once you play with this, you'll soon find that defaults can be very useful but should be limited to cmdlets where they make sense. 
# Fortunately, $PSDefaultParameterValues will forget all values once you restart PowerShell, so it's easy to get rid of unwanted defaults.


# Defaults you found useful should be defined in one of your profile scripts (for example, the one specified in $profile).

###########################################################################################################################################################################################

# Tip 73: Using Help Window as Text Dialog

# Note: the code works immediately in PowerShell 3.0 ISE editor. If you want to run it in the PowerShell 3.0 console, you need to first load the necessary assembly like this:

Add-Type -AssemblyName Microsoft.PowerShell.GraphicalHost

# Here's the code (it displays text found in winrm.ini but you can display any text - just assign it to $text).

$text = Get-Content $env:windir\system32\winrm\*\winrm.ini -ReadCount 0 | Out-String

$helpWindow = New-Object Microsoft.Management.UI.HelpWindow $text -Property @{

    Title = "My Text Viewer"
    Background='#011f51'
    Foreground='#FFFFFFFF'
}

$helpWindow.ShowDialog()

###########################################################################################################################################################################################

# Tip 74: Discover New "Internet"-Cmdlets in PowerShell 3.0

# With Invoke-WebRequest and Invoke-RestMethod, PowerShell 3.0 now has powerful support for downloading information from the Internet as well as communicating with Internet services. 

# Some people use Invoke-WebRequest to easily read RSS information. For example, this line is supposed to get the contents of the PowerShell team blog RSS feed:

$rss = "http://blogs.msdn.com/b/powershell/rss.aspx"

Invoke-RestMethod $rss | Select-Object -Property Title, Link, PubDate

# As it turns out, this works, but the results are only partial. To get the complete team blog, you would have to use Invoke-WebRequest like this:

$rss = 'http://blogs.msdn.com/b/powershell/rss.aspx'

$webPage = Invoke-WebRequest $rss
$xml = [xml]$webPage.Content
$xml.rss.channel.item | Select-Object -Property Title, Link, PubDate

# So to use these cmdlets, you need to play with them and find out how they work. 
# Invoke-WebRequest always retrieves the complete content from a web page, but leaves it to you to pick the data and data type you need.

###########################################################################################################################################################################################

# Tip 75: Creating String Arrays without Quotes

# Often, you may need a list of strings and want to store it in a variable. The common coding practice is like this:

$machineType = "Native", "I386", "Itanium", "X64"


# A much easier approach does not require quotes or commas. It simply submits the strings you want in your list as parameters to Write-Output. 
# Write-Output is a cmdlet, so it does not require quotes around strings (unless they contain special characters or spaces):

$machineType = Write-Output Native I386 Itanium X64

###########################################################################################################################################################################################

# Tip 76: Identifying 32-bit Executables


# here's a small function that returns the architecture an executable was compiled for:

function Get-FileArchitecture
{
    param($filePath = "$env:windir\notepad.exe")

    $architecture = "Native, I386, Itanium, X64".Split(",")

    $data = New-Object System.Byte[] 4096
    $stream = New-Object System.IO.FileStream -ArgumentList $filePath, Open, Read
    [void]$stream.Read($data, 0, 60)

    $PE_HEADER_ADDR = [System.BitConverter]::ToInt32($data, 60)
    $architecture[[System.BitConverter]::ToUInt16($data, $PE_HEADER_ADDR + 4)]
}

Get-FileArchitecture

# Simply submit a path to an exe file, and you get back the architecture it was compiled for.


# This line would dump all non-64-bit applications from your Windows folder (which you would, of course, only run on a 64-bit system):

Get-ChildItem $env:windir -Filter *.exe -ErrorAction SilentlyContinue -Recurse | ForEach-Object {

    $arch = Get-FileArchitecture $_.FullName

    if("X64", "Native" -notcontains $arch)
    {
        $object = $_ | Select-Object -Property Name, Architecture, FullName
        $object.Architecture = $arch

        $object
    }
}

###########################################################################################################################################################################################

# Tip 77: Converting Low-Level Error Numbers into Help Messages

# Sometimes, native commands such as net.exe return cryptic error numbers instead of full error messages. 
# Traditionally, you could use the following command to convert these error numbers into full messages:

net helpmsg 3534               # Output: The service did not report an error.
net helpmsg 1                  # Output: Incorrect function.
net helpmsg 4323               # Output: The transport cannot access the medium.


# A better way may be to use winrm.exe because this command can do the very same - and more:

winrm helpmsg 3534             # Output: The service did not report an error. 
winrm helpmsg 1                # Output: Incorrect function. 
winrm helpmsg 4323             # Output: The transport cannot access the medium. 


# While net.exe can only convert a certain range of error messages, winrm.exe is more flexible and can for example also convert Remoting-specific error codes:

winrm helpmsg 0x80338104       # Output: The WS-Management service cannot process the request. The WMI service returned an 'access denied' error.

net helpmsg 0x80338104

# Output: -->> Error

#        net : The syntax of this command is:
#        At line:1 char:1
#        + net helpmsg 0x80338104
#        + ~~~~~~~~~~~~~~~~~~~~~~
#            + CategoryInfo          : NotSpecified: (The syntax of this command is::String) [], RemoteException
#            + FullyQualifiedErrorId : NativeCommandError
#         
#        NET HELPMSG
#        message#


# As you see, winrm.exe returns the correct error message whereas net.exe falls back to a standard template and cannot translate the number. 
# Winrm.exe therefore is the more generic approach that you can safely use to translate any low-level API error code to text.

###########################################################################################################################################################################################

# Tip 78: Working With TimeSpan Objects


# TimeSpan objects represent a given amount of time. They are incredibly useful when you calculate with dates or times because they can represent the amount of time between two dates, 
# or can add a day (or a minute) to a date to create relative dates.

New-TimeSpan -Days 1 -Hours 3                        # get a timespan representing one day and 3 hours

# Output:
#        Days              : 1
#        Hours             : 3
#        Minutes           : 0
#        Seconds           : 0
#        Milliseconds      : 0
#        Ticks             : 972000000000
#        TotalDays         : 1.125
#        TotalHours        : 27
#        TotalMinutes      : 1620
#        TotalSeconds      : 97200
#        TotalMilliseconds : 97200000

New-TimeSpan -End "2014-12-24 18:30:00"              # get a timespan representing the time difference between now and next Christmas


[DateTime]"2014-12-24 18:30:00" - (Get-Date)         # get a timespan by subtracting two dates
(Get-Date) - [TimeSpan]"1.00:00:00"                  # get a timespan by subtracting a timespan representing one day from a date


# negating a timespan
$timeSpan = New-TimeSpan -Days 1
$timeSpan.Negate()
$timeSpan

# creating a negative timespan directly
New-TimeSpan -Days -1

###########################################################################################################################################################################################

# Tip 79: Validating Active Directory User Account and Password


function Test-ADCredential
{
    param([System.Management.Automation.Credential()]$Credential)

    Add-Type -AssemblyName System.DirectoryServices.AccountManagement 

    $info = $Credential.GetNetworkCredential()

    if ($info.Domain -eq '') { $info.Domain = $env:USERDOMAIN }

    $TypeDomain = [System.DirectoryServices.AccountManagement.ContextType]::Domain

    try
    {
        $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $TypeDomain,$info.Domain
        $pc.ValidateCredentials($info.UserName,$info.Password)
    }
    catch
    {
        Write-Warning "Unable to contact domain '$($info.Domain)'. Original error: $_"
    }
}

# Simply submit a credential object or a string in the format "domain\username".

Test-ADCredential -Credential "domain\username"

###########################################################################################################################################################################################

# Tip 80: Resolving URLs

# Sometimes you may stumble across URLs like this one: http://go.microsoft.com/fwlink/?LinkID=13517

# As it turns out, these are just "pointers" to the real web address. 
# In PowerShell 3.0, the new cmdlet Invoke-WebRequest can resolve these URLs and return the real address that it points to:

$urlRaw = 'http://go.microsoft.com/fwlink/?LinkID=135173'
(Invoke-WebRequest -Uri $urlRaw -MaximumRedirection 0 -ErrorAction Ignore).Headers.Location

# Output: http://technet.microsoft.com/library/hh847743.aspx

###########################################################################################################################################################################################

# Tip 81: Displaying IPv4 address as Binary

$ipV4 = "192.168.0.1"

-join ($ipV4.Split(".") | ForEach-Object {[System.Convert]::ToString($_, 2).PadLeft(8, "0")})     # Output: 11000000101010000000000000000001


# You can also change the code a little:

$ipV4.Split(".") | ForEach-Object {

    "{0,5} : {1}" -f $_, [System.Convert]::ToString($_, 2).PadLeft(8, "0")
}

# Output:
#        192 : 11000000
#        168 : 10101000
#          0 : 00000000
#          1 : 00000001

###########################################################################################################################################################################################

# Tip 82: Finding User Account with WMI

# WMI represents all kinds of physical and logical entities on your machine. It also has classes that represent user accounts which include both local and domain accounts.

Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:username' and Domain='$env:userdomain'"     # returns the user account of the currently logged on user

# As always, the returned object holds a lot more information when you make PowerShell display all of its properties:
Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:username' and Domain='$env:userdomain'" | Select-Object *


# The next function, for example, checks whether the currently logged on user is a local account or a domain account:

function Test-LocalUser
{   
    param(
      $UserName = $env:username,
      $Domain = $env:userdomain
    )

    (Get-WmiObject -Class Win32_UserAccount -Filter "Name='$UserName' and Domain='$Domain'")
    
    .LocalAccount    # Exception here.
}

###########################################################################################################################################################################################

# Tip 83: Finding User Group Memberships

function Get-GroupMemberShip
{
    param($userName = $env:USERNAME, $domain = $env:userdomain)

    $user = Get-WmiObject -Class Win32_UserAccount -Filter "Name='$userName' and Domain='$domain'"
    $user.GetRelated("Win32_Group")
}

# By default, it returns the group memberships of the account that runs the script but you can use the parameters -UserName and -Domain to also specify a different account.

# If you want to access a local account on a different machine, then add the parameter -ComputerName to Get-WmiObject. 
# If you want to use PowerShell remoting rather than DCOM to remotely connect to another machine, 
# you may want to use the new CIM cmdlets instead (Get-CimInstance instead of Get-WmiObject).

###########################################################################################################################################################################################

# Tip 84: Converting Binary Data to IP Address (and vice versa)

$ipV4 = "192.168.0.1"
[Convert]::ToString(([IPAddress][String]([IPAddress]$ipV4).Address).Address, 2)             # Output: 11000000101010000000000000000001


# turn a binary into an IP address
$IPBinary = "11000000101010000000000000000001"
([System.Net.IPAddress]"$([System.Convert]::ToInt64($IPBinary, 2))").IPAddressToString      # Output: 192.168.0.1

###########################################################################################################################################################################################

# Tip 85: Prompt for Credentials without a Dialog Box

# Whenever a PowerShell script asks for credentials, PowerShell pops up a dialog box. You can view this by running this command:

Get-Credential

# PowerShell is a console-based scripting language, and so it may be unwanted to open additional dialogs. 
# That's why you can change the basic behavior and ask PowerShell to accept credential information right inside the console. 
# This is a per-machine setting, so you need local Administrator privileges and must run these two lines from an elevated PowerShell console:

$key = "HKLM:\SOFTWARE\Microsoft\PowerShell\1\ShellIds"
Set-ItemProperty -Path $key -Name ConsolePrompting -Value $true

###########################################################################################################################################################################################

# Tip 86: Running Portions of Code in 32-bit or 64-bit

# To execute code in 32-bit from within a 64-bit environment (or vice versa), you can create appropriate aliases:

# In a 32-bit PowerShell, you create:
Set-Alias Start-PowerShell64 "$env:windir\sysnative\WindowsPowerShell\v1.0\powershell.exe"

# And in a 64-bit PowerShell, you would create:
Set-Alias Start-PowerShell32 "$env:windir\syswow64\WindowsPowerShell\v1.0\powershell.exe"

# The next example runs in a 64-bit PowerShell (as a proof, pointer sizes are 8). 
# When you run the code in a 32-bit PowerShell, you get a pointer size of 4 which again proves that your code indeed is running in a 32-bit environment:

[IntPtr]::Size                          # Output: 8

Start-PowerShell32 { [IntPtr]::Size }   # Output: 4

# Note that the alias will return rich (serialized) objects that you can process as usual:

Start-PowerShell32 { Get-Service } | Select-Object -Property DisplayName, Status

###########################################################################################################################################################################################

# Tip 87: Convert IP address to decimal value (and back)

# Sometimes you may want to convert an IP address to its decimal value, for example, because you want to use binary operators to set bits. 
# Here are two simple filters that make this a snap:

filter Convert-IPtoDecimal
{
    ([IPAddress][String]([IPAddress]$_)).Address
}

filter Convert-DecimaltoIP
{
    ([System.Net.IPAddress]$_).IPAddressToString
}

"192.168.0.1" | Convert-IPtoDecimal                # Output: 16820416
16820416 | Convert-DecimaltoIP                     # Output: 192.168.0.1


###########################################################################################################################################################################################

# Tip 88: Calculate Broadcast Address

function Get-BroadcastAddress
{
    param([Parameter(Mandatory = $true)]$IPAddress, $subNetMask = "255.255.255.0")

    filter Convert-IPtoDecimal
    {
        ([IPAddress][String]([IPAddress]$_)).Address
    }

    filter Convert-DecimaltoIP
    {
        ([System.Net.IPAddress]$_).IPAddressToString
    }

    [UInt32]$ip = $IPAddress | Convert-IPtoDecimal
    [UInt32]$subnet = $subNetMask | Convert-IPtoDecimal
    [Uint32]$broadcast = $ip -band $subnet

    $broadcast -bor -bnot $subnet | Convert-DecimaltoIP
}

# This function demonstrates how you calculate network information by bitwise manipulation. 
# The broadcast address, for example, is calculated by using the bits from the subnet and setting all remaining bits to "1".

Get-BroadcastAddress -IPAddress "192.168.0.1"          # Output: 192.168.0.255
Get-BroadcastAddress -IPAddress "127.0.0.1"            # Output: 127.0.0.255

###########################################################################################################################################################################################

# Tip 89: Converting String Array in String

# When you use Get-Content to read the content of a text file, you always get back a string array. So each line of your text file is kept separately.

# If you want to convert a string array into one large string, use the operator -join:

$file = "$env:windir\WindowsUpdate.log"
$text = (Get-Content -Path $file -ReadCount 0) -join "`n"

$text = Get-Content -Path $file -Raw                               # In PowerShell 3.0, you can do the same with the new parameter -Raw

###########################################################################################################################################################################################

# Tip 90: Quickly Replace Words in Text File

$oldfile = "$env:windir\WindowsUpdate.log"
$newfile = "$env:temp\newfile.txt"

$text = (Get-Content -Path $oldfile -ReadCount 0) -join "`n"       # In Powershell 3.0:  $text = Get-Content -Path $newfile -Raw
$text -replace "error", "ALERT" | Set-Content -Path $newfile
Invoke-Item -Path $newfile

###########################################################################################################################################################################################

# Tip 91: Writing Text Information Fast

# If you want to write plain text information to a file, don't use Out-File. Instead, use Set-Content. It is much faster:

$tempfile1 = "$env:temp\tempfile1.txt"
$tempfile2 = "$env:temp\tempfile2.txt"
$tempfile3 = "$env:temp\tempfile3.txt"

$text = Get-Content -Path C:\Windows\WindowsUpdate.log

Measure-Command { $text | Out-File $tempfile1 }                                   # TotalSeconds      : 2.4570528
Measure-Command { $text | Set-Content -Path $tempfile2 -Encoding Unicode }        # TotalSeconds      : 0.77849
Measure-Command { Set-Content -Path $tempfile3 -Encoding Unicode -Value $text }   # TotalSeconds      : 0.1289937

# Depending on how large the windowsupdate.log file is on your machine, you get back varying results. 
# However, Set-Content is more than twice as fast, 
# and if you submit the text to the parameter -Value rather than sending it over the (slow) pipeline, the overall result is even 6x as fast or better.

# Use Out-File only if you want to write objects to file. The primary reason why Out-File is slower is: it tries and converts objects to plain text. 
# Set-Content writes the text as-is - which is all you need if you wanted to write plain text in the first place.

###########################################################################################################################################################################################

# Tip 92: Testing Event Log Names and Sources

# Write-EventLog lets you write custom entries to event logs, and New-EventLog can add new event logs and event sources. 
# Which raises the question: how can you test in advance whether a given event log or event log source exists?

[System.Diagnostics.EventLog]::Exists("Application")     # check if event log name exists

[System.Diagnostics.EventLog]::SourceExists("DCOM")      # check whether source exists

# Note that SourceExists() may raise an exception if you do not have Administrator privileges because then it cannot search all event logs.

###########################################################################################################################################################################################

# Tip 93: Generating Random Passwords

# If you need a simple way of creating random passwords, then this piece of code may help you:

Add-Type -AssemblyName System.Web
[System.Web.Security.Membership]::GeneratePassword(10,4)

# GeneratePassword() takes two arguments. The first sets the length of the password. The second sets the number of non-alphanumeric characters that you want in it.

###########################################################################################################################################################################################

# Tip 94: Turning CSV-Files into "Databases"

# Let's assume you have a CSV file with information that you need to frequently look up. 
# For example, the CSV file may contain server names and certain configuration settings for them.

# To easily look up items in your CSV file, you can turn it into a hash table. Let's first create a test CSV file to play with:

$file = "$env:temp\testfile.csv"

$content = @"
Servername, Year, ID, Description, Metric
Server12,2012,100,'test',99
Server98,2011,187,'production',61
S_EXCH1,2010,877,'mail',98
MEMS77,2011,300,'data',7
"@

# 以上内容的设置应尽量避免换行，换行会对后面的结果产生影响，从而使输出信息非预期所示

$content | Set-Content -Path $file              # create test CSV file

# turn this CSV file into a lookup table, using the column "Servername" as key column
$content = Import-Csv $file -Encoding UTF8
$lookup = $content | Group-Object -AsHashTable -AsString -Property Servername

$lookup.Keys                                    # listing CSV file keys
# Output:
#        Name                           Value                                                                                                                                                                                 
#        ----                           -----                                                                                                                                                                                 
#        Server98                       {@{Servername=Server98; Year=2011; ID=187; Description='production'; Metric=61}}                                                                                                      
#        Server12                       {@{Servername=Server12; Year=2012; ID=100; Description='test'; Metric=99}}                                                                                                            
#        MEMS77                         {@{Servername=MEMS77; Year=2011; ID=300; Description='data'; Metric=7}}                                                                                                               
#        S_EXCH1                        {@{Servername=S_EXCH1; Year=2010; ID=877; Description='mail'; Metric=98}} 

# Note how the code uses Group-Object to create the lookup table. Note also that its parameter -Property determines the CSV file column it uses to index the information. 
# You just need to make sure that the information in this column is unique (has no duplicate entries).

$lookup["Server98"]
$lookup["Server98"].Description                 # In PowerShell 2.0, try this:  $($lookup["Server98"]).Description

$lookup.Keys -contains "Server12"

###########################################################################################################################################################################################

# Tip 95: Counting Number of Files - Fast!

# Method 1: using plain cmdlets and determines the number of files in the Windows folder
Get-ChildItem -Path $env:windir -force | Where-Object { $_.PSIsContainer -eq $false } | Measure-Object | Select-Object -ExpandProperty Count     # Output: 45
(Get-ChildItem -Path $env:windir -File).Count

$str1 = Get-ChildItem -Path $env:windir -File -Force | Select-Object -ExpandProperty Name | Sort-Object
$str2 = Get-ChildItem -Path $env:windir -force | Where-Object { $_.PSIsContainer -eq $false } | Select-Object -ExpandProperty Name | Sort-Object

$str1.Count  # Output: 44
$str2.Count  # Output: 45

foreach($s2 in $str2)
{
    if($str1 -notcontains $s2)
    {
        $s2
    }
}

# Method 2: uses .NET methods and is shorter and is about 20x as fast
[System.IO.Directory]::GetFiles($env:windir).Count

# And here's the code to count the number of files recursively, including all subfolders:

Get-ChildItem -Path $env:windir -Force -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer -eq $false } | Measure-Object | Select-Object -ExpandProperty Count

[System.IO.Directory]::GetFiles($env:windir, "*", "AllDirectories").Count

# Again, the .NET approach is much faster, but it has one severe disadvantage. Whenever it comes across a file or folder it cannot access, the entire operation breaks. 
# The cmdlets are smarter. Get-ChildItem and its parameter -ErrorAction can ignore errors and continue with the remaining files.

###########################################################################################################################################################################################

# Tip 96: Creating New Objects - Alternative

# There are many ways in PowerShell to create new objects. One of the more advanced approaches uses a hash table to determine object properties and their values:

$content = @{

    User = $env:USERNAME
    OS = (Get-WmiObject Win32_OperatingSystem).Caption
    BIOS = (Get-WmiObject Win32_BIOS).Version
    ID = 12
}

$object = New-Object -TypeName PSObject -Property $content
$object
# Output:
#        BIOS                        User                      OS                                              ID
#        ----                        ----                      --                                              --
#        HPQOEM - 20090825           m-sihe                    Microsoft Windows 7 Enterprise                  12



# As you may note, the order of columns isn't the order you specified in your hash table, though. That's because hash tables are unordered by default. 

# In PowerShell 2.0, to order columns, you would have to post-process your object with Select-Object like this:
$object | Select-Object User, OS, BIOS, ID

# In PowerShell 3.0, you can use this alternative:

$content = [Ordered]@{

    User = $env:USERNAME
    OS = (Get-WmiObject Win32_OperatingSystem).Caption
    BIOS = (Get-WmiObject Win32_BIOS).Version
    ID = 12
}

# This creates an ordered hash table which preserves the order of columns. Note however that once you use this technique, your code is no longer compatible with PowerShell 2.0

###########################################################################################################################################################################################

# Tip 97: Running Portions of Code Elevated

# Let's assume your script may or may not need to do a privileged operation, for example write a value to a HKEY_LOCAL_MACHINE branch, depending on some prerequisites. 
# Instead of having to run the entire script with Administrator privileges, you can run only those parts requiring it:

function Start-ElevatedCode
{
    param([ScripBlock]$code)
    
    Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList $code
}

# This will launch a new elevated PowerShell console, and runs the code you submit. Of course, the user has to click the UAC dialog or submit user credentials to run it.

Start-ElevatedCode { New-Item HKLM:\SOFTWARE\Test }   # In this example, you'd create a HKLM reg key which you can only do elevated.

###########################################################################################################################################################################################

# Tip 98: Reversing Text Strings


# Method 1: Reverse() method is very useful to reverse the order of any array
$text = "Hello Silence!"
$text = $text.ToCharArray()
[Array]::Reverse($text)

-join $text                                       # Output: !ecneliS olleH

# The Reverse() method is very useful to reverse the order of any array. Since texts are just character arrays, it can be used to reverse texts. 
# The resulting character array then just needs to be converted back to a string using the operator -join.



# Method 2: let PowerShell walk the array backwards for you
$text = "Hello Silence!"
-join $text[-$text.Length..-1]                    # Output: !ecneliS olleH



# Method 3: use Regex to do the task
$text = "Hello Silence!"
-join [Regex]::Matches($text,".","RightToLeft")   # Output: !ecneliS olleH

###########################################################################################################################################################################################

# Tip 99: Examining Scheduled Tasks


# There is a COM interface that you can use to select and dump any scheduled task definition. Just make sure you are running PowerShell with full Administrator privileges.

# This example dumps all tasks in Task Scheduler Library\Microsoft\Windows\DiskDiagnostic:

$service = New-Object -ComObject Schedule.Service
$service.Connect($env:COMPUTERNAME)

$folder = $service.GetFolder("\Microsoft\Windows\DiskDiagnostic")
$tasks = $folder.GetTasks(1)

$count = $tasks.Count                                                                   # Number of tasks in that container:
"There are $count tasks."

$tasks | Select-Object -Property Name,Enabled,LastRunTime,LastTaskResult,NextRunTime    # Task statistics

# Output:
#        Name           : Microsoft-Windows-DiskDiagnosticDataCollector
#        Enabled        : False
#        LastRunTime    : 2012/11/12 09:10:17
#        LastTaskResult : 0
#        NextRunTime    : 2014/11/23 01:00:00
#        
#        Name           : Microsoft-Windows-DiskDiagnosticResolver
#        Enabled        : False
#        LastRunTime    : 1899/12/30 00:00:00
#        LastTaskResult : 1
#        NextRunTime    : 1899/12/30 00:00:00

# You can easily find out how many tasks are defined in a given container, and whether tasks are enabled, their last and their next runtime.

###########################################################################################################################################################################################

# Tip 100: Changing Scheduled Tasks with PowerShell

# PowerShell can read and also change any part of a scheduled task. Just make sure you have appropriate permissions and run PowerShell elevated.

# It picks the scheduled task "RACTask" in the scheduled task container "Microsoft\Windows\RAC". 
# It then looks at the task definition and selects the task trigger "RACTimeTrigger". 
# It finally makes sure this trigger is enabled, then writes back the updated task definition, effectively enabling this trigger:


# connect to Task Scheduler
$service = New-Object -ComObject Schedule.Service
$service.Connect($env:COMPUTERNAME)

# pick a specific task in a container:
$folder = $service.GetFolder("\Microsoft\Windows\RAC")
$task = $folder.GetTask("RACTask")

# get task definition and change it
$definition = $task.Definition
$definition.triggers | Where-Object { $_.ID -eq "RACTimeTrigger" } | ForEach-Object { $_.Enabled = $true }

# write back changed task definition:
# 4 = Update
$folder.RegisterTaskDefinition($task.Name, $definition, 4, $null, $null, $null)

# RegisterTaskDefinition() returns the task definition it just updated, so you can double-check whether your changes were correct. 
# You can also examine the task definition to find additional settings you may want to automatically change and add to your code.

# Beginning with Windows 8 / Server 2012, there is finally a module devoted to scheduled task management. This is how you can list the cmdlets available:

Get-Command -Module ScheduledTasks        # The new module "ScheduledTasks" is great but won't help you if you run scripts on older Windows versions.

###########################################################################################################################################################################################