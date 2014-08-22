# Call Method Example: & '.\HandleExcel - V2.ps1' 2014/07/28

#Location for the "AzureSDK 2.4 test execution.xlsx"
$file = "C:\Users\v-sihe\Desktop\Temp\HandleExcel\AzureSDK 2.4 test execution.xlsx"

# Get Excel info
Get-Item -Path $file | select name,lastwritetime

Write-Host
Write-Host "Warning: Please make sure the task scenario excel is the latest! (Y/N)" -ForegroundColor Yellow
$isLatest = Read-Host

Write-Host "AzureSDK 2.4 test execution.xlsx will be open, please get date and total rows from this file!"
start $file

Write-Host "Please input the date according to the excel file: (like 2014/07/28)"
$ETATime = Read-Host
Write-Host "Please input int data to custom how many rows need to filter with the excel file: "
$endRow  = Read-Host

if($ETATime -eq $null -or $endRow -eq $null)
{
    Write-Host "Parameter is not completed, exit!"
    break
}

function initProcess
{
    $isexcelExist = @((Get-Process | Select-Object processname).processname).Contains("excel")
    if($isexcelExist)
    {
        # Stop all excel process before start collect data
        Get-Process excel | Select-Object id | ForEach-Object { Stop-Process -Id $_.Id -Force}
    }

    
    # Open and read "AzureSDK 2.4 test execution.xlsx"[excel], then filter the right data to a temp excel[excel2]
    # The var marked as global: because these var need to be visible for other function
    $global:excel =  New-Object -ComObject Excel.Application
    
    $global:excel2 = New-Object -ComObject Excel.Application
    $excel2.Visible = $true
    $workbook = $excel2.Workbooks.add()
    $global:sheet2 = $workbook.worksheets.Item(1)
    
    #Set col title
    $sheet2.cells.item(1,1) = "Scenario"
    $sheet2.cells.item(1,2) = "Category"
    $sheet2.cells.item(1,3) = "ETA"
    $sheet2.cells.item(1,4) = "Recommend OS"
    $sheet2.cells.item(1,5) = "Machine"
    $sheet2.cells.item(1,6) = "Status"
    $sheet2.cells.item(1,7) = "Issues"

    #Set font
    for($i = 1;$i -le 7;$i++)
    {
        $sheet2.cells.item(1,$i).font.bold = $true
    }
}


function cleanUP
{
    #This script should lanch two excel here, kill the first one according to the start time
    Stop-Process -id ((get-process excel | select id,starttime | Sort-Object -Descending)[0].id) -Force
}


$Global:temp = 2
function collectData($tabID = 1)
{
    # Fiter data from "AzureSDK 2.4 test execution.xlsx"
    $book = $null
    $s


    $book = $excel.Workbooks.Open($file)
    $sheet = $book.Worksheets.Item($tabID)
    $cells = $sheet.Cells

    $row = 3
    $vsSKU = $null # VSSKU: VS2012/VS2013
    $count = $Global:temp  # $count marked the row to insert right data

    
    if($tabID -eq 1)     #$tabID = 1: "Dev2013 Azure bundle 2.4"
    {
        $vsSKU = "VS2013"
    }
    elseif($tabID -eq 2) #$tabID = 2: "Dev2012 Azure bundle 2.4"
    {
        $vsSKU = "VS2012"
    }

    Write-Host
    Write-Host "                          Branch: $($vsSKU) Azure bundle 2.4                                           " -BackgroundColor Cyan -ForegroundColor Black
    
    $flag = $true
    $regex = [regex]"^(\d+\/\d+\/\d+)"  # this regex try to get the right date from excel: like 2014/07/28(OGF) 
    $mark = 0                           # give a friendly message format and let the line info wirte to a new line each 20 counts

    while($flag)
    {
        #Scenario
        $title1  =  $cells.Item($row,2).Value2
        #Category
        $title2  =  $cells.Item($row,3).Value2
        
        #ETA/Completed
        $latest = "1900/1/1"
    
        for($i = 4; $i -le 12; $i++)
        {
            $date = $null
            $value = $cells.Item($row,$i).Value2
            
            if($value -eq $null)
            {
                continue
            }
    
            if($value.GetType().FullName -eq "System.Double")
            {
                #$date = [DateTime]::FromOADate([System.Convert]::ToDouble($value)).ToString("yyyy/MM/dd")
                $date = [DateTime]::FromOADate($value).ToString("yyyy/MM/dd")
            }
            elseif($value.GetType().FullName -eq "System.String")
            {
                #$regex = [regex]"^(\d+\/\d+\/\d+)"
                $date  = $regex.Matches(($value)).Value
            }
    
            $time1 = [DateTime]$latest
            $time2 = [DateTime]$date
    
            if(($time1 - $time2).TotalMilliseconds -ge 0)
            {
                $latest = $latest
            }
            else
            {
                $latest = $date
            }
        }
    
        # Filter the test scenario according to the input date
        if(($latest).ToString() -ne ($ETATime).ToString())
        {
            Write-Host $row " ⇨ " -NoNewline -ForegroundColor Yellow
    
            $row++
            $mark++

            # when $mark = 20 write a new line
            # For Example:
             #3  ⇨ 54  ⇨ 55  ⇨ 56  ⇨ 57  ⇨ 58  ⇨ 59  ⇨ 60  ⇨ 61  ⇨ 62  ⇨ 63  ⇨ 64  ⇨ 65  ⇨ 66  ⇨ 67  ⇨ 68  ⇨ 69  ⇨ 70  ⇨ 71  ⇨ 72  ⇨ 
             #73  ⇨ 74  ⇨ 75  ⇨ 76  ⇨ 77  ⇨ 78  ⇨ 79  ⇨ 80  ⇨ 81  ⇨ 82  ⇨ 83  ⇨ 84  ⇨ 85  ⇨ 86  ⇨ 87  ⇨ 88  ⇨ 89  ⇨ 90  ⇨ 91  ⇨ 92  ⇨ 

            if($mark -eq 20)
            {
                Write-Host
                $mark = 0
            }

    
            if($row -eq $endRow)
            {
                Write-Host
                Write-Host "Done." -ForegroundColor Green
    
                $flag = $false
                $count += 2
                $Global:temp = $count
            }
            continue
        }
    
        #Recommend OS
        $title12 =  $cells.Item($row,13).Value2
        
        Write-Host
        Write-Host "Line:          " $row -ForegroundColor Yellow
        Write-Host "Scenario:      " $title1 -ForegroundColor Green
        Write-Host "Category:      " $title2
        Write-Host "ETA:           " $latest
        Write-Host "Recommend OS:  " $title12
        Write-Host "======================================================================================================="
    
        $sheet2.cells.item($count,1) = $title1
        $sheet2.cells.item($count,2) = $title2
        $sheet2.cells.item($count,3) = $latest
        $sheet2.cells.item($count,4) = $title12
        $sheet2.cells.item($count,5) = "IIS-SBH"
        $sheet2.usedRange.EntireColumn.AutoFit() | Out-Null
    
        $row++
        $count++
        $mark = 0  #Reset mark value
    
        if($row -eq $endRow)
        {
            Write-Host
            Write-Host "Done." -ForegroundColor Green
    
            $flag = $false
            $count += 2
            $Global:temp = $count
        }
    }
}

trap
{
    Write-Host "Exception here." -ForegroundColor Red
    Get-Process excel | Select-Object id | ForEach-Object { Stop-Process -Id $_.Id -Force}  # Stop all excel process
}

if($isLatest.ToUpper() -eq "Y")
{
    initProcess

    1..2 | ForEach-Object { collectData -tabID $_ }
     
    cleanUP
}
elseif($isLatest.ToUpper() -eq "N")
{
    Write-Host "The test scenario excel is not the latest, please get the right one!" -ForegroundColor Yellow
    start \\bpdfiles01\CommonShare\xuezhain\AzureOneSDK\DotNET
}
else
{
    Write-Host "Invalid input, please try again!" -ForegroundColor Yellow
}

