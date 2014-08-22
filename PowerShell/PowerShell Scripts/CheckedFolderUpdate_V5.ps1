
param($path = $(throw "please supply the path to monitor."))

$seconds = 0

while($true)
{
    $exfileInfo = @{}
    $currentFileInfo = @{}

    #V5 new feature: Add -Recures to monitor all files under a folder and show fullname for the changed file
    $oriItemStatus = $null
    $oriItemStatus = @((Get-ChildItem $path -Recurse | Sort-Object -Property name))

    Start-Sleep -Seconds 1

    $currentItemStatus = $null
    $currentItemStatus = @((Get-ChildItem $path -Recurse | Sort-Object -Property name))

    for($i = 0; $i -lt $oriItemStatus.Count; $i++)
    {
        $exfileInfo[$oriItemStatus[$i].FullName] = $oriItemStatus[$i].LastWriteTime
    }

    for($i = 0; $i -lt $currentItemStatus.Count; $i++)
    {
        $currentFileInfo[$currentItemStatus[$i].FullName] = $currentItemStatus[$i].LastWriteTime
    }

    if($oriItemStatus.Count -ge $currentItemStatus.Count)
    {
        foreach($name in $oriItemStatus.FullName)
        {
            $falg = $false
            foreach($key in $currentFileInfo.Keys)
            {
                if($name -eq $key)
                {
                    $falg = $true
                }
            }
            if($falg)
            {
                if($exfileInfo[$name] -ne $currentFileInfo[$name])
                {
                    Write-Host "$($name) was updated!" -ForegroundColor Yellow
                }
            }
            if(-not $falg)
            {
                Write-Host "$($name) was deleted!" -ForegroundColor Red
            }
        }

        #Note: Add this foreach code is to fix a bug: when you create a new file and delete another file, this tool will not catch this
        foreach($name in $currentItemStatus.FullName)
        {
            $falg = $false
            foreach($key in $exfileInfo.Keys)
            {
                if($name -eq $key)
                {
                    $falg = $true
                }
            }
            if(-not $falg)
            {
                Write-Host "$($name) was created!" -ForegroundColor Green
            }
        }
    }
    else
    {
        foreach($name in $currentItemStatus.FullName)
        {
            $falg = $false
            foreach($key in $exfileInfo.Keys)
            {
                if($name -eq $key)
                {
                    $falg = $true
                }
            }
            if(-not $falg)
            {
                Write-Host "$($name) was created!" -ForegroundColor Green
            }
        }             
    } 

    $seconds++

    Write-Host "." -NoNewline

    if($seconds -eq 60)
    {
        Write-Host
        $seconds = 0
    }
}