#$src = @("http://download.microsoft.com/download/4/A/B/4ABC96F2-A611-4255-ADEE-7D2D53B9B7CF/WebDeploy_amd64_cs-CZ.msi")
$compareTools = "\\iisdist\privates\v-weissh\ComputeFileHash.exe"
$des = "D:\Test"
#$shareFolder = "\\iisdist\release\oob\wdeploy\drops\Current\1929.0\sign\msdeploy\msi\signed"
$shareFolder = "D:\Test\1"
$src =
    "http://download.microsoft.com/download/4/A/B/4ABC96F2-A611-4255-ADEE-7D2D53B9B7CF/WebDeploy_amd64_cs-CZ.msi",
    "http://download.microsoft.com/download/4/A/B/4ABC96F2-A611-4255-ADEE-7D2D53B9B7CF/WebDeploy_x86_cs-CZ.msi",
    "http://download.microsoft.com/download/4/3/F/43F7FC1A-F70C-49C2-99A1-7CE6B7CAF60F/WebDeploy_amd64_de-DE.msi",
    "http://download.microsoft.com/download/4/3/F/43F7FC1A-F70C-49C2-99A1-7CE6B7CAF60F/WebDeploy_x86_de-DE.msi",
    "http://download.microsoft.com/download/0/1/D/01DC28EA-638C-4A22-A57B-4CEF97755C6C/WebDeploy_amd64_en-US.msi",
    "http://download.microsoft.com/download/0/1/D/01DC28EA-638C-4A22-A57B-4CEF97755C6C/WebDeploy_x86_en-US.msi",
    "http://download.microsoft.com/download/8/7/2/87257D73-6306-4FD9-B47C-E971F4BD0A77/WebDeploy_amd64_es-ES.msi",
    "http://download.microsoft.com/download/8/7/2/87257D73-6306-4FD9-B47C-E971F4BD0A77/WebDeploy_x86_es-ES.msi",
    "http://download.microsoft.com/download/4/0/8/408B6C0D-585D-44E2-81ED-15536D4E321F/WebDeploy_amd64_fr-FR.msi",
    "http://download.microsoft.com/download/4/0/8/408B6C0D-585D-44E2-81ED-15536D4E321F/WebDeploy_x86_fr-FR.msi",
    "http://download.microsoft.com/download/0/D/C/0DC7AA01-0286-4E17-A610-4BE32AB656BE/WebDeploy_amd64_it-IT.msi",
    "http://download.microsoft.com/download/0/D/C/0DC7AA01-0286-4E17-A610-4BE32AB656BE/WebDeploy_x86_it-IT.msi",
    "http://download.microsoft.com/download/0/D/5/0D50B63D-EE55-4834-8312-89CDEFABDE44/WebDeploy_amd64_ja-JP.msi",
    "http://download.microsoft.com/download/0/D/5/0D50B63D-EE55-4834-8312-89CDEFABDE44/WebDeploy_x86_ja-JP.msi",
    "http://download.microsoft.com/download/3/C/4/3C40387C-E975-4A5A-B44B-963F736AA097/WebDeploy_amd64_ko-KR.msi",
    "http://download.microsoft.com/download/3/C/4/3C40387C-E975-4A5A-B44B-963F736AA097/WebDeploy_x86_ko-KR.msi",
    "http://download.microsoft.com/download/B/2/2/B22FB61D-06ED-4FD7-9AD5-F538E4569FD0/WebDeploy_amd64_pl-PL.msi",
    "http://download.microsoft.com/download/B/2/2/B22FB61D-06ED-4FD7-9AD5-F538E4569FD0/WebDeploy_x86_pl-PL.msi",
    "http://download.microsoft.com/download/C/B/D/CBD42703-6E26-4C9B-B341-2F4C84088853/WebDeploy_amd64_pt-BR.msi",
    "http://download.microsoft.com/download/C/B/D/CBD42703-6E26-4C9B-B341-2F4C84088853/WebDeploy_x86_pt-BR.msi",
    "http://download.microsoft.com/download/7/0/E/70E54287-6810-4C50-AE1F-EEDE83AF41C4/WebDeploy_amd64_ru-RU.msi",
    "http://download.microsoft.com/download/7/0/E/70E54287-6810-4C50-AE1F-EEDE83AF41C4/WebDeploy_x86_ru-RU.msi",
    "http://download.microsoft.com/download/2/A/4/2A48835C-D08F-48F0-977F-8522EF08638A/WebDeploy_amd64_tr-TR.msi",
    "http://download.microsoft.com/download/2/A/4/2A48835C-D08F-48F0-977F-8522EF08638A/WebDeploy_x86_tr-TR.msi",
    "http://download.microsoft.com/download/5/6/4/56418889-EAC9-4CE6-93C3-E0DA3D64A0D8/WebDeploy_amd64_zh-CN.msi",
    "http://download.microsoft.com/download/5/6/4/56418889-EAC9-4CE6-93C3-E0DA3D64A0D8/WebDeploy_x86_zh-CN.msi",
    "http://download.microsoft.com/download/5/4/4/5441215F-78DB-4686-AD05-3FAD23C406D1/WebDeploy_amd64_zh-TW.msi",
    "http://download.microsoft.com/download/5/4/4/5441215F-78DB-4686-AD05-3FAD23C406D1/WebDeploy_x86_zh-TW.msi"


$fileName = $null
function DownloadExeFile($path = $des)
{
    foreach($url in $src)
    {
        $fileName = $url.Split('/')[-1]
        $target = Join-Path $des $fileName

        if(!(Test-Path $target))
        {
            Invoke-WebRequest -Uri $url -OutFile $target
        }
        else
        {
            Write-Host "The file is already exists!" -ForegroundColor Red
        }
    }
}

function GetSHA1($path = $des)
{
    $fileInfo = @{}
    $process = @(cmd /c $compareTools $path)
    
    foreach($value in $process)
    {
        $key = $value.ToString().Split("`t")[-1]
        $value = $value.ToString().Split("`t")[0]
        try
        {
            $fileInfo += @{$key = $value}
        }
        catch
        {
            Write-Host "The key in HashTable already exist!" -ForegroundColor Red
        }      
    }
    return $fileInfo
}

function CompareSHA1
{
    $internetFileInfo = GetSHA1 -path $des
    $shareFolderInfo  = GetSHA1 -path $shareFolder

    foreach($key1 in $internetFileInfo.Keys)
    {
        if($shareFolderInfo.ContainsKey($key1))
        {
            if($internetFileInfo[$key1] -eq $shareFolderInfo[$key1])
            {
                Write-Host $("File: {0,25} `t SHA1: {1} `t Result: Passed" -f $key1,$internetFileInfo[$key1]) -ForegroundColor Green
            }
            else
            {
                Write-Host $("File: {0,25} `t SHA1: {1} `t SHA1: {2} `tResult: Failed" -f $key1,$internetFileInfo[$key1],$shareFolderInfo[$key1]) -ForegroundColor Red
            }
        }
    }
}
#DownloadExeFile
CompareSHA1
