
function getTextDataFromClipboard()
{
    $data = [System.Windows.Forms.Clipboard]::GetDataObject()

    if($data.GetDataPresent([System.Windows.Forms.DataFormats]::Text))
    {
        $txtData = $data.GetData([System.Windows.Forms.DataFormats]::Text).ToString()
        Write-Host "Ctrl + C: --->> " -ForegroundColor Yellow -NoNewline
        Write-Host $txtData -BackgroundColor Magenta -NoNewline
        
        return $txtData
    }
    else
    {
        [System.Windows.Forms.MessageBox]::Show("The data in Clipboard can't be translated to Text!", "Error")
    }
}

function testPath()
{
    $filePath = getTextDataFromClipboard

    if(Test-Path $filePath -ErrorAction Ignore)
    {
        Write-Host "  The path is valid! " -ForegroundColor Green
    }
    else
    {
        Write-Host "  The path is invalid! " -ForegroundColor Red
    }
}

testPath




while($true)
{
    Start-Sleep -Seconds 1
    testPath
}


[System.Windows.Forms.Clipboard]::GetDataObject()
[System.Windows.Forms.Clipboard]::Clear()
[System.Windows.Forms.Clipboard]::ContainsText()