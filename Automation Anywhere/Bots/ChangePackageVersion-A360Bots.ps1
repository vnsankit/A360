Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

$ZipFileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    Filter = 'Compressed (zipped) Folder (*.zip)|*.zip'
    Title = 'Please select the file that you exported from the A360 Control Room'
}
$Result = $ZipFileBrowser.ShowDialog()

if (($Result -ne 'OK') -OR (($ZipFileBrowser.FileName).Split(".")[-1] -ne 'zip'))
{
    [System.Windows.MessageBox]::Show('No bots file selected, please try again.','Error','Ok','Error')
    EXIT
}

$CsvFileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    Filter = 'CSV (Comma delimited) (*.csv)|*.csv'
    Title = 'Please select the file that contains package details to be updated'
}
$Result = $CsvFileBrowser.ShowDialog()

if (($Result -ne 'OK') -OR (($CsvFileBrowser.FileName).Split(".")[-1] -ne 'csv'))
{
    [System.Windows.MessageBox]::Show('No Package details file selected, please try again.','Error','Ok','Error')
    EXIT
}

Function Write-Log {
    [CmdletBinding()]
    Param(
        [String]$File,
        [string]$PackageName,
        [string]$OldVersion,
        [string]$PackageVersion,
        [string]$Status,
        [string]$Comment,
        [string]$LogFilePath
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss.fff")
   
    $Content = [PSCustomObject]@{
        DateTime = (Get-Date).ToString()
        File = $File
        'Package Name' = $PackageName
        'Old Version' = $OldVersion
        'New Version' = $PackageVersion
        'Status' = $Status
        'Comment' = $Comment
    }
    If ($null -ne $LogFilePath) {
        try {
            Export-Csv -InputObject $Content -Path $LogFilePath -NoTypeInformation -Append;
            
        }
        catch {
            [System.Windows.MessageBox]::Show($_.Exception.Message ,'Error','Ok','Error')
            EXIT          
        }
    }
    Else {
        [System.Windows.MessageBox]::Show($Content ,'Error','Ok','Error')
        EXIT
    }
} 

$ZipFile = $ZipFileBrowser.FileName
$LogFilePath = $ZipFile.Replace(".zip", "_"+(Get-Date).toString("yyyyMMdd")+".csv")
$CsvFilePath = $CsvFileBrowser.FileName

try
{
    Expand-Archive $ZipFile -DestinationPath $ZipFile.Replace(".zip", "") -Force

    $botFiles = Get-ChildItem $ZipFile.Replace(".zip", "\Automation Anywhere") -Attributes !Directory -Recurse

    foreach ($file in $botFiles)
    {
        if ($file.Extension -ne '')
        {
            continue;
        }

        $UpdateRequired = 'False'

        $Content = Get-Content $file.PSPath -Force

        Import-Csv $CsvFilePath |`
        ForEach-Object {
            $PackageName = $_."PackageName"
            $PackageVersion = $_."PackageVersion"

                if ($Content -match """$PackageName"",""version"":""\d+.\d+.\d+-\d+-\d+""" -AND $PackageVersion -match "\d+.\d+.\d+-\d+-\d+")
                {
                    $Found = $Content -match """$PackageName"",""version"":""\d+.\d+.\d+-\d+-\d+"""
                    $Found = $matches[0] -match "\d+.\d+.\d+-\d+-\d+"
                    $OldVersion = $matches[0]
        
                    if ($OldVersion -ne $PackageVersion)
                    {
                        $Content = $Content -replace """$PackageName"",""version"":""\d+.\d+.\d+-\d+-\d+""", """$PackageName"",""version"":""$PackageVersion"""

            
                        if ($Content -match """$PackageName"",""version"":""$PackageVersion""")
                        {
                            $UpdateRequired = 'True'
                            Write-Log ($file.FullName).Substring(($file.FullName).IndexOf('\Bots\')) $PackageName $OldVersion $PackageVersion 'PASS' 'Package updated' $LogFilePath
                        }
                        else
                        {
                            Write-Log ($file.FullName).Substring(($file.FullName).IndexOf('\Bots\')) $PackageName $OldVersion $PackageVersion 'FAIL' 'Package update failed' $LogFilePath
             
                        }
                    }
                    else
                    {
                        Write-Log ($file.FullName).Substring(($file.FullName).IndexOf('\Bots\')) $PackageName $OldVersion $PackageVersion 'SKIP' 'No update required' $LogFilePath
             
                    }
                }
                else
                    {
                        $OldVersion = "N/A"
                        Write-Log ($file.FullName).Substring(($file.FullName).IndexOf('\Bots\')) $PackageName $OldVersion $PackageVersion 'SKIP' 'Package not used or Incorrect name/version' $LogFilePath
             
                    }
            }

        if ($UpdateRequired -eq 'True')
        { Set-Content $file.PSPath $Content }
    }

    Compress-Archive $ZipFile.Replace(".zip", "\*") -DestinationPath $ZipFile -Update
    Remove-Item -LiteralPath $ZipFile.Replace(".zip", "") -Force -Recurse

	[System.Windows.MessageBox]::Show("Package update completed, please check the updated bots and log file.`n`nBots file: "+$ZipFile+"`n`nLog file: "+$LogFilePath,'Information','Ok','Info')
}
 catch
 {
    [System.Windows.MessageBox]::Show($_.Exception.Message ,'Error','Ok','Error')
 }