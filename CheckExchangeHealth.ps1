Function Add-Zip
{ 
    Param([String]$SourceFile,[String]$ZipFile) 
      
    $file = Get-ChildItem $SourceFile
    if (!(Test-Path $ZipFile)) {
        Set-Content $ZipFile ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
    }
    $ShellApplication = New-Object -Com Shell.Application
    $ZipPackage = $ShellApplication.NameSpace($ZipFile)
    $ZipPackage.CopyHere(($file.FullName))
}

Function Test-FileLockStatus
{
    Param ([String]$FilePath)
    $FileLocked = $False
    $FileInfo = New-Object System.IO.FileInfo $FilePath
    Trap {
        Set-Variable -name Filelocked -value $True -scope 1
        Continue
    	}
    $FileStream = $FileInfo.Open( [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None )
    if ($FileStream) {$FileStream.Close()}
    $FileLocked
}

#Set Path
$DirectoryPath = "D:\Scripts\"
$ReportFile = "$($DirectoryPath)ExchangeAllServersReport.html"
$Logfile = "$($DirectoryPath)Logs\CheckExchangeHealth_$((get-date).year).txt"

cd $DirectoryPath
Add-PSSnapin Microsoft.Exchange*

#Update HealthChecker.ps1
#.\HealthChecker.ps1 -ScriptUpdateOnly

#Clear Report
try {
    $Reportfiles = Get-ChildItem HealthChecker-*.xml,HealthChecker-*.txt,ExchangeAllServersReport.html
    foreach ($file in $Reportfiles) {
        $ReportFile | Remove-Item
        Add-Content "0,$(get-date -f 'yyyyMMdd-HH:ss'),RemoveFile,$($file.name)" -Path $Logfile
    } 
}
catch {
}


#Main
try {
    #Run MS HealthChecker.ps1
    Get-ExchangeServer | ?{$_.AdminDisplayVersion -Match "^Version 15"} | %{.\HealthChecker.ps1 -Server $_.Name}; .\HealthChecker.ps1 -BuildHtmlServersReport
    try {
        #Report Manager
        if ((Test-Path $ReportFile) -and (Get-Item $ReportFile).LastWriteTime.Date -eq (Get-Date).Date) {
            #Send Report to Admins
            $Report =  Get-Content $ReportFile -Raw
            $smtpServer = ""
            $smtpFrom = "MessagerBot@henlius.com"
            $smtpTo = ""
            #$smtpCC =
            #$smtpBCC = 
            $messageSubject = "Exchange Servers Health Report $(get-date -f 'yyyy-MMdd')"
            $message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
            #$message.CC.Add($smtpcc)
            #$message.BCC.Add($smtpbcc)
            $message.Subject = $messageSubject
            $message.IsBodyHTML = $true
            $message.Body = $Report
            #$message.Attachments.Add($reportfile)
            $smtp = New-Object Net.Mail.SmtpClient($smtpServer, 587)
            $smtp.EnableSsl = $true
            $smtp.Credentials = New-Object System.Net.NetworkCredential("", "");
            $smtp.Send($message)
            Add-Content "0,$(get-date -f 'yyyyMMdd-HH:ss'),SendMail,$smtpTo" -Path $Logfile

            #Zip
            $CheckFiles = Get-ChildItem HealthChecker-*.xml,HealthChecker-*.txt,ExchangeAllServersReport.html
            $ZIPTargetFiles = $DirectoryPath+'Backup\'+"$(get-date -f 'yyyyMMddHHss').zip"
            foreach ($file in $CheckFiles) {
                Add-Zip -SourceFile $file.FullName -ZipFile $ZIPTargetFiles
                Do {
                    Start-Sleep -Seconds 5
                    $FileLockStatus = Test-FileLockStatus -FilePath $ZIPTargetFiles
                } 
                While ($FileLockStatus -eq $True)
                Remove-Item -Path $file.FullName -Confirm:$false
            }
            Add-Content "0,$(get-date -f 'yyyyMMdd-HH:ss'),ZipFile,$ZIPTargetFiles" -Path $Logfile
        }
    }
    catch {
        $Exception = $(($Error[0].Exception | Out-String).Trim())
        Add-Content "1,$(get-date -f 'yyyyMMdd-HH:ss'),$Exception," -Path $Logfile
        #Clear Report
        try {
            $CheckFiles = Get-ChildItem HealthChecker-*.xml,HealthChecker-*.txt,ExchangeAllServersReport.html
            foreach ($file in $CheckFiles) {
                $file | Remove-Item
                Add-Content "0,$(get-date -f 'yyyyMMdd-HH:ss'),RemoveFile,$($file.name)" -Path $Logfile
            } 
        }
        catch {
        }
    }

}
catch {
    $Exception = $(($Error[0].Exception | Out-String).Trim())
    Add-Content "1,$(get-date -f 'yyyyMMdd-HH:ss'),$Exception," -Path $Logfile
}
