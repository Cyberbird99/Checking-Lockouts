Import-Module C:\Github Projects\Powershell\SendEmail\MailModule.psm1
$MailAccount=Import-Clixml -Path C:\Github Projects\Powershell\outlook.xml
$MailPort=587
$MailSMTPServer="smtp-mail.outlook.com
$MailFrom=$MailAccount.UserName
$MailTo="test@outlook.com"

$LogPath="C:\Github Projects\Powershell\Checking Lockouts"

$LogFile="Lockouts - $(Get-Date -Format "yyyyy-MM-dd hh-mm").csv"

$LockedOutUsers=Search-ADAccount -LockedOut -Server TestServer

$Export=[System.Collections.ArrayList]@()

foreach($LockedOutUser in $LockedOutUsers){
    $ADUser=Get-ADUser -Identity $LockedOutUser.SamAccountName -Server TestServer -Properties *
    $Entry=New-Object -TypeName psobject
    Add-Member -InputObject $Entry -MemberType NoteProperty -Name "Name" -Value "$($ADUser.GivenName) $($ADUser.Surname)"
    Add-Member -InputObject $Entry -MemberType NoteProperty -Name "UserName" -Value $ADUser.SamAccountName
    Add-Member -InputObject $Entry -MemberType NoteProperty -Name "LockoutTime" -Value $([datetime]::FromFileTime($ADUser.lockoutTime))
    Add-Member -InputObject $Entry -MemberType NoteProperty -Name "LastBadPasswordAttempt" -Value $ADUser.LastBadPasswordAttempt
    [void]$Export.Add($Entry)

}

if($Export){
    $Export | Export-Csv -Path "$LogPath\$LogFile" -Delimiter ',' -NoTypeInformation
}

if(Test-Path -path "$LogPath\$LogFile"){
    $Subject="Account Lockouts"
    $Body="the lockedout accounts"
    $Attachment="$LogPath\$LogFile"
    Send-MailKitMessage -From $MailFrom -To $MailTo -SMTPServer $MailSMTPServer -Port $MailPort -Credential $MailAccount -Subject $Subject -Body $Body -Attachments $Attachment 
}