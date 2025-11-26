<#
.SYNOPSIS
	App for help desk techs
.DESCRIPTION
	Creates a window for managing user accounts, and collecting user and computer info.
.EXAMPLE
	PS> ./helpdesktool.ps1
.NOTES
	Author: Brian Widom
#>

$MainWindow = .\CreateWindow.ps1 -Path '../Windows/MainWindow.xaml'
$MainWindow.Title = 'Service Desk Tool'

$dgAccountInfo = $MainWindow.FindName("dgAccountInfo")
$cbSearchCriteria = $MainWindow.FindName("cbSearchCriteria")
$tbSearchUser = $MainWindow.FindName("tbSearchUser")
$tbEmployeeID = $MainWindow.FindName("tbEmployeeID")
$tbSAMAccountName = $MainWindow.FindName("tbSAMAccountName")
$tbComputerSearch = $MainWindow.FindName("tbComputerSearch")
$tbSessions = $MainWindow.FindName("tbSessions")
$iDisabledIcon = $MainWindow.FindName('iDisabledIcon')
$lbSessions = $MainWindow.FindName("lbSessions")
$tbComputerName = $MainWindow.FindName('tbComputerName')
$tbFreeDiskSpace = $MainWindow.FindName('tbFreeDiskSpace')
$tbMemoryUsage = $MainWindow.FindName('tbMemoryUsage')
$tbLastBootTime = $MainWindow.FindName('tbLastBootTime')

$tbSearchUser.Focus() | Out-Null

$dataTable = New-Object System.Data.DataTable

[void]$dataTable.Columns.Add("DC Name", [string])
[void]$dataTable.Columns.Add("LastBadPassword", [string])
[void]$dataTable.Columns.Add("PasswordLastSet", [string])
[void]$dataTable.Columns.Add("PasswordExpirationDate", [string])
[void]$dataTable.Columns.Add("LockedOut", [string])
[void]$dataTable.Columns.Add("BadLogonCount", [int])

#Properties for Get-ADUser command
$properties = @("LastBadPasswordAttempt", "PasswordLastSet", "msDS-UserPasswordExpiryTimeComputed", "BadLogonCount", "LockedOut", "EmployeeID", "SAMAccountName", "LockoutTime")

$dgAccountInfo.ItemsSource = $dataTable.DefaultView

$dcs = @(Get-ADDomainController -Filter * | Sort-Object -Property Name)
$rows = [Object[]]::new($dcs.Count)


for($i=0; $i -lt $dcs.Count; $i++){
    $rows[$i] = $dataTable.NewRow()
    $dataTable.Rows.Add($rows[$i])
}

#Set rows on data grid. 
function Set-Rows{
    param(
        [Parameter(Position=0)]
        [int]$RowIndex,
        [Parameter(Position=1)]
        [string]$LastBadPassword=[string]::Empty,
        [Parameter(Position=2)]
        [string]$PasswordLastSet=[string]::Empty,
        [Parameter(Position=3)]
        [string]$PasswordExpirationDate=[string]::Empty,
        [Parameter(Position=4)]
        [string]$LockedOut=[string]::Empty,
        [Parameter(Position=5)]
        $BadLogonCount=[DBNull]::Value,
        [Parameter(Position=6)]
        [string]$DCName = [string]::Empty
    )
    $rows[$RowIndex]["LastBadPassword"] = $LastBadPassword
    $rows[$RowIndex]["PasswordLastSet"] = $PasswordLastSet
    $rows[$RowIndex]["PasswordExpirationDate"] = $PasswordExpirationDate
    $rows[$RowIndex]["LockedOut"] = $LockedOut
    $rows[$RowIndex]["BadLogonCount"] = $BadLogonCount
    $rows[$RowIndex]["DC Name"] = $DCName
}


function Search-User{
    $tbEmployeeID.Text = "Collecting data..."
    $tbSAMAccountName.Text = ""
    $iDisabledIcon.Visibility="Hidden"
    #Force the text box controls to update
    [System.Windows.Forms.Application]::DoEvents()
    switch($cbSearchCriteria.SelectedIndex){
        0{
            $filter = "(EmployeeID -eq '$($tbSearchUser.Text)') "
        }
        1{
            $x = "*"+$tbSearchUser.Text+"*"
            $filter = "Name -like '$x' -OR SAMAccountName -like '$x'"          
        }
    }
    
    #Get count of users who match criteria. If more than one, diplay matching users.
    $countUser = @(Get-ADUser -Filter $filter)

    if($countUser.Count -eq 1){
        for($i = 0; $i -lt $dcs.Count; $i++){             
                $userInfoOnServer = @(Get-ADUser -Server $dcs[$i] -Filter $filter -Properties $properties -SearchBase $countUser[0].DistinguishedName)                
                Set-Rows $i `
                    $(if($userInfoOnServer.LastBadPasswordAttempt){$userInfoOnServer.LastBadPasswordAttempt}else{'None'}) `
                    $(if($userInfoOnServer.PasswordLastSet){$userInfoOnServer.PasswordLastSet}else{"Change Password"}) `
                    $(if($userInfoOnServer.PasswordLastSet){[datetime]::FromFileTime($userInfoOnServer.'msDS-UserPasswordExpiryTimeComputed')}else{"N/A"}) `
                    $(if($userInfoOnServer.LockoutTime -gt 0){"Locked"}else{"Unlocked"}) `
                    $(if($userInfoOnServer.BadLogonCount){$userInfoOnServer.BadLogonCount}else{0}) `
                    $($dcs[$i].Name)
        }
        if($countUser[0].Enabled){$iDisabledIcon.Visibility='Hidden'}else{$iDisabledIcon.Visibility='Visible'}
        $tbEmployeeID.Text = $userInfoOnServer.EmployeeID
        $tbSAMAccountName.Text = $userInfoOnServer.SAMAccountName
    }elseif($countUser.Count -eq 0){
        for($i = 0; $i -lt $dcs.Count; $i++){       
            Set-Rows -RowIndex $i      
        }
        throw "User not found"
        $tbEmployeeID.Text = ""
        $tbSAMAccountName.Text = ""
    }elseif($countUser.Count -gt 1){
        $tbEmployeeID.Text = ""
        $tbSAMAccountName.Text = ""
        $iDisabledIcon.Visibility="Hidden"
        Create-SelectUserWindow
    }
}

function Unlock-User{
    #Check if a user is selected.
    if($tbSAMAccountName.Text){
        foreach($dc in $dcs){
                Unlock-ADAccount -Identity $tbSAMAccountName.Text -Server $dc
        }        
    }else{
        throw "No User Selected"
    }
}

function Create-PasswordWindow{
    if($tbSAMAccountName.Text){
        $passwordSetting = Get-Content -Path '..\config.json' | ConvertFrom-Json

        $ChangePasswordWindow = .\CreateWindow.ps1 -Path '..\Windows\ChangePasswordWindow.xaml'
        $ChangePasswordWindow.Owner = $MainWindow
        $ChangePasswordWindow.WindowStartupLocation = 'CenterOwner'      

        $lChangePasswordPrompt = $ChangePasswordWindow.FindName("lChangePasswordPrompt")
        $lChangePasswordPrompt.Content = "Change " + $tbSAMAccountName.Text + "'s password to:"

        $attribute = Get-ADUser -Identity $tbSAMAccountName.Text -Properties $passwordSetting.userAttribute | Select-Object -ExpandProperty $passwordSetting.userAttribute
        if($passwordSetting.partOfAttribute -gt 0){
            $password = $passwordSetting.passwordText + $attribute.Substring($passwordSetting.partOfAttribute - 1)
        }elseif($passwordSetting.partOfAttribute -lt 0){
            $password = $passwordSetting.passwordText + $attribute.Substring($attribute.Length - ($passwordSetting.partOfAttribute*-1))
        }else{
            $password = $passwordSetting.passwordText
        }

        $tbNewPassword = $ChangePasswordWindow.FindName("tbNewPassword")
        $tbNewPassword.Text = $password

        function Change-UserPassword{
            Write-Host "Changing Password of $($tbSAMAccountName.Text) to $($tbNewPassword.Text)"
            $u = Set-ADAccountPassword -Identity $tbSAMAccountName.Text -NewPassword (ConvertTo-SecureString -AsPlainText $tbNewPassword.Text -Force) -PassThru           
            if($passwordSetting.makePasswordTemporary){
                Set-ADUser -Identity $tbSAMAccountName.Text -ChangePasswordAtLogon $true
            }            
            [System.Windows.Forms.MessageBox]::Show("Password Changed")            
            Search-User
            $ChangePasswordWindow.Close()
        }

        $bConfirm = $ChangePasswordWindow.FindName("bConfirm")
        $bConfirm.Add_Click({Change-UserPassword})

        $bCancel = $ChangePasswordWindow.FindName("bCancel")
        $bCancel.Add_Click({$ChangePasswordWindow.Close()})

        $ChangePasswordWindow.ShowDialog() | Out-Null
    }else{
        throw "No user selected"
    }
}

function Create-SelectUserWindow{
    $SelectUserWindow = .\CreateWindow.ps1 -Path '..\Windows\SelectUserWindow.xaml'

    $SelectUserWindow.Owner = $MainWindow
    $SelectUserWindow.WindowStartupLocation = 'CenterOwner'

    $lbUsers = $SelectUserWindow.FindName('lbUsers')
    $userInfoOnServer = @(Get-ADUser -Filter $filter)
    foreach($user in $userInfoOnServer){
        $lbUsers.AddChild($user.SAMAccountName)
    }

    $bSelectUser = $SelectUserWindow.FindName('bSelectUser')
    $bSelectUser.Add_Click({Select-User})

    $bCancel = $SelectUserWindow.FindName('bCancel')
    $bCancel.Add_Click({
        $SelectUserWindow.Close()
        $tbEmployeeID.Text = ''
        $tbSAMAccountName.Text = ''
    })

    #Select user from window displaying list of users retrieved from Search-User command. When user selected, their
    #information is filled in the main window.
    function Select-User{
        if(!$lbUsers.SelectedItem){
            throw "No user selected"
        }
        $tbEmployeeID.Text = "Collecting data..."
        $tbSAMAccountName.Text = ""
        $user = $lbUsers.SelectedItem
        $iDisabledIcon.Visibility="Hidden"
        for($i = 0; $i -lt $dcs.Count; $i++){ 
                $userInfoOnServer = @(Get-ADUser $user -Server $dcs[$i] -Properties $properties)
                Set-Rows $i `
                    $(if($userInfoOnServer.LastBadPasswordAttempt){$userInfoOnServer.LastBadPasswordAttempt}else{'None'}) `
                    $(if($userInfoOnServer.PasswordLastSet){$userInfoOnServer.PasswordLastSet}else{"Change Password"}) `
                    $(if($userInfoOnServer.PasswordLastSet){[datetime]::FromFileTime($userInfoOnServer.'msDS-UserPasswordExpiryTimeComputed')}else{"N/A"}) `
                    $(if($userInfoOnServer.LockoutTime -gt 0){"Locked"}else{"Unlocked"}) `
                    $(if($userInfoOnServer.BadLogonCount){$userInfoOnServer.BadLogonCount}else{0}) `
                    $($dcs[$i].Name)
        }
        if($userInfoOnServer.Enabled){$iDisabledIcon.Visibility='Hidden'}else{$iDisabledIcon.Visibility='Visible'}
        $tbEmployeeID.Text = $userInfoOnServer.EmployeeID
        $tbSAMAccountName.Text = $userInfoOnServer.SAMAccountName
        $SelectUserWindow.Close()
    }
    
    $SelectUserWindow.ShowDialog() | Out-Null
}

function Clear-Window{
    for($i = 0; $i -lt $dcs.Count; $i++){             
        $rows[$i]["LastBadPassword"] = [string]::Empty
        $rows[$i]["PasswordLastSet"] = [string]::Empty
        $rows[$i]["PasswordExpirationDate"] = [string]::Empty
        $rows[$i]["LockedOut"] = [string]::Empty
        $rows[$i]["BadLogonCount"] = [DBNull]::Value
        $rows[$i]["DC Name"] = [string]::Empty
    }
    $tbEmployeeID.Text = "User Not Found"
    $tbSAMAccountName.Text = ""
}

function Search-Computer{    
    $lbSessions.Items.Clear()
    $tbComputerName.Text = ''
    #$tbIPAddress.Text =  ''
    $tbFreeDiskSpace.Text = ''
    $tbMemoryUsage.Text = ''
    $tbLastBootTime.Text = ''
    
    if($tbComputerSearch.Text){
        $computerName = @(Get-ADComputer -Identity $tbComputerSearch.Text)
        if(!(Test-Connection $tbComputerSearch.Text -Quiet -Count 1)){
            Throw "Unable to connect to $($tbComputerSearch.Text)" 
        }

        $alAvailableSessions = [System.Collections.ArrayList]::new()
            
        #Get sessions on computer from native Windows command qwinsta as string
        #Parse name, id and state of the active sessions from the string by getting 
        #index of where each of these fields are on each row.
        $sessions = (qwinsta /server $tbComputerSearch.Text).split("`n")
        $usernameIndex = $sessions[0].IndexOf('USERNAME')
        $IDIndex = $sessions[0].IndexOf('ID') - 2
        $stateIndex = $sessions[0].IndexOf('STATE')

        for($i = 1; $i -lt $sessions.count; $i++){
            if($sessions[$i].Substring($usernameIndex,1).Trim() -ne [string]::Empty){
                [void] $alAvailableSessions.Add([pscustomObject]@{
                    sessionName = $sessions[$i].Substring($usernameIndex,20).Trim()
                    sessionID = $sessions[$i].Substring($IDIndex,5).Trim()
                    sessionState = $sessions[$i].Substring($stateIndex,6).Trim()
                })
            }
        }

        foreach($session in $alAvailableSessions){
            $lbSessions.AddChild("$($session.sessionName)                  $($session.sessionID)                  $($session.sessionState)")
        }
        $tbComputerName.Text = $computerName.Name          
        $tbFreeDiskSpace.Text = "$((Get-WMIObject -ComputerName $computerName.Name -ClassName Win32_LogicalDisk | Where-Object {$_.DeviceID -eq 'C:'} | Select-Object  @{Name="FreeSpacePercent"; Expression={[Math]::Round(($_.FreeSpace / $_.Size) * 100)}}).FreeSpacePercent)%"
        try{          
            $AvailableMBytes = ((Get-Counter -ComputerName $computerName.Name -Counter '\Memory\Available MBytes').CounterSamples.CookedValue)
            $TotalPhysicalMemory = Get-WmiObject -ComputerName $computerName.Name -Class win32_ComputerSystem | Select-Object -ExpandProperty TotalPhysicalMemory 
            $Percent = [Math]::Round((1 - ($AvailableMBytes * 1MB)/$TotalPhysicalMemory) * 100)
            $tbMemoryUsage.Text = "$Percent%"
        }catch{
            $tbMemoryUsage.Text = 'Unavailable'
        }        
        $tbLastBootTime.Text = [Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject -ComputerName $computerName.Name -Class Win32_OperatingSystem).LastBootUpTime)
    }else{
        [System.Windows.Forms.MessageBox]::Show('No computer selected.')
    }
        
}

#Shadow selected session based on session information collected in Search-Computer function
function Start-Shadow{
    if($lbSessions.SelectedItem){
        $selectedSession = $lbSessions.SelectedItem
        $sessionID = (-split $selectedSession)[1]
        mstsc.exe /v:$($tbComputerName.Text) /shadow:$sessionID /f /span /control
    }else{
        [System.Windows.Forms.MessageBox]::Show("No session selected.")
    }
}

function Send-Email{
    $EmailWindow = .\CreateWindow.ps1 -Path '..\Windows\EmailWindow.xaml'
    if($tbSAMAccountName.Text){
        $user = Get-AdUser -Identity $tbSAMAccountName.Text -Properties EmailAddress
    }else{
        throw "No user selected"
    }
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.createItem(0)
    $mail.To = $user.EmailAddress
    $mail.Display()
    $signature = $mail.HTMLBody

    $EmailWindow.Owner = $MainWindow
    $EmailWindow.WindowStartupLocation = 'CenterOwner'

    $cbTemplate = $EmailWindow.FindName("cbTemplate")    
    $bSelectTemplate = $EmailWindow.FindName('bSelectTemplate')
    $bAddTemplate = $EmailWindow.FindName('bAddTemplate')
    $tbTemplateName = $EmailWindow.FindName('tbTemplateName')
    $bDeleteTemplate = $EmailWindow.FindName('bDeleteTemplate')

    $csv = Import-Csv '..\EmailTemplates.csv'
    $csv | ForEach-Object{$cbTemplate.AddChild($_.Name)}

    $bSelectTemplate.Add_Click({Select-Template})
    $bAddTemplate.Add_Click({Add-Template})
    $bDeleteTemplate.Add_Click({Delete-Template})

    function Select-Template{
        $csv = Import-Csv '..\EmailTemplates.csv'
        $mail.HTMLBody = ''
        $templateName = $cbTemplate.SelectedItem
        $template = $csv | Where-Object{$_.Name -eq $templateName} | Select-Object -ExpandProperty Template        
        $subject = $csv | Where-Object{$_.Name -eq $templateName} | Select-Object -ExpandProperty Subject        
        $bodyBuffer = ''

        $mail.Subject = "$subject"
        $template | ForEach-Object {$bodyBuffer += $_}
        $Mail.HTMLBody = "$(Get-EmailHeader) <br><br> $bodyBuffer <br> $signature"
    }

    Function Get-EmailHeader{
        $timeOfDay = if((Get-Date).Hour -lt 12){'morning'}else{'afternoon'}
        return "Good $timeOfDay, $($user.GivenName),`n`n"
    }

    function Add-Template{
        $fileSize = Get-Item '..\EmailTemplates.csv' | Select-Object -ExpandProperty Length
        if($fileSize -lt 5MB){
            $newTemplate = [pscustomobject]@{
                Name = $tbTemplateName.Text
                Subject = $mail.Subject
                Template = $mail.HTMLBody
            }|ConvertTo-Csv

            for($i = 2; $i -lt $newTemplate.Length; $i++){
                $newTemplate[$i] | Out-File -FilePath '..\EmailTemplates.csv' -Append
            }       

            $cbTemplate.Items.Clear()        
        
            $csv = Import-Csv '..\EmailTemplates.csv'
            $csv | ForEach-Object{$cbTemplate.AddChild($_.Name)}  
        }else{
            [System.Windows.Forms.MessageBox]::Show('Email Templates file is too large. Delete templates to make more room.')
        } 
    }  

    function Delete-Template{
        $confirmEmailDeleteWindow = .\CreateWindow.ps1 -Path '..\Windows\ConfirmEmailDeleteWindow.xaml'
        $bConfirmDelete = $confirmEmailDeleteWindow.FindName('bConfirm')
        $bCancelDelete = $confirmEmailDeleteWindow.FindName('bCancel')
        $lDeleteEmailPrompt = $confirmEmailDeleteWindow.FindName('lDeleteEmailPrompt')
        $lDeleteEmailPrompt.Content = "Are you sure you want to delete email template $($cbTemplate.SelectedItem)?"

        $bConfirmDelete.Add_Click(
            {
                $templateName = $cbTemplate.SelectedItem
                $csv = Import-Csv -Path '..\EmailTemplates.csv'
                $templates = [System.Collections.ArrayList]::new(@($csv))
        
                for($i = 0; $i -lt $templates.Count; $i++){
                    if($templates[$i].Name -eq $templateName){
                        $index = $i
                    }
                }

                $templates.RemoveAt($index)

                if($templates.Count -eq 0){
                    '"Name","Subject","Template"'| Out-File '..\EmailTemplates.csv'
                }else{
                    $templates | ConvertTo-Csv | Out-File '..\EmailTemplates.csv'
                }                
                               
                $cbTemplate.Items.Clear()        
        
                $csv = Import-Csv '..\EmailTemplates.csv'
                $csv | ForEach-Object{$cbTemplate.AddChild($_.Name)} 
                $confirmEmailDeleteWindow.close()
            }
        )

        $bCancelDelete.Add_Click({$confirmEmailDeleteWindow.close()})

        $confirmEmailDeleteWindow.ShowDialog() | Out-Null
    }

    $EmailWindow.ShowDialog()|out-null
}

function Create-UserInfoWindow{
    if($tbSAMAccountName.Text){
        $UserInfoWindow = .\CreateWindow.ps1 -Path '..\Windows\UserInfoWindow.xaml'
        $UserInfoWindow.Title = $tbSAMAccountName.Text
        $UserInfoWindow.Owner = $MainWindow
        $UserInfoWindow.WindowStartupLocation = 'CenterOwner'

        $tbEmailAddress = $UserInfoWindow.FindName('tbEmailAddress')
        $tbDescription = $UserInfoWindow.FindName('tbDescription')
        $tbAddress = $UserInfoWindow.FindName('tbAddress')
        $tbTelephone = $UserInfoWindow.FindName('tbTelephone')
        $tbMobilePhone = $UserInfoWindow.FindName('tbMobilePhone')
        $tbOtherLoginWorkstation = $UserInfoWindow.FindName('tbOtherLoginWorkstation')
        $tbCanonicalName = $UserInfoWindow.FindName('tbCanonicalName')
        $tbProfilePath = $UserInfoWindow.FindName('tbProfilePath')
        $tbExpiresOn = $UserInfoWindow.FindName('tbExpiresOn')
        $lbMemberOf = $UserInfoWindow.FindName('lbMemberOf')
        
        $Properties = @('EmailAddress','Description','Office','telephoneNumber','MobilePhone','otherLoginWorkstations','CanonicalName','HomeDirectory','AccountExpirationDate','MemberOf')

        $User = Get-ADUser -Filter {SAMAccountName -eq $tbSAMAccountName.Text} -Properties $Properties
        $tbAddress.Text = $User.Office
        $tbDescription.Text = $User.Description
        $tbEmailAddress.Text = $User.EmailAddress
        $tbTelephone.Text = $User.telephoneNumber
        $tbMobilePhone.Text = $User.MobilePhone
        $tbOtherLoginWorkstation.Text = if($User.otherLoginWorkstations){$User.otherLoginWorkstations}else{"Not Set"}
        $tbCanonicalName.Text = $User.CanonicalName
        $tbProfilePath.Text = $User.HomeDirectory
        $tbExpiresOn.Text = if($User.AccountExpirationDate){$User.AccountExpirationDate}else{"Never"}               
        $user.MemberOf | Sort-Object -Property {($_ -split ',')[0].Substring(3)} | ForEach-Object {$lbMemberOf.AddChild(($_ -split ',')[0].Substring(3))}

        $UserInfoWindow.ShowDialog() | Out-Null
    }else{
        throw "No user Selected"
    }
}


$bSearch = $MainWindow.FindName("bSearch")
$bSearch.Add_Click({try{Search-User}
    catch{
        [System.Windows.Forms.MessageBox]::Show($_)
        $tbEmployeeID.Text = ""
        $tbSAMAccountName.Text = ""
    }
})

$bUnlock = $MainWindow.FindName("bUnlock")
$bUnlock.Add_Click({try{Unlock-User}catch{[System.Windows.Forms.MessageBox]::Show($_)}})

$bChangePassword = $MainWindow.FindName("bChangePassword")
$bChangePassword.Add_Click({try{Create-PasswordWindow}catch{[System.Windows.Forms.MessageBox]::Show($_)}})

$bSearchComputer = $MainWindow.FindName("bSearchComputer")
$bSearchComputer.Add_Click({try{Search-Computer}catch{[System.Windows.Forms.MessageBox]::Show($_)}})

$bShadow = $MainWindow.FindName("bShadow")
$bShadow.Add_Click({try{Start-Shadow}catch{[System.Windows.Forms.MessageBox]::Show($_)}})

$bSendEmail = $MainWindow.FindName('bSendEmail')
$bSendEmail.Add_Click({try{Send-Email}catch{[System.Windows.Forms.MessageBox]::Show($_)}})

$bMoreUserInfo = $MainWindow.FindName('bMoreUserInfo')
$bMoreUserInfo.Add_Click({try{Create-UserInfoWindow}catch{[System.Windows.Forms.MessageBox]::Show($_)}})


$MainWindow.ShowDialog() | Out-Null
$MainWindow.BringIntoView()