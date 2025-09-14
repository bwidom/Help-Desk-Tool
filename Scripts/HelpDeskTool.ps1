$MainWindow = .\CreateWindow.ps1 -Path '../Windows/MainWindow.xaml'

$dgAccountInfo = $MainWindow.FindName("dgAccountInfo")
$cbSearchCriteria = $MainWindow.FindName("cbSearchCriteria")
$tbSearchUser = $MainWindow.FindName("tbSearchUser")
$tbEmployeeID = $MainWindow.FindName("tbEmployeeID")
$tbSAMAccountName = $MainWindow.FindName("tbSAMAccountName")
$tbComputerSearch = $MainWindow.FindName("tbComputerSearch")
$tbSessions = $MainWindow.FindName("tbSessions")
$tbComputerName = $MainWindow.FindName('tbComputerName')
$tbIPAddress = $MainWindow.FindName('tbIPAddress')
$tbFreeDiskSpace = $MainWindow.FindName('tbFreeDiskSpace')
$tbMemoryUsage = $MainWindow.FindName('tbMemoryUsage')
$tbLastBootTime = $MainWindow.FindName('tbLastBootTime')
$iDisabledIcon = $MainWindow.FindName('iDisabledIcon')

$tbSearchUser.Focus() | Out-Null

$dataTable = New-Object System.Data.DataTable

[void]$dataTable.Columns.Add("DC Name", [string])
[void]$dataTable.Columns.Add("LastBadPassword", [string])
[void]$dataTable.Columns.Add("PasswordLastSet", [string])
[void]$dataTable.Columns.Add("PasswordExpirationDate", [string])
[void]$dataTable.Columns.Add("LockedOut", [string])
[void]$dataTable.Columns.Add("BadLogonCount", [int])

#Properties for Get-ADUser command
$properties = @("LastBadPasswordAttempt", "PasswordLastSet", "msDS-UserPasswordExpiryTimeComputed", "BadLogonCount", "LockedOut", "EmployeeID", "SAMAccountName")

$dgAccountInfo.ItemsSource = $dataTable.DefaultView

$dcs = @(Get-ADDomainController -Filter * | Sort-Object -Property Name)
$rows = [Object[]]::new($dcs.Count)
$domainDistinguishedName = (Get-ADDomain).DistinguishedName


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
    $countUser = @(Get-ADUser -Filter $filter -SearchBase $domainDistinguishedName)

    if($countUser.Count -eq 1){
        for($i = 0; $i -lt $dcs.Count; $i++){             
                $userInfoOnServer = @(Get-ADUser -Server $dcs[$i] -Filter $filter -Properties $properties -SearchBase $domainDistinguishedName)
                Set-Rows $i `
                    $(if($userInfoOnServer.LastBadPasswordAttempt){$userInfoOnServer.LastBadPasswordAttempt}else{'None'}) `
                    $(if($userInfoOnServer.PasswordLastSet){$userInfoOnServer.PasswordLastSet}else{"Change Password"}) `
                    $(if($userInfoOnServer.PasswordLastSet){[datetime]::FromFileTime($userInfoOnServer.'msDS-UserPasswordExpiryTimeComputed')}else{"N/A"}) `
                    $(if((Get-ADUser -Filter $filter -Properties * | Select-Object -ExpandProperty lockoutTime) -gt 0){"Locked"}else{"Unlocked"}) `
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
        $tbEmployeeID.Text = "User Not Found"
        $tbSAMAccountName.Text = ""
    }elseif($countUser.Count -gt 1){
        Create-SelectUserWindow
    }
}

function Unlock-User{
    #Check if a user is selected.
    if($tbSAMAccountName.Text){
        foreach($dc in $dcs){
                Unlock-ADAccount -Identity $tbSAMAccountName.Text -Server $dc
        }
        Search-User
    }else{
        Write-Host "No User Selected"
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
            Set-ADUser -Identity $tbSAMAccountName.Text -ChangePasswordAtLogon $true
            [System.Windows.Forms.MessageBox]::Show("Password Changed")
            Write-Host "$(($u.Name)) password changed to $($tbNewPassword.Text). Password must change at login"
            Search-User
            $ChangePasswordWindow.Close()
        }

        $bConfirm = $ChangePasswordWindow.FindName("bConfirm")
        $bConfirm.Add_Click({Change-UserPassword})

        $bCancel = $ChangePasswordWindow.FindName("bCancel")
        $bCancel.Add_Click({$ChangePasswordWindow.Close()})

        $ChangePasswordWindow.ShowDialog() | Out-Null
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
        $lEmployeeID.Text = ''
        $lSAMAccountName.Text = ''
    })

    #Select user from window displaying list of users retrieved from Search-User command. When user selected, their
    #information is filled in the main window.
    function Select-User{
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
                    $(if((Get-ADUser -Filter $filter -Properties * | Select-Object -ExpandProperty lockoutTime) -gt 0){"Locked"}else{"Unlocked"}) `
                    $(if($userInfoOnServer.BadLogonCount){$userInfoOnServer.BadLogonCount}else{0}) `
                    $($dcs[$i].Name)
        }
        if($countUser[0].Enabled){$iDisabledIcon.Visibility='Hidden'}else{$iDisabledIcon.Visibility='Visible'}
        $tbEmployeeID.Text = $userInfoOnServer.EmployeeID
        $tbSAMAccountName.Text = $userInfoOnServer.SAMAccountName
        $SelectUserWindow.Close()
    }
    
    $SelectUserWindow.ShowDialog() | Out-Null
}

$bSearch = $MainWindow.FindName("bSearch")
$bSearch.Add_Click({Search-User})

$bUnlock = $MainWindow.FindName("bUnlock")
$bUnlock.Add_Click({Unlock-User})

$bChangePassword = $MainWindow.FindName("bChangePassword")
$bChangePassword.Add_Click({Create-PasswordWindow})

$MainWindow.ShowDialog() | Out-Null