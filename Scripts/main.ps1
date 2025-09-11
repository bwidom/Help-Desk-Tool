Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
[xml]$XAML = Get-Content '../Windows/MainWindow.xaml'

$XAML.Window.RemoveAttribute('x:Class')
$XAML.Window.RemoveAttribute('mc:Ignorable')
$XAMLReader = New-Object System.Xml.XmlNodeReader $XAML
$MainWindow = [Windows.Markup.XamlReader]::Load($XAMLReader)
$XAML.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $MainWindow.FindName($_.Name)}

$dgAccountInfo = $MainWindow.FindName("dgAccountInfo")
$cbSearchCriteria = $MainWindow.FindName("cbSearchCriteria")
$tbSearchUser = $MainWindow.FindName("tbSearchUser")
$lEmployeeID = $MainWindow.FindName("lEmployeeID")
$lSAMAccountName = $MainWindow.FindName("lSAMAccountName")
$tbComputerSearch = $MainWindow.FindName("tbComputerSearch")
$lbSessions = $MainWindow.FindName("lbSessions")
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
    $lEmployeeID.Text = "Collecting data..."
    $lSAMAccountName.Text = ""
    $iDisabledIcon.Visibility="Hidden"
    #Forces the text box controls to update
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
        $lEmployeeID.Text = $userInfoOnServer.EmployeeID
        $lSAMAccountName.Text = $userInfoOnServer.SAMAccountName
    }elseif($countUser.Count -eq 0){
        for($i = 0; $i -lt $dcs.Count; $i++){       
            Set-Rows -RowIndex $i      
        }
        $lEmployeeID.Text = "User Not Found"
        $lSAMAccountName.Text = ""
    }elseif($countUser.Count -gt 1){
        Create-SelectUserWindow
       
    }
}

$bSearch = $MainWindow.FindName("bSearch")
$bSearch.Add_Click({Search-User})


$MainWindow.ShowDialog() | Out-Null