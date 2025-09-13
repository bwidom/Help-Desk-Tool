param(
    [string] $Path
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
[xml]$XAML = Get-Content $Path

$XAML.Window.RemoveAttribute('x:Class')
$XAML.Window.RemoveAttribute('mc:Ignorable')
$XAMLReader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($XAMLReader)
$XAML.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $Window.FindName($_.Name)}

return $Window