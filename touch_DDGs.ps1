#Requires -Module ExchangeOnlineManagement,ImportExcel
[CmdletBinding()]
Param()

#Test And Connect To Microsoft Exchange Online If Needed
try {
    Write-Verbose -Message "Testing connection to Microsoft Exchange Online"
    Get-Mailbox -ErrorAction Stop | Out-Null
    Write-Verbose -Message "Already connected to Microsoft Exchange Online"
}
catch {
    Write-Verbose -Message "Connecting to Microsoft Exchange Online"
    Connect-ExchangeOnline
}

$ExcelDoc = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'SpreadSheet (*.xlsx)|*.xlsx'
}
$null = $ExcelDoc.ShowDialog()

try
{
    $DDGs = Import-Excel -Path $ExcelDoc.Filename -ErrorAction Stop
}
catch
{
    Write-Verbose "Cancelled."
}

foreach ($DDG in $DDGs) {
    $recipientFilter = "(HiddenFromAddressListsEnabled -eq '$False' -and RecipientTypeDetails -eq 'UserMailbox')"
    if ($ddg.city -and $ddg.state) {
        $recipientFilter += " -and (city -eq '$($ddg.city)' -and stateorprovince -eq '$($ddg.state)')"
    }
    if ($ddg.Type) {
        $recipientFilter += " -and (CustomAttribute1 -eq '$($ddg.type)')"
    }
    $SMTP = $DDG.Alias + "@velocityclinical.com"
    New-DynamicDistributionGroup -Name $DDG.Group -Alias $DDG.alias -PrimarySMTPAddress $SMTP -RecipientFilter $recipientFilter
}