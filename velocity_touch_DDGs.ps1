#Requires -Module ExchangeOnlineManagement,ImportExcel
#Allows -Verbose To Work
[CmdletBinding()]
Param()
#Include GUI Elements in Script
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()
Clear-Variable SelectionFormResult -ErrorAction SilentlyContinue
Clear-Variable Exceldoc -ErrorAction SilentlyContinue

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

# Create Selection Form 
$SelectionForm = New-Object System.Windows.Forms.Form
$SelectionForm.Autosize = $True
$SelectionForm.MaximizeBox = $False
$SelectionForm.StartPosition = "CenterScreen"
$SelectionForm.TopMost = $True

# Create "Edit DDG" Button
$RemoveDDGButton = New-Object System.Windows.Forms.Button
$RemoveDDGButton.TabIndex = 2
$RemoveDDGButton.Dock = [System.Windows.Forms.DockStyle]::Top
$RemoveDDGButton.Text = 'Remove DDG'
$RemoveDDGButton.DialogResult = 3
$SelectionForm.Controls.Add($RemoveDDGButton)

# Create "Edit DDG" Button
$EditDDGButton = New-Object System.Windows.Forms.Button
$EditDDGButton.TabIndex = 1
$EditDDGButton.Dock = [System.Windows.Forms.DockStyle]::Top
$EditDDGButton.Text = 'Edit DDG'
$EditDDGButton.DialogResult = 2
$SelectionForm.Controls.Add($EditDDGButton)

# Create "Create DDG" Button
$CreateDDGButton = New-Object System.Windows.Forms.Button
$CreateDDGButton.TabIndex = 0
$CreateDDGButton.Dock = [System.Windows.Forms.DockStyle]::Top
$CreateDDGButton.Text = 'Create DDG'
$CreateDDGButton.DialogResult = 1
$SelectionForm.Controls.Add($CreateDDGButton)

$SelectionFormResult = $SelectionForm.ShowDialog()

if($SelectionFormResult -eq 1){
    # Request Excel Document in File Dialog Window, Default Location: User's Desktop
    $ExcelDoc = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        Filter = 'SpreadSheet (*.xlsx)|*.xlsx'
    }
    $null = $ExcelDoc.ShowDialog()
    
    # Import From Chosen Excel File
    $IncDDGs = Import-Excel -Path $ExcelDoc.Filename -ErrorAction Stop
    
    $DomainSelectionForm = New-Object System.Windows.Forms.Form
    $DomainSelectionForm.Autosize = $True
    $DomainSelectionForm.MaximizeBox = $False
    $DomainSelectionForm.StartPosition = "CenterScreen"
    $DomainSelectionForm.TopMost = $True
    
    $DomainOKButton = New-Object System.Windows.Forms.Form
    $DomainOKButton.TabIndex = 6
    $DomainOKButton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $DomainOKButton.Text = 'OK'
    $DomainOKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $DomainOKButton.Enabled = $false
    $DomainSelectionForm.AcceptButton = $DomainOKButton
    $DomainSelectionForm.Controls.Add($DomainOKButton)

    $domainComboBox = New-Object System.Windows.Forms.ComboBox
    $domainComboBox.TabIndex = 0
    $domainComboBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $domainComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $domainComboBox.FormattingEnabled = $true
    foreach($domain in Get-AzureADDomain){
        $null = $domainComboBox.Items.add($domain.Name)
    }
    $DomainSelectionForm.Controls.Add($domainComboBox)

    $DomainSelectionResult = $DomainSelectionForm.ShowDialog()

    if($DomainSelectionResult = [System.Windows.Forms.DialogResult]::OK){
        foreach ($DDG in $IncDDGs) {
            $recipientFilter = "(HiddenFromAddressListsEnabled -eq '$False' -and RecipientTypeDetails -eq 'UserMailbox')"
            if ($ddg.city -and $ddg.state) {
                $recipientFilter += " -and (city -eq '$($ddg.city)' -and stateorprovince -eq '$($ddg.state)')"
            }
            if ($ddg.Type) {
                $recipientFilter += " -and (CustomAttribute1 -eq '$($ddg.type)')"
            }
            $SMTP = $DDG.Alias + "@" + $DomainComboBox.Text
            New-DynamicDistributionGroup -Name $DDG.Group -Alias $DDG.alias -PrimarySMTPAddress $SMTP -RecipientFilter $recipientFilter
        }
    }
}
elseif($SelectionFormResult -eq 2){
    Write-Host "You Picked Edit, Which Is Not Functional At This Time.  Hello World!"
}
elseif ($SelectionFormResult -eq 3) {
    Write-Host "You Picked Remove, Which Is Not Functional At This Time.  Hello World!"
}