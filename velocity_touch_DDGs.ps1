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
Clear-Variable IncDDGs -ErrorAction SilentlyContinue

# Test And Connect To AzureAD If Needed
try {
    Write-Verbose -Message "Testing connection to Azure AD"
    Get-AzureAdDomain -ErrorAction Stop | Out-Null
    Write-Verbose -Message "Already connected to Azure AD"
}
catch {
    Write-Verbose -Message "Connecting to Azure AD"
    Connect-AzureAD
}

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
$RemoveDDGButton.Text = 'Remove DDG(s)'
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
$CreateDDGButton.Text = 'Create DDG(s) From Excel File'
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
    
    # Create Domain Selection Form
    $DomainSelectionForm = New-Object System.Windows.Forms.Form
    $DomainSelectionForm.Text = "Please Select Domain"
    $DomainSelectionForm.Autosize = $True
    $DomainSelectionForm.MaximizeBox = $False
    $DomainSelectionForm.StartPosition = "CenterScreen"
    $DomainSelectionForm.TopMost = $True
    
    # Create OK Button for Domain Selection Form
    $DomainOKButton = New-Object System.Windows.Forms.Button
    $DomainOKButton.TabIndex = 1
    $DomainOKButton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $DomainOKButton.Text = 'OK'
    $DomainOKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $DomainSelectionForm.AcceptButton = $DomainOKButton
    $DomainSelectionForm.Controls.Add($DomainOKButton)

    # Create Combobox to Select Domain
    $DomainCombobox = New-Object System.Windows.Forms.ComboBox
    $DomainCombobox.TabIndex = 0
    $DomainCombobox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $DomainCombobox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $DomainCombobox.FormattingEnabled = $true
    foreach($domain in Get-AzureADDomain){
        $null = $DomainCombobox.Items.add($domain.Name)
    }
    $DomainComboBox.SelectedIndex = 0
    $DomainSelectionForm.Controls.Add($DomainCombobox)

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
    # Pull Dynamic Distribution Groups and Present Selection Out-Gridview
    $ActiveDDGs = Get-DynamicDistributionGroup | Out-GridView -Passthru -Title "Select DDG(s) To Remove"
    #Verify Active DDGs
    if ($ActiveDDGs){
        #Remove Each Selected DDG
        foreach($ActiveDDG in $ActiveDDGs){
            Remove-DynamicDistributionGroup -Identity $ActiveDDG.Name -Confirm:$False
        }
        Write-Verbose "All Selected Dynamic Distribution Groups Removed"
    }
    else { 
        Write-Verbose "No Active Dynamic Distribution Groups OR Cancel Button Was Selected"
        Throw
    }
}