# Load Exchange 2010 Management Snap-In if not already loaded
$snapinName = "Microsoft.Exchange.Management.PowerShell.E2010"
if (-not (Get-PSSnapin | Where-Object { $_.Name -eq $snapinName })) {
    Add-PSSnapin $snapinName -ErrorAction Stop
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select and Process CSV File for UPN Change'
$form.Size = New-Object System.Drawing.Size(400,200)
$form.StartPosition = 'CenterScreen'

# Create the browse button
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(10,10)
$browseButton.Size = New-Object System.Drawing.Size(100,23)
$browseButton.Text = 'Browse'
$form.Controls.Add($browseButton)

# Create the text box for file path
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(120,10)
$textBox.Size = New-Object System.Drawing.Size(260,23)
$textBox.ReadOnly = $true
$form.Controls.Add($textBox)

# Create OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(10,50)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.Enabled = $false
$form.Controls.Add($okButton)

# Create Cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(100,50)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$form.Controls.Add($cancelButton)

# Browse button click event
$browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = 'CSV Files (*.csv)|*.csv'
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBox.Text = $openFileDialog.FileName
        $okButton.Enabled = $true
    }
})

# OK button click event
$okButton.Add_Click({
    ProcessFile $textBox.Text
})

# Cancel button click event
$cancelButton.Add_Click({
    $form.Close()
})

# Function to process the file
function ProcessFile($filePath) {
    try {
        # Reading CSV file
        $users = Import-Csv $filePath

        foreach ($user in $users) {
            try {
                Write-Host "Processing user $($user.samAccountName)"
                # Change UPN and email address
                Set-RemoteMailbox -Identity $user.samAccountName -UserPrincipalName $user.NewUPN -PrimarySmtpAddress $user.NewUPN -ErrorAction Stop

                Write-Host "Updated user $($user.samAccountName)"
            } catch {
                $errorMessage = $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show("Error updating user $($user.samAccountName): $errorMessage")
                break
            }
        }

        [System.Windows.Forms.MessageBox]::Show("File processing completed.")
    } catch {
        $errorMessage = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Error: $errorMessage")
    }
}

# Show the form
$form.ShowDialog()
