#Rave case coding for on-prem teams 
#ajayx V1 16th May 2019 

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select a SME area'
$form.Size = New-Object System.Drawing.Size(400,200)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,25)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,25)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please select support area'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(360,50)
$listBox.Height = 90

[void] $listBox.Items.Add('<SEARCH>')
[void] $listBox.Items.Add('<ADMIN>')
[void] $listBox.Items.Add('<LISTLIBRARY>')
[void] $listBox.Items.Add('<PERFORMANCE>')
[void] $listBox.Items.Add('<WORKFLOW>')
[void] $listBox.Items.Add('<UPGRADE>')
[void] $listBox.Items.Add('<AUTHENTICATION>')
[void] $listBox.Items.Add('<UPA>')
[void] $listBox.Items.Add('<SETUP>')
[void] $listBox.Items.Add('<BI>')
[void] $listBox.Items.Add('<OCSI>')
[void] $listBox.Items.Add('<INFOPATH>')
[void] $listBox.Items.Add('<HYBRID>')
[void] $listBox.Items.Add('<PROJECT>')
[void] $listBox.Items.Add('<FS14>')
[void] $listBox.Items.Add('<ESP>')
[void] $listBox.Items.Add('<SPD>')
[void] $listBox.Items.Add('<SPDEV>')

$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $listBox.SelectedItem
    #$x
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'SP Version'
$form.Size = New-Object System.Drawing.Size(400,200)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,25)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,25)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'SharePoint Version'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(360,50)
$listBox.Height = 90

[void] $listBox.Items.Add('<SP 2010>')
[void] $listBox.Items.Add('<SP 2013>')
[void] $listBox.Items.Add('<SP 2016>')
[void] $listBox.Items.Add('<SP 2019>')
[void] $listBox.Items.Add('<SPO>')
[void] $listBox.Items.Add('<NON>')

$form.Controls.Add($listBox)
$form.Topmost = $true
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $y = $listBox.SelectedItem
    #$y
}


$form = New-Object System.Windows.Forms.Form
$form.Text = 'Issue description'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Enter Issue description below'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $z = $textBox.Text
    #$z
}


Write-host -ForegroundColor Green $x "|" $y  "|"  $z
$x +"|"+ $y + "|" +" "+ $z |Set-Clipboard

if ($Host.Name -eq "ConsoleHost")
{
    Write-Host "Result is saved to clipboard. Press enter to continue..."
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") > $null
}

