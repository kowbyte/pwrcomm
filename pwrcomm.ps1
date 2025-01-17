<#Written by Kaden Spletzer#>

#Dependancies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#List of comports on the system
$comports = [System.IO.Ports.SerialPort]::GetPortNames()
Write-Host "! Starting powercomm !" -BackgroundColor Blue -ForegroundColor Black
#Write keystrokes to interface
function writeConsole($port){
    if ([Console]::KeyAvailable){
        $key = [Console]::ReadKey($true)
        switch ($key.key) {
            UpArrow{$port.write($([char]27 + "[" + "A")); return}
            DownArrow{$port.write($([char]27 + "[" + "B")); return}
            RightArrow{$port.write($([char]27 + "[" + "C")); return}
            LeftArrow{$port.write($([char]27 + "[" + "D")); return}
            Delete{$port.write($([char]27 + "[" + "3" + "~")); return}
            {($_ -eq [ConsoleKey]::F1) -or ($_ -eq [ConsoleKey]::F2)}{$port.close();Write-Host "`nPort has been closed!" -ForegroundColor Red; return}
            Default{$port.write($key.KeyChar)}
        }
        <#$char = $key.KeyChar
        switch ($char) {
            "~" { $port.close();Write-Host "`nPort has been closed!" -ForegroundColor Red }
            Default {$port.write($char)}
        }#>
    }
}

#Read data from interface
function readConsole($port){
    try {
        $line = $port.ReadExisting()
        if ($line){
            Write-Host -NoNewline $line
        }
    }
    catch [System.Exception]{
        return
    }
}

#Form creation for specifying COM port and Baud rate
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(315,230)
$form.Text = 'pwrcomm'
$form.StartPosition = 'CenterScreen'

$acceptBtn = New-Object System.Windows.Forms.Button
$acceptBtn.Location = New-Object System.Drawing.Point (75, 160)
$acceptBtn.Size = New-Object System.Drawing.Size(75,23)
$acceptBtn.Text = 'Start'
$acceptBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $acceptBtn
$form.Controls.Add($acceptBtn)

$negativeBtn = New-Object System.Windows.Forms.Button
$negativeBtn.Location = New-Object System.Drawing.Point (150, 160)
$negativeBtn.Size = New-Object System.Drawing.Size(75,23)
$negativeBtn.Text = 'Cancel'
$negativeBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.AcceptButton = $negativeBtn
$form.Controls.Add($negativeBtn)

#labels
$comLabel = New-Object System.Windows.Forms.Label
$comLabel.Location = New-Object System.Drawing.Point(10,20)
$comLabel.Size = New-Object System.Drawing.Size(280,20)
$comLabel.Text = 'Select a COM port from the available list:'
$form.Controls.Add($comLabel)

$baudLabel = New-Object System.Windows.Forms.Label
$baudLabel.Location = New-Object System.Drawing.Point(10,110)
$baudLabel.Size = New-Object System.Drawing.Size(280,20)
$baudLabel.Text = 'Enter a baud rate:'
$form.Controls.Add($baudLabel)

#input fields
$baud = New-Object System.Windows.Forms.TextBox
$baud.Location = New-Object System.Drawing.Point(10,130)
$baud.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($baud)

$comList = New-Object System.Windows.Forms.ListBox
$comList.Location = New-Object System.Drawing.Point(10, 40)
$comList.Size = New-Object System.Drawing.Size(280,20)
$comList.Height = 80

$form.Controls.Add($acceptBtn)

#Add system comports to comport list
foreach($com in $comports)
{
    [void] $comList.Items.Add($com)
}

$form.Controls.Add($comList)

$form.Topmost = $true

function main(){
    $formResult = $form.ShowDialog()
    if($formResult -eq [System.Windows.Forms.DialogResult]::OK){
        try{
            $baud = [System.Convert]::ToInt32($baud.Text)
        }catch{
            Write-Host -BackgroundColor Yellow "The inputted value is not a number!"
            main
            Return
        }
        Write-Host "### Press F1 or F2 to quit ### " -BackgroundColor Yellow -ForegroundColor Black
        try{
            $port = New-Object System.IO.Ports.SerialPort $comList.SelectedItem,$baud,None,8,one
            $port.DtrEnable = $true
            $port.RtsEnable = $true
            $port.Open()
            Write-Host -BackgroundColor Green "Connection to $($comList.SelectedItem) established!" -ForegroundColor Black
        }catch{
            Write-Host [System.Exception]
        }
        do{
            readConsole($port)
            writeConsole($port)
        }while($port.IsOpen)
        Write-Host -BackgroundColor Red " Closing powercomm " -ForegroundColor Black
    }else{
        Write-Host -BackgroundColor Red " Closing powercomm " -ForegroundColor Black
        exit
    }
}
main