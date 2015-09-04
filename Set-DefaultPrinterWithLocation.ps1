#Setting Variables

$clientPC = Get-ADComputer -Identity $env:CLIENTNAME -Properties Location

#Ici, si location est vide, exit
if ($clientpc.Location -eq $null){ exit }

$clientPrinter = Get-Printer -computerName $clientpc.Name | Where-Object {$_.Name -contains $clientPC.Location}
$clientPrinterPath = "\\" + $clientPrinter.ComputerName + "\" + $clientPrinter.Name
Add-Printer -ConnectionName $clientPrinterPath


$WMIPrinter = Get-WmiObject -Class Win32_Printer | where {$_.Name -like $clientPrinterPath }

#[System.Windows.Forms.MessageBox]::Show("Settings " + $WMIPrinter.DeviceID + " as default Printer")

$WMIPrinter.setdefaultprinter()

Exit