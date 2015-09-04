<#Import Modules#>
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") > $null


#Setting Variables

$clientPC = Get-ADComputer -Identity $env:CLIENTNAME -Properties Location
if ($clientpc.Location -eq $null){ exit }
$clientPrinter = Get-Printer -computerName $clientpc.Name | Where-Object {$_.Name -contains $clientPC.Location}
$clientPrinterPath = "\\" + $clientPrinter.ComputerName + "\" + $clientPrinter.Name
Add-Printer -ConnectionName $clientPrinterPath


$WMIPrinter = Get-WmiObject -Class Win32_Printer | where {$_.Name -like $clientPrinterPath }
#$_.Name -match $clientPrinter.Name -and 

#(Get-WmiObject -ComputerName . -Class Win32_Printer -Filter "Name='HP LaserJet 5Si'").SetDefaultPrinter()

[System.Windows.Forms.MessageBox]::Show("Settings " + $WMIPrinter.DeviceID + " as default Printer")
$WMIPrinter.setdefaultprinter()

Exit

<#
$clientPrinterPath = "\\" + $clientPrinter.ComputerName + "\" + $clientPrinter.Name
(New-Object -ComObject WScript.Network).AddWindowsPrinterConnection($clientPrinterPath)

#>