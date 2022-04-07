$wshell = New-Object -ComObject Wscript.Shell
$wshell = New-Object -ComObject Wscript.Shell
$printersFolder = "C:\Windows\System32\spool\PRINTERS\*"

function CleanFolder {
    param ($Folder)

    Remove-Item $Folder
}

Stop-Service -Name Spooler -Force
CleanFolder($printersFolder)
Start-Service -Name Spooler

$wshell.Popup("Serviço de impressão reiniciado. Caso o problema de comunicação com a impressora persista, contate a TI no ramal 204.",0,"Done",0x1)

