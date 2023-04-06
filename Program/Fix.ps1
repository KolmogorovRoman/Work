$wshell = New-Object -ComObject Wscript.Shell
$SessionID = (Get-Process -PID $PID).SessionId
$id_1c = (Get-Process "1cv8" -ErrorAction Ignore |? {$_.SI -eq $SessionID}).id
if ($id_1c -eq $null) { echo "Должна быть запущена 1С"; return }
elseif ($id_1c -is [System.Object[]]) { echo "Должна быть запущена одна 1С"; return}
else {$wshell.AppActivate($id_1c) > $null}

$wshell.SendKeys("^{HOME}{RIGHT}{DOWN}") > $null
$wshell.SendKeys("{F10} {ESC}{RIGHT}{RIGHT} {DOWN}{RIGHT}{DOWN}{DOWN}{ENTER}") > $null
$wshell.SendKeys("{F10} {ESC}{RIGHT}{RIGHT} {DOWN}{RIGHT}{DOWN}{DOWN}{DOWN}{DOWN}{ENTER}") > $null
$wshell.SendKeys("%{ENTER}{DOWN}^{DOWN}{DOWN}{DOWN}{DOWN}{ENTER}") > $null
$wshell.SendKeys("{F10} {ESC}{RIGHT}{RIGHT} {UP}{UP}{RIGHT}{UP}{UP}{ENTER}{TAB}{TAB}{TAB}{ }{TAB}{ }{ENTER}") > $null
