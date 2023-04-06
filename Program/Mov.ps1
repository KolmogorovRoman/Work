$Global:name = [string]$args[0]
if (-not (Test-Path $work\$name)) {(mkdir $work\$name) > $null}
if (-not (Test-Path $work\$name\ТН)) {(mkdir $work\$name\ТН) > $null}
if (-not (Test-Path $work\$name\Счет-фактуры)) {(mkdir $work\$name\Счет-фактуры) > $null}
if (-not (Test-Path $work\$name\CMR)) {(mkdir $work\$name\CMR) > $null}
if (-not (Test-Path $work\$name\Сводные)) {(mkdir $work\$name\Сводные) > $null}
cd $work\$name
$host.ui.RawUI.WindowTitle = "Script $name"
if ($sorting -ne $null) {Stop-Process $sorting}
$Global:sorting = start powershell -ArgumentList "-noexit", "-WindowStyle Minimized", "$program\Sorting.ps1 $work $name" -PassThru