if ($args[0] -in @("set", "ыуе"))
{
    $sets_info = log set
    $sets_names = $sets_info | Where-Object {$_.Файл -ne $null} | ForEach-Object {"=$($_.Клиент)"} 
    & $program\ClipVals.ps1 @sets_names "e27" "N27" "P67" "p38" "n40" "sera"
}
elseif ($args[0] -in @("sample", "ыфьзду"))
{
    Get-Content $data\ClipSample.txt | echo
    Get-Content $data\ClipSample.txt | Set-Clipboard
}
elseif ($args[0] -in @("last", "дфые"))
{
    Get-Content $data\ClipLast.txt | echo
    Get-Content $data\ClipLast.txt | Set-Clipboard
}
elseif ($args[0] -in @("save", "ыфму"))
{
    Get-Content $data\ClipLast.txt > $data\ClipSample.txt
}
else
{
    ("clip " + ($args |  ForEach-Object {
        $arg = $_ -replace "`"", '`"'
        "`"$arg`"" 
    }) -join " ") -join " " > $data\ClipLast.txt
    echo "Для выхода нажать Esc"
    $argsline = $args | ForEach-Object {"`"$_`""}
    & $program\ClipVals.ps1 @args
}