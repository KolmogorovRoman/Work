if ($args[0] -in @("set", "ыуе"))
{
    Remove-Item .\Сводные\*.mxl
    echo "Сводные удалены"
}
else
{
    $num = [string]$args[0]
    $nums = ((Get-ChildItem .\ТН\).BaseName | Select-String ("\d*"+$num+"$"))
    if ($nums.Length -eq 1)
    {
        $num = $nums[0].Matches.Value
        Remove-Item .\ТН\$num.mxl
        Remove-Item .\Счет-фактуры\$num.mxl
        Remove-Item .\CMR\$num.mxl
        echo "$num удалена"
    }
    elseif ($nums.Length -gt 1) {echo "Подходящих номеров несколько:" $nums; break}
    else {echo ([string]$num + " не найдена"); return}
}