$num = [string]$args[0]
$nums = ((Get-ChildItem .\ТН\).BaseName | Select-String ("\d*"+$num+"$"))
if ($nums.Length -eq 1)
{
    $num = $nums[0].Matches.Value
}
elseif ($nums.Length -gt 1) {echo "Подходящих номеров несколько:" $nums; return}
else
{
    echo ([string]$num + " не найдена")
    $Global:open_next = ([int]$num+1)
    return
}
if ($num -in $Global:skip_numbers)
{
    echo ([string]$num + " уже пропущена")    
}
else
{
    $Global:skip_numbers.Add($num) > $null
    echo ([string]$num + " пропущена")
}