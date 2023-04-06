$products = Import-Csv C:\Users\KolmogorovRV\Work\Data\Products.csv -UseCulture -Encoding Default
if ($args[0] -in @("set", "ыуе"))
{
    $sets_info = log set dict
    Push-Location .\Сводные
    $sets = (Get-ChildItem . | Sort-Object -Property {
        $match_info = $_.BaseName | Select-String -Pattern "(\d+)$"
        $number = $match_info.Matches[0].Groups[1].Value
        return $sets_info[$number].Водитель, $sets_info[$number].Номер
    })
    [array]::reverse($sets)

    & $program\1cv8fv\bin\1cv8fv.exe ($sets)
    Pop-Location
}
elseif ($args[0] -in @("bill", "ишдд"))
{
    Push-Location .\Счет-Фактуры
    & $program\1cv8fv\bin\1cv8fv.exe (Get-ChildItem . | Sort-Object -Descending) 
    Pop-Location
}
elseif ($args[0] -in @("cmr", "ськ", "цмр", "смр"))
{
    Push-Location .\CMR
    & $program\1cv8fv\bin\1cv8fv.exe (Get-ChildItem . | Sort-Object -Descending) 
    Pop-Location
}
else
{
    $num = [string]$args[0]
    if ($num -eq "")
    {
        $num = $Global:open_next
    }
    $nums = ((Get-ChildItem .\ТН\).BaseName | Select-String ("\d*"+$num+"$"))
    if ($nums.Length -eq 1)
    {
        $num = $nums[0].Matches.Value
        $Global:open_next = ([int]$num+1)
    }
    elseif ($nums.Length -gt 1) {echo "Подходящих номеров несколько:" $nums; return}
    else
    {
        echo ([string]$num + " не найдена")
        $Global:open_next = ([int]$num+1)
        return
    }

    $organisation = ""
    if ((Select-String .\ТН\$num.mxl -Pattern "открытое акционерное общество").Matches.Success) {$organisation += "   "}
    if ((Select-String .\ТН\$num.mxl -Pattern "Производственный цех г.Новолукомль").Matches.Success) {$organisation += "Нов"}
    if ((Select-String .\ТН\$num.mxl -Pattern "производственный цех г.п.Шумилино").Matches.Success) {$organisation += "Шум"}

    echo ($num, $organisation -join " ")
    if ($num -in $Global:skip_numbers)
    {
        echo "Пропущена"
        $Global:open_next = ([int]$num+1)
        return
    }
    $match_info = Select-String .\ТН\$num.mxl -Pattern '"#","(\d{7})"'
    if ($match_info.Matches.Count -eq 0)
    {
        echo "Без номера"
        $Global:open_next = ([int]$num+1)
        return
    }
     
    $product = New-Object System.Collections.Generic.HashSet[string]
    $product_group = New-Object System.Collections.Generic.HashSet[string]
    $product_org = New-Object System.Collections.Generic.HashSet[string]
    $file = ".\ТН\$num.mxl"
    $products | ForEach-Object {
        if ((Select-String $file -Pattern $_.Search).Matches.Success) 
            {
                $product.Add($_.Name) > $null;
                $product_group.Add($_.Group) > $null;
                $product_org.Add($_.Org) > $null
            }
    }
    if ($product.Count -eq 1) {$product = (new-Object System.Collections.Generic.List[string] -ArgumentList $product)[0]}
    elseif ($product_group.Count -eq 1) {$product = (new-Object System.Collections.Generic.List[string] -ArgumentList $product_group)[0]}
    echo $product

    if ((Select-String .\ТН\$num.mxl -Pattern "Товар согласно приложению").Matches.Success)
    {
        echo "Большая накладная!"
        & $program\1cv8fv\bin\1cv8fv.exe .\CMR\$num.mxl .\Счет-фактуры\$num.mxl .\ТН\$num.mxl "$(Get-Location)\ТН\$num.mxl"
    }
    else
    {
        & $program\1cv8fv\bin\1cv8fv.exe .\CMR\$num.mxl .\Счет-фактуры\$num.mxl .\ТН\$num.mxl
    }
}