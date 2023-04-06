$ProductsTable = @{}
if (Test-Path $data\ProductsTable.txt)
{
    (Get-Content $data\ProductsTable.txt | ConvertFrom-Json).PSobject.Properties |
        ForEach-Object { $ProductsTable[$_.Name] = $_.Value }
}
else
{
    $Excel = New-Object -ComObject Excel.Application
    if ($data -eq $null) { $data = Get-Location }
    $File_Products = "$data\ProductsTable.xlsx"
    $ProductsWorkBook = $Excel.WorkBooks.Open($File_Products)
    $ProductsWorkSheet = $ProductsWorkBook.Worksheets[1]
    $prbt = ($ProductsWorkSheet.range("a2").end(-4121).address()  -replace ('^\$', '') -split '\$')[1]
    $prrt = ($ProductsWorkSheet.range("c1").end(-4161).address()  -replace ('^\$', '') -split '\$')[0]
    $prcrds = "C", "2", $prrt, $prbt
    $ProductsBase = $ProductsWorkSheet.Range("A1:A$($prcrds[3])") |
        ForEach-Object {$_.Text}
    $ProductsCargo = $ProductsWorkSheet.Range("B1:B$($prcrds[3])") |
        ForEach-Object {$_.Text}
    $ProductsWorkSheet.Range("$($prcrds[0])$($prcrds[1]):$($prcrds[2])$($prcrds[3])") |
        Where-Object {$_.Text -ne ""} |
        ForEach-Object {
            $row = ($_.Address() -replace ('^\$', '') -split '\$')[1]
            $product_request = $_.Text -replace "\s+", "" -replace "`n", ""
            $ProductsTable[$product_request] = @{Base = $ProductsBase[+$row-1]; Cargo = $ProductsCargo[+$row-1]}
            $p = [int]((+$row)/(+$prbt)*100)
            Write-Progress -Activity "Загрузка" -Status "$p%" -PercentComplete $p
        }
    $ProductsWorkBook.Saved = $true
    $ProductsWorkBook.Close()
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ProductsWorkBook) > $null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) > $null
    $ProductsTable | ConvertTo-Json > $data\ProductsTable.txt
    Write-Progress -Activity "Загрузка" -Status "100%" -PercentComplete 100 -Completed
}

if ($args[0] -eq $null) { $File_Application = (Read-Host "Файл заявки") -replace "^`"", "" -replace "`"$", "" }
else { $File_Application = $args[0] }
$Excel = New-Object -ComObject Excel.Application
$ApplicationWorkBook = $Excel.WorkBooks.Open($File_Application)
$Excel.Visible = $true
echo "Выделите область в Excel и нажмите любую клавишу"
$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') > $null
if ($Excel -eq $null) {return}
$ApplicationWorkSheet = $Excel.ActiveSheet
$crds = $Excel.Selection.Address() -split ':' -replace ('^\$', '') -split '\$'

$Application = $crds[1]..$crds[3] |
    Where-Object {$ApplicationWorkSheet.Range("$($crds[2])$_").MergeArea(1,1).Text -notin @($null, "", "0")} |
    ForEach-Object {
        @{
            product_request = $ApplicationWorkSheet.Range("$($crds[0])$_").MergeArea(1,1).Text
            count = $ApplicationWorkSheet.Range("$($crds[2])$_").MergeArea(1,1).Text
        }
    }

$Excel.DisplayAlerts = $false
$ApplicationWorkBook.Saved = $true
$ApplicationWorkBook.Close()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ApplicationWorkBook) > $null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) > $null

$data = @()
$info = @()
$not_found = 0
echo "", ""
$Application |
    ForEach-Object {
        $product_request = $_.product_request -replace "\s+", "" -replace "`n", ""
        if ($ProductsTable[$product_request] -ne $null)
        {
            $product_base = $ProductsTable[$product_request].Base
            $cargo = $ProductsTable[$product_request].Cargo
        }
        else
        {
            $product_base = "НЕ НАЙДЕНО: $($_.product_request)"
            $cargo = "НЕ НАЙДЕНО"
            $not_found++
        }
        echo "Заявка: $($_.product_request)", "База: $product_base", $_.count, ""
        #$info += [PSCustomObject]@{База = $product_base; Заявка = $_.product_request; Число = $_.count}
        $data += $product_base, $cargo, $_.count
    }
#echo $info
echo ""

if ($not_found -gt 0) {echo "Не найдено: $not_found"}

echo "Для выхода нажать Esc"
if ($program -eq $null) { $program = Get-Location }
$vals = $data | ForEach-Object {"=$_"}
& $program\ClipVals.ps1 @vals "sEr"
if ($args[0] -ne $null) { return }