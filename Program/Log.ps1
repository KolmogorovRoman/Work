$quals = @()
$cmrs = @()
$waybills = @()
$arg = $args[0]
$places = @{}
$gates = @{}
Import-Csv C:\Users\KolmogorovRV\Work\Data\Places.csv -UseCulture -Encoding Default | ForEach-Object {
    if ($places.ContainsKey($_.Adress))
    {
	$places[$_.Adress] = @($places[$_.Adress]) + $_.Name
	$gates[$_.Adress] = $_.Gate
    }
    else
    {
	$places[$_.Adress] = $_.Name
	$gates[$_.Adress] = $_.Gate
    }
}
$products = Import-Csv C:\Users\KolmogorovRV\Work\Data\Products.csv -UseCulture -Encoding Default
#$places = & $data\Places.ps1
if ($arg -in @("set", "ыуе"))
{
    $sets = @{}
    $set_numbers = @{}
    $set_places = @{}
    $set_drivers = @{}
    Get-ChildItem .\ТН\ | ForEach-Object {
        $number = ""
        $match_info = Select-String $_ -Pattern '"#","(\d{7})"'
        $number = $match_info.Matches[0].Groups[1].Value

        $number_set = ""
        $match_info = Select-String $_ -Pattern "Свод № (\d+)"
        $number_set = $match_info.Matches[0].Groups[1].Value

        $driver = ""
        $match_info = Select-String $_ -Pattern "Товар к доставке принял:.+\s+(\w+\s+\w\.(\w\.)?)"
        $driver = $match_info.Matches[0].Groups[1].Value -replace "\s+", " "

        $place = ""
        $match_info = Select-String .\Счет-фактуры\$number.mxl -Pattern "Пункт разгрузки:\s+(.+)\s*`"}"
        $adress = $match_info.Matches[0].Groups[1].Value.Trim() -replace "`"`"", "`""
        $place = $places[$adress]
        if ($place -is [Array]) {$place = $place | Where-Object {Select-String .\Счет-фактуры\$number.mxl -Pattern $_}}

        if ($number_set -ne "") {$set_numbers[$number_set] = $number_set}
        if ($place -ne "") {$set_places[$number_set] = $place}
        if ($driver -ne "") {$set_drivers[$number_set] = $driver}
    }
    $set_numbers.Keys | ForEach-Object {
        $sets[$_] = [PSCustomObject]@{
            Номер = $_;
            Водитель = $set_drivers[$_];
            Клиент = $set_places[$_];
            Файл = Get-ChildItem .\Сводные\*$_.mxl | ForEach-Object {"$($_.BaseName)$($_.Extension)"}
        }
    }
    if ($args[1] -eq "dict") { return $sets }
    else
    {
        Push-Location .\Сводные\
        Get-ChildItem . | ForEach-Object {
            $match_info = $_.BaseName | Select-String  -Pattern "(\d+)$"
            if (($match_info.Matches[0].Success) -and ($_.Extension -eq ".mxl"))
            {
                $number = $match_info.Matches[0].Groups[1].Value
                Rename-Item $_ -NewName "$($sets[$number].Клиент) $number.mxl"}
                $sets[$number].Файл = "$($sets[$number].Клиент) $number.mxl"
            }
        Pop-Location
        return $sets.Values | Sort-Object Водитель, Номер
    }
}
Get-ChildItem .\ТН\ | ForEach-Object {
    $by_bill = $False
    $number_qual = ""
    $match_info = Select-String $_ -Pattern "Качественное удостоверение № (\d+)"
    $number_qual = +$match_info.Matches[0].Groups[1].Value

    $number = ""
    $match_info = Select-String $_ -Pattern '"#","(\d{7})"'
    $number = $match_info.Matches[0].Groups[1].Value
    if ($number -eq "")
	{
	    $match_info = Select-String $_ -Pattern "Счет-фактура № (\d{7})"
        $number = $match_info.Matches[0].Groups[1].Value
        $by_bill = $True
	}

    $number_set = ""
    $match_info = Select-String $_ -Pattern "Свод № (\d+)"
    $number_set = $match_info.Matches[0].Groups[1].Value

    $place = ""
    $match_info = Select-String .\Счет-фактуры\$number.mxl -Pattern "Пункт разгрузки:\s+(.+)\s*`"}"
    $adress = $match_info.Matches[0].Groups[1].Value.Trim() -replace '""', '"'
    $place = $places[$adress]
    if ($place -is [Array]) {$place = $place | Where-Object {Select-String .\Счет-фактуры\$number.mxl -Pattern $_}}

    $number_bill = ""
    $match_info = Select-String $_ -Pattern "Счет-фактура № (\d{7})"
    $number_bill = $match_info.Matches[0].Groups[1].Value

    $number_cmr = ""
    $match_info = Select-String $_ -Pattern "CMR № (\d+)"
    $number_cmr = $match_info.Matches[0].Groups[1].Value

    $cargo_count_text = ""
    $match_info = Select-String $_ -Pattern "Всего количество грузовых мест:\s*(.+)\s*`"}"
    $cargo_count_text = $match_info.Matches[0].Groups[1].Value
    $cargo_count = "$(& $program\StringToNumber.ps1 $cargo_count_text)"

    $mass_gross = ""
    $match_info = Select-String .\Счет-фактуры\$number.mxl -Pattern "Вес БРУТТО:\s*(.+)\s*кг"
    $mass_gross = $match_info.Matches[0].Groups[1].Value

    $mass_net = ""
    $match_info = Select-String .\Счет-фактуры\$number.mxl -Pattern "Вес НЕТТО:\s*(.+)\s*кг"
    $mass_net = $match_info.Matches[0].Groups[1].Value

    $organisation = ""
    if ((Select-String $_ -Pattern "Производственное подразделение в г. Витебске").Matches.Success) {$organisation += "Вит"}
    if ((Select-String $_ -Pattern "Производственный цех г. Новолукомль").Matches.Success) {$organisation += "Нов"}
    if ((Select-String $_ -Pattern "Производственный цех г.п. Шумилино").Matches.Success) {$organisation += "Шум"}

    $product = New-Object System.Collections.Generic.HashSet[string]
    $product_group = New-Object System.Collections.Generic.HashSet[string]
    $product_org = New-Object System.Collections.Generic.HashSet[string]
    $file = $_
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

    $driver = ""
    $match_info = Select-String $_ -Pattern "Товар к доставке принял:.+\s+(\w+\s+\w\.(\w\.)?)"
    $driver = $match_info.Matches[0].Groups[1].Value -replace "\s+", " "

    $match_info = Select-String $_ -Pattern "Второй водитель.*\s+(\w+\s+\w\.\w\.)"
    $second_driver = ""
    if ($match_info.Matches.Success) {$second_driver = $match_info.Matches[0].Groups[1].Value -replace "\s+", " "}
    else
    {
        #$match_info = Select-String $_ -Pattern "Второй водитель(\w+\s+\w+\s+\w+)"
	$match_info = Select-String $_ -Pattern "Второй водитель(.+)\s_+"
        if ($match_info.Matches.Success) {$second_driver = $match_info.Matches[0].Groups[1].Value -replace "\s+", " "}
    }

    $border = ""
    $match_info = Select-String .\CMR\$number.mxl -Pattern 'Через ППУ ""(.+)"""'
    $border = $match_info.Matches[0].Groups[1].Value

    $state = ""
    if ($by_bill) {$state += "Б/Н"}
    if ($number -ne $number_bill) {$state += "Счет не совп."}
    if ($quals -contains $number_qual) {$state += "Кач. повт. "}
    if ($cmrs -contains $number_cmr) {$state += "CMR повт. "}
    if ((Get-ChildItem .\Счет-фактуры\).BaseName -notcontains $number) {$state += "Счета нет "}
    if ((Get-ChildItem .\CMR\).BaseName -notcontains $number) {$state += "CMR нет "}
    $product_org.Add($organisation) > $null
    if ($product_org.Count -gt 1) {$state += "Цех не совп."}
    if ($organisation -eq 'Вит') {$organisation = '   '}
    $gate_by_place = $gates[$adress]
    # if (-not (Select-String .\CMR\$number.mxl -Pattern $gate_by_place -Quiet)) {$state += "Граница"}
    if ($number -in $Global:skip_numbers) {$state += "Пропущена"}

    $waybill = [PSCustomObject]@{
        Статус = $state;
        КАЧ = $number_qual;
        Клиент = $place;
        Цех = $organisation;
        Продукция = $product;
        Номер = $number;
        CMR = $number_cmr;
        Водитель = $driver;
        "Второй водитель" = $second_driver;
	ППУ = $border
        Мест = $cargo_count;
        Масса = $mass_gross;
        Нетто = $mass_net;
        Свод = $number_set
    }

    $quals = $quals + $number_qual
    $cmrs = $cmrs + $number_cmr
    $waybills = $waybills + $waybill
}
if ($arg -in @("vet", "муе"))
{
    $Text = $waybills | Sort-Object -Property Водитель, КАЧ | Format-Table -Property Водитель, КАЧ, Цех, Продукция, Клиент, Номер, CMR, Мест, Нетто | Out-String
    $Word = New-Object -ComObject Word.Application
    $Doc = $Word.Documents.Add()
    $Word.Selection.ParagraphFormat.SpaceAfter = 0
    $Word.Selection.ParagraphFormat.LineSpacing = 12
    $Word.Selection.Font.Name = "Consolas"

    if ($args[1] -in @("hor", "рщк"))
    {
        $Word.Selection.Font.Size = 14
        $Doc.PageSetup.Orientation = 1
        $Doc.PageSetup.LeftMargin = 100
    }
    else
    {
        $Word.Selection.Font.Size = 11
        $Doc.PageSetup.Orientation = 0
        $Doc.PageSetup.LeftMargin = 60
    }
    $Doc.PageSetup.RightMargin = 20
    $Doc.PageSetup.TopMargin = 20
    $Doc.PageSetup.BottomMargin = 20

    $Word.Selection.TypeText($name)
    $Word.Selection.TypeText($Text)

    $Doc.Saved = $true
    $Word.Visible = $true
    $doc.ExportAsFixedFormat("$(Get-Location)\$name.pdf", 17)
}
else
{
    $waybills | Sort-Object -Property Водитель, КАЧ | Format-Table -Property Статус, КАЧ, Клиент, Цех, Номер, CMR, Водитель, "Второй водитель", ППУ, Масса
}