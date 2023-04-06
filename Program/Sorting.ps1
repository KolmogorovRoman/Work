$work = [string]$args[0]
$name = [string]$args[1]
$dest = "$work\$name"
$program = "$work\Program"
$input = "$work\Input"
$watcher = New-Object System.IO.FileSystemWatcher $input
$host.ui.RawUI.WindowTitle = "Sorting $name"
Remove-Item $input\*
echo "Сортировка запущена: $input -> $dest"
echo ""
while ($true)
{
    $result = $watcher.WaitForChanged([System.IO.WatcherChangeTypes]::Created)
    $name = $result.name
    Write-Host $name -NoNewLine
    $text = $null
    while ($text -eq $null)
    {
        try { $text = Get-Content $input\$name -ErrorAction Stop }
        catch { sleep -m 200 }
    }
    $number = $null
    if ((Select-String $input\$name -Pattern "С товаром переданы документы").Matches.Success)
    {
        $match_info = Select-String $input\$name -Pattern '"#","(\d{7})"'
        $number = $match_info.Matches[0].Groups[1].Value
	if ($number -eq $null)
	{
	    $match_info = Select-String $input\$name -Pattern "Счет-фактура № (\d{7})"
            $number = $match_info.Matches[0].Groups[1].Value
	}
        Copy-Item $input\$name -Destination $dest\ТН\$number.mxl
        Remove-Item $input\$name
        echo " -> \$dest\ТН\$number.mxl"
    }
    elseif ((Select-String $input\$name -Pattern "СЧЕТ-ФАКТУРА № \d{7} от").Matches.Success)
    {
        $match_info = Select-String $input\$name -Pattern "СЧЕТ-ФАКТУРА № (\d{7})"
        $number = $match_info.Matches[0].Groups[1].Value
        Copy-Item $input\$name -Destination $dest\Счет-фактуры\$number.mxl
        Remove-Item $input\$name
        echo " -> \$dest\Счет-фактуры\$number.mxl"
    }
    elseif ((Select-String $input\$name -Pattern "Internationaler Frachtbrif").Matches.Success)
    {
        $match_info = Select-String $input\$name -Pattern "ТН № \w{2}(\d{7})"
        $number = $match_info.Matches[0].Groups[1].Value
	if ($number -eq $null)
	{
	    $match_info = Select-String $input\$name -Pattern "Счет-фактура № (\d{7})"
            $number = $match_info.Matches[0].Groups[1].Value
	}
        Copy-Item $input\$name -Destination $dest\CMR\$number.mxl
        Remove-Item $input\$name
        echo " -> \$dest\CMR\$number.mxl"
    }
    elseif ((Select-String $input\$name -Pattern "Сводное задание на отгрузку").Matches.Success)
    {
        $match_info = Select-String $input\$name -Pattern "Сводное задание на отгрузку № (\d+)"
        $number = $match_info.Matches[0].Groups[1].Value
        Copy-Item $input\$name -Destination $dest\Сводные\$number.mxl
        Remove-Item $input\$name
        echo " -> \$dest\Сводные\$number.mxl"
    }
    else
    {
        echo " не определен"
    }
}