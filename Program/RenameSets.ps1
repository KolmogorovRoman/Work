$sets_info = log set dict
Push-Location .\Сводные\
Get-ChildItem . | ForEach-Object {
    $match_info = $_.BaseName | Select-String  -Pattern "(\d+)$"
    if (($match_info.Matches[0].Success) -and ($_.Extension -eq ".mxl"))
    {
        $number = $match_info.Matches[0].Groups[1].Value
        Rename-Item $_ -NewName "$($sets_info[$number].Клиент) $number.mxl"}
    }
Pop-Location