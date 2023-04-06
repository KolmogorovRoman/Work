$str = (([string]::Join(" ", $args)).Split() | Out-String -Stream).ToLower()
$number_unknown = 0
$number = 0
$str | ForEach-Object {
    if ($_ -in @("один", "одна")) { $number_unknown += 1 }
    elseif ($_ -in @("два", "две")) { $number_unknown += 2 }
    elseif ($_ -in @("три")) { $number_unknown += 3 }
    elseif ($_ -in @("четыре")) { $number_unknown += 4 }
    elseif ($_ -in @("пять")) { $number_unknown += 5 }
    elseif ($_ -in @("шесть")) { $number_unknown += 6 }
    elseif ($_ -in @("семь")) { $number_unknown += 7 }
    elseif ($_ -in @("восемь")) { $number_unknown += 8 }
    elseif ($_ -in @("девять")) { $number_unknown += 9 }
    elseif ($_ -in @("десять")) { $number_unknown += 10 }
    elseif ($_ -in @("одиннадцать")) { $number_unknown += 11 }
    elseif ($_ -in @("двенадцать")) { $number_unknown += 12 }
    elseif ($_ -in @("тринадцать")) { $number_unknown += 13 }
    elseif ($_ -in @("четырнадцать")) { $number_unknown += 14 }
    elseif ($_ -in @("пятнадцать")) { $number_unknown += 15 }
    elseif ($_ -in @("шестнадцать")) { $number_unknown += 16 }
    elseif ($_ -in @("семнадцать")) { $number_unknown += 17 }
    elseif ($_ -in @("восемнадцать")) { $number_unknown += 18 }
    elseif ($_ -in @("девятнадцать")) { $number_unknown += 19 }
    elseif ($_ -in @("двадцать")) { $number_unknown += 20 }
    elseif ($_ -in @("тридцать")) { $number_unknown += 30 }
    elseif ($_ -in @("сорок")) { $number_unknown += 40 }
    elseif ($_ -in @("пятьдесят")) { $number_unknown += 50 }
    elseif ($_ -in @("шестьдесят")) { $number_unknown += 60 }
    elseif ($_ -in @("семьдесят")) { $number_unknown += 70 }
    elseif ($_ -in @("восемьдесят")) { $number_unknown += 80 }
    elseif ($_ -in @("девяносто")) { $number_unknown += 90 }
    elseif ($_ -in @("сто")) { $number_unknown += 100 }
    elseif ($_ -in @("двести")) { $number_unknown += 200 }
    elseif ($_ -in @("триста")) { $number_unknown += 300 }
    elseif ($_ -in @("четыреста")) { $number_unknown += 400 }
    elseif ($_ -in @("пятьсот")) { $number_unknown += 500 }
    elseif ($_ -in @("шестьсот")) { $number_unknown += 600 }
    elseif ($_ -in @("семьсот")) { $number_unknown += 700 }
    elseif ($_ -in @("восемьсот")) { $number_unknown += 800 }
    elseif ($_ -in @("девятьсот")) { $number_unknown += 900 }
    elseif ($_ -in @("тысяча", "тысячи", "тысяч"))
    {
        $number += $number_unknown * 1000
        $number_unknown = 0
    }
    elseif ($_ -in @("миллион", "миллиона", "миллионов"))
    {
        $number += $number_unknown * 1000000
        $number_unknown = 0
    }
    elseif ($_ -in @("миллиард", "миллиарда", "миллиардов"))
    {
        $number += $number_unknown * 1000000000
        $number_unknown = 0
    }
    elseif ([string]::IsNullOrWhiteSpace($_))
    {}
    else
    {
        $number += 42000000000000
    }
}
$number += $number_unknown
$number