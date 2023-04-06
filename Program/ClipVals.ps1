function new_line
{
    Write-Host ""
}
function out_line
{
    [Console]::SetCursorPosition(0, $args[0].start)
    if ($args[1] -eq "high")
    {
        $args[0].vals | Select-Object -SkipLast 1 | ForEach-Object {"$_ "} | Write-Host -NoNewline
        Write-Host $args[0].vals[-1] -NoNewline -ForegroundColor $host.ui.RawUI.BackgroundColor -BackgroundColor $host.ui.RawUI.ForegroundColor
    }
    else
    {
        $args[0].vals -join " " | Write-Host -NoNewline
    }
}
function reout_line
{
    $CurrentLine = $Host.UI.RawUI.CursorPosition.Y
    $ConsoleWidth = $Host.UI.RawUI.BufferSize.Width
    $args[0].start..$args[0].end | ForEach-Object {
        [Console]::SetCursorPosition(0, $_)
        Write-Host (" " * $ConsoleWidth ) -NoNewline
    }
    [Console]::SetCursorPosition(0, $args[0].start)
    out_line $args[0] $args[1]
}

$each_newline = $false
$hwnd = (Get-Process -Id $PID).MainWindowHandle
$vals = @()
$true_vals = @()
$keys = @()
$settings = "s"
$last_str = "lКонец"
$args | ForEach-Object {
    if ($_[0] -eq "w")
    {
        $hwnd = $_.Substring(1)
    }
    elseif ($_[0] -eq "s")
    {
        $each_newline = "e" -in $_.Substring(1).ToCharArray()
        $settings = $_
    }
    elseif ($_[0] -eq "l")
    {
        $last_str = $_.Substring(1)
    }
    elseif ($_[0] -in @("n", "p", "e"))
    {
        $keys += $_
    }
    else
    {
        $vals += ($_ -replace "`"", "\`"")
        $true_vals += $_
    }
}
if ($keys.Count -eq 0)
{
    #$keys = @("N86", "P67", "e27")
    $keys = @("N86", "P20", "e27")
}

if ("A" -notin $settings.ToCharArray())
{
    if ("N86" -notin $keys)
    {
        $settings += "a"
    }
}

$lines = New-Object System.Collections.ArrayList
$l = @{
    vals = New-Object System.Collections.ArrayList
    start = $Host.UI.RawUI.CursorPosition.Y
    end = $Host.UI.RawUI.CursorPosition.Y
    }
$lines.Add($l) > $null
$WindowTitle = $host.UI.RawUI.WindowTitle
$ClipboardContent = Get-Clipboard

& $program\Clip.exe @vals @keys "w$hwnd" $last_str $settings | ForEach-Object {
    $out = $_.Split(" ", 4)
    $val = $true_vals[$out[1]]
    if ($val[0] -eq "=") {$val = $val.Substring(1)}
    elseif ($val[0] -eq "+")
    {
        $length = $val.Substring(1).Length
        $val = [string] (+$val.Substring(1) + $out[2])
        $zeroes_count = $length - $val.Length
        if ($zeroes_count -lt 0) {$zeroes_count = 0}
        $val = "0" * ($zeroes_count) + $val
    }
    $add_new_line = $false
    if ($each_newline) 
    {
        if ($out[0] -eq "V" -and $out[2] -ne "0") { $add_new_line = $true }
        if ($out[0] -eq ">") {$out[0] = "V"}
        if ($out[0] -eq "<") {$out[0] = "^"}
    }
    if ($out[0] -eq "V")
    {
        reout_line $lines[-1]
        new_line
        if ($add_new_line) { new_line }

        $l = @{
            vals = New-Object System.Collections.ArrayList
            start = $Host.UI.RawUI.CursorPosition.Y
            }
        $l.vals.Add($val) > $null
        $lines.Add($l) > $null
        out_line $lines[-1] "high"
        $lines[-1].end = $Host.UI.RawUI.CursorPosition.Y
    }
    elseif ($out[0] -eq ">")
    {
        $lines[-1].vals.Add($val) > $null
        reout_line $lines[-1] "high"
        $lines[-1].end = $Host.UI.RawUI.CursorPosition.Y
    } 
    elseif ($out[0] -eq "<")
    {
        $lines[-1].vals.RemoveAt($lines[-1].vals.Count - 1)
        reout_line $lines[-1] "high"
        $lines[-1].end = $Host.UI.RawUI.CursorPosition.Y
    } 
    elseif ($out[0] -eq "^")
    {
        $lines[-1].vals.RemoveAt($lines[-1].vals.Count - 1)
        reout_line $lines[-1]
        $lines.RemoveAt($lines.Count - 1)
        reout_line $lines[-1] "high"
    }
    elseif ($out[0] -eq ".")
    {
        reout_line $lines[-1]
        new_line
        if ($add_new_line) { new_line }

        $l = @{
            vals = New-Object System.Collections.ArrayList
            start = $Host.UI.RawUI.CursorPosition.Y
            }
        $l.vals.Add("Конец") > $null
        $lines.Add($l) > $null
        out_line $lines[-1] "high"
        $lines[-1].end = $Host.UI.RawUI.CursorPosition.Y
    }
    else
    {
        echo $_
    }

    $host.UI.RawUI.WindowTitle = $lines[-1].vals[-1]
}

reout_line $lines[-1]
new_line
new_line
$host.UI.RawUI.WindowTitle = $WindowTitle
Set-Clipboard $ClipboardContent