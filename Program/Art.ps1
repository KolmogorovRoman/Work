$width = $Host.UI.RawUI.WindowSize.Width
$height = $Host.UI.RawUI.WindowSize.Height
$snowflakes1 = @()
$snowflakes2 = @()
1..[int]($width*$height/75) | ForEach-Object {
    $snowflakes1 += [PSCustomObject]@{
        x = (Get-Random -Maximum $width)
        y = (Get-Random -Maximum ($height))
    }
}
1..[int]($width*$height/100) | ForEach-Object {
    $snowflakes2 += [PSCustomObject]@{
        x = (Get-Random -Maximum $width)
        y = (Get-Random -Maximum ($height*2))
    }
}
$light = New-Object int[][] $height, $width

$e=[char]27

function draw_snow
{
    #$screen = $args[0]
    $lines = @((1..($height)) |
        ForEach-Object {
            New-Object System.Text.StringBuilder (" " * $width)
        })
    $snowflakes2 | ForEach-Object {
        if ($_.y % 2 -eq 0)
        {
            if ($lines[[math]::floor($_.y/2)][$_.x] -eq ".")
            {
                $lines[[math]::floor($_.y/2)][$_.x] = ":"
            }
            else
            {
                $lines[[math]::floor($_.y/2)][$_.x] = "'"
            }
        }
        else
        {
            if ($lines[[math]::floor($_.y/2)][$_.x] -eq "'")
            {
                $lines[[math]::floor($_.y/2)][$_.x] = ":"
            }
            else
            {
                $lines[[math]::floor($_.y/2)][$_.x] = "."
            }
        }
    }
    $snowflakes1 | ForEach-Object {
        $lines[$_.y][$_.x] = "*"
    }
    $str = ""
    #$snowflakes1 | ForEach-Object {
        #$str += "$e[38;5;$($screen[$_.y][$_.x])m$e[$($_.y + 1);$( $_.x + 1 )H*"
        #$str += "$e[$($_.y + 1);$( $_.x + 1 )H*"
    #}
    #$snowflakes2 | ForEach-Object {
        #$c = $lines[[math]::floor($_.y/2)][$_.x]
        #$str += "$e[38;5;$($screen[[math]::floor($_.y/2)][$_.x])m$e[$([math]::floor($_.y/2) + 1);$( $_.x + 1 )H$($lines[[math]::floor($_.y/2)][$_.x])"
        #$str += "$e[$([math]::floor($_.y/2) + 1);$( $_.x + 1 )H$($lines[[math]::floor($_.y/2)][$_.x])"
    #}
    $lines = $lines  | ForEach-Object {$_.ToString()}
    return ($lines -join "$e[1E")
    #return $str
}
function snow_steep1
{
    $snowflakes1 | ForEach-Object {
        $xrandom = Get-Random -Maximum 100
        $yrandom = Get-Random -Maximum 100
        if ($yrandom -le 90)
        {
            if (($_.y -eq ($height-1)) -and ($yrandom -le 0)) {}
            else
            {
                $_.y++
                if ($xrandom -le 80) { $_.x-- }
                elseif ($xrandom -le 90) { $_.x++ }
            }
        }
        if ($_.y -ge ($height)) { $_.y = 0 }
        if ($_.x -ge ($width)) { $_.x = 0 }
        if ($_.x -le 0) { $_.x = $width - 1 }
    }
}
function snow_steep2
{
    $snowflakes2 | ForEach-Object {
        $xrandom = Get-Random -Maximum 100
        $yrandom = Get-Random -Maximum 100
        if ($yrandom -le 90)
        {
            $_.y++
            if ($xrandom -le 80) { $_.x-- }
            elseif ($xrandom -le 90) { $_.x++ }
        }
        if ($_.y -ge ($height * 2)) { $_.y = 0 }
        if ($_.x -ge ($width)) { $_.x = 0 }
        if ($_.x -le 0) { $_.x = $width - 1 }
    }
}
function make_sprite
{
    return $args[0].Split("`n") |
        Select-Object -Skip 1 | Select-Object -SkipLast 1 |
        ForEach-Object {  
            if (-not [string]::IsNullOrWhiteSpace($_)) {[PSCustomObject]@{
                start = ( $_ | Select-String "^\s*(\S)" ).Matches[0].Groups[1].Index
                end  =  ( $_ | Select-String "\S(\s*)$" ).Matches[0].Groups[1].Index
                line = $_
                }
            }
            else {[PSCustomObject]@{
                start = 0
                end  =  0
                line = ""
                }
            }
        }
}
function draw_sprite
{
    $x = $args[1]
    $offset_left = 0
    if ($x -lt 0) { $offset_left = -$x }
    $y = $args[2]
    $c = $args[3]
    if ($c -eq $null) {$c = 15}
    $str = ""
    $args[0] | ForEach-Object {
        $skip = $offset_left - $_.start
        if ($skip -lt 0) { $skip = 0 }
        $trim = 0
        if ($x - 1 + $_.line.Length -gt $width) {$trim = $x - 1 + $_.line.Length - $width}
        #$str += "$e[38;5;$($c)m$e[$($y + 1);$( $x + $_.start + 1 + $offset )H$($_.line.Substring($skip))"
        if (($skip + $trim) -lt ($_.end - $_.start))
        {
            $str += "$e[38;5;$($c)m$e[$($y + 1);$( $x + $_.start + 1 + $skip )H$($_.line.Substring($_.start + $skip, $_.end - $_.start - $skip - $trim))"
        }
        $y++
    }
    return $str
}
function make_light
{
    $colors = $args[1]
    return $args[0].Split("`n") |
        Select-Object -Skip 1 | Select-Object -SkipLast 1 |
        ForEach-Object {
            $str = $_
            $line = [PSCustomObject]@{
                start = 0
                end  =  0
                colors = $null
                }
            $line.start = ( $_ | Select-String "^\s*(\S)" ).Matches[0].Groups[1].Index
            $line.end  =  ( $_ | Select-String "\S(\s*)$" ).Matches[0].Groups[1].Index
            $line.colors = New-Object int[] ($line.end - $line.start)
            (0 .. ($line.end - $line.start - 1)) | ForEach-Object {
                $line.colors[$_] = $colors[[string]($str[$line.start+$_])]
                #$line.colors[$_] = $colors[[string]($str[4])]
            }
            return $line
        }
}
function draw_light
{
    $light = $args[1]
    $x = $args[2]
    $y = $args[3]
    $args[0] | ForEach-Object {
        $line = $_
        (0 .. ($_.end - $_.start-1)) | ForEach-Object {
            ($light[$y][$x + $line.start + $_]) = ($line.colors[$_])
        }
        $y++
    }
}
$star = make_sprite "
     .
  __/ \__
  \     /
  /_.^._\
"
$star_light = make_light "
      111 
    1111111 
  11111 11111 
 111       111 
1111       1111 
1111       1111 
 111       111 
" @{"1"=210}
$truck_light = make_light "
             111111111
           111111111111
         111111111111111
       111111111111111111
     111111111111111111111
   11111111111111111111111
 111111111111111111111111
111111111111111111111111
 1111111111111111111111
   1111111111111111111
" @{"1"=210}
$tree = make_sprite "
          /   \
         /     \
        /       \
       /         \
      /_         _\
       /*-.._..-*\
      /           \
     /             \
    /               \
   /_               _\
    /*--...___...--*\
   /                 \
  /                   \
 /                     \
/_                     _\
  *---...,_____,...---*
           | |
"
$guirland = make_sprite "
             _,*
         _,-*    
        *
           
                _,
            _,-*   
        _,-*        
     ,-*            
             
                 _,-*
             _,-*    
         _,-*          
     _,-*
    *
"
$balls = "
          1
               2
         
           
            3
              
       4
                1
             
       
      2            4
    
                3
"
$balls_1 = make_sprite ($balls -replace "1", "0" -replace "2", " " -replace "3", " " -replace "4", " ")
$balls_2 = make_sprite ($balls -replace "1", " " -replace "2", "0" -replace "3", " " -replace "4", " ")
$balls_3 = make_sprite ($balls -replace "1", " " -replace "2", " " -replace "3", "0" -replace "4", " ")
$balls_4 = make_sprite ($balls -replace "1", " " -replace "2", " " -replace "3", " " -replace "4", "0")
$truck = make_sprite "
 ______________________________   ______
|                              |  |  |  \
|         М О Л О К О          |  |__|___\
|                              |  |  |    |
|                              |__|  |    |
|______________________________|__|_______|

"
$wheels = @(
(make_sprite "
  (x)(+)                (x)(+)        (x)
"),
(make_sprite "
  (+)(x)                (+)(x)        (+)
"))
$t = 0
$ball1_colors = @(54, 207)
$ball3_colors = @(94, 226)
$ball2_colors = @(28, 49)
$ball4_colors = @(88, 196)
$ball_modes = @(
@(0, 1, 1, 1, 1, 0, 0, 0),
@(0, 0, 1, 1, 1, 1, 0, 0),
@(0, 0, 0, 1, 1, 1, 1, 0),
@(0, 0, 0, 0, 1, 1, 1, 1)
)
#$ball1_mode = @(1, 1, 0, 0, 0, 1, 1, 1)
#$ball2_mode = @(1, 1, 1, 1, 0, 0, 0, 1)
#$ball3_mode = @(0, 1, 1, 1, 1, 1, 0, 0)
#$ball4_mode = @(0, 0, 0, 1, 1, 1, 1, 1)
$guirland_colors = @(243, 158)
$star_colors = @(88, 9)

if ($args[0] -eq "colors")
{
    Write-Output ((0..7|%{"$e[38;5;$($_)m$_"}) -join " ")
    Write-Output ((8..15|%{"$e[38;5;$($_)m$_"}) -join " ")
    0..5 | ForEach-Object {
            $c = $_
            (0..5) | ForEach-Object {
                Write-Output (((16+$c*36+$_*6)..(16+$c*36+($_+1)*6-1) | ForEach-Object {
                    $s = "0"*(3-$_.ToString().Length)+$_
                    "$e[38;5;$($_)m$s"
                }) -join " ")
            }
        }
    Write-Output ((232..255|%{"$e[38;5;$($_)m$_"}) -join " ")
    return
}


Write-Output "$e[?25l"
while ($true)
{
    $str = "$e[0m$e[1;1H"
    #$lines = @((1..($height)) |
    #    ForEach-Object {
    #        New-Object System.Text.StringBuilder (" " * $width)
    #    })
    #$str += ($lines -join "$e[1E")
    
    $time = Measure-Command {
        if ($t % 2 -eq 0)
        {
            snow_steep2
            snow_steep1
        }
        $x = [int]($width/2) - ([int]($tree[-3].line.Length/2))
        $y = $height - $tree.Count
        
        #$light = New-Object int[][] $height, $width
        #(0..($height-1)) | ForEach-Object {
        #    $_y = $_
        #    (0..($width-1)) | ForEach-Object {
        #        $light[$_y][$_] = 15
        #    }
        #}
        #if ([math]::floor($t/24) % 2 -eq 1)
        #{
        #    draw_light $star_light $light ($x+5) ($y-6)
        #}
        $str += draw_snow #$light
        $t++
        $truck_x = $t % ($width+200) - 100
        $str += draw_sprite $truck ($truck_x) ($height - $truck.Count) 15
        $str += draw_sprite $wheels[($truck_x/1) % 2] ($truck_x) ($height - $wheels1.Count) 15

        $str += draw_sprite $star ($x+7) ($y-4) $($star_colors[[math]::floor($t/22) % 2])
        $str += draw_sprite $tree $x $y 34
        $str += draw_sprite $guirland $x ($y+2) $($guirland_colors[[math]::floor($t/22) % 2])
        if (([math]::floor($t/9) % 8) -eq 0)
        {
            $ball_modes = Get-Random $ball_modes -Count $ball_modes.Count
        }
        $str += draw_sprite $balls_1 $x ($y+2) $($ball1_colors[$ball_modes[0][[math]::floor($t/9) % 8]])
        $str += draw_sprite $balls_2 $x ($y+2) $($ball2_colors[$ball_modes[1][[math]::floor($t/9) % 8]])
        $str += draw_sprite $balls_3 $x ($y+2) $($ball3_colors[$ball_modes[2][[math]::floor($t/9) % 8]])
        $str += draw_sprite $balls_4 $x ($y+2) $($ball4_colors[$ball_modes[3][[math]::floor($t/9) % 8]])
        $str += "$e[0m$e[1;1H"
    }
    Write-Output $str
    $sleep = 40 - $time.TotalMilliseconds
    if ($sleep -lt 0) {$sleep = 0}
    Sleep -Milliseconds $sleep
}