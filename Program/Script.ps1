$Global:work = Get-Location
$Global:program = "$work\Program"
$Global:data = "$work\Data"
$Global:skip_numbers = New-Object System.Collections.ArrayList

function MovArgsComplete{
    param ( $commaName,
            $parameterName,
            $wordToComplete,
            $commandAst,
            $fakeBoundParameters )

    (Get-ChildItem $work -Directory) |? {$_ -notin "Data", "Input", "Program"} |
        Where-Object {$_ -match "$wordToComplete"} |
        Sort-Object LastWriteTime -Descending |
        ForEach-Object {$_}
}

function Mov {[CmdletBinding()] param (
    [ArgumentCompleter({MovArgsComplete @args})]
    $Folder)
    & $program\Mov.ps1 $Folder
}
function №щм {[CmdletBinding()] param (
    [ArgumentCompleter({MovArgsComplete @args})]
    $Folder)
    & $program\Mov.ps1 $Folder
}
echo "Mov [каталог] - перейти в каталог"

function Log {& $program\Log.ps1 @args}
function ƒщп {Log @args}
echo "Log - журнал"
echo "Log vet - журнал дл€ вет-врача"

function Open {& $program\Open.ps1 @args}
function ўзут {Open @args}
echo "Open [номер] - открыть накладную"
echo "Open - открыть следующую накладную"
echo "Open bill - открыть счет-фактуры"
echo "Open cmr - открыть CMR"
echo "Open set - открыть сводные"

function Skip {& $program\Skip.ps1 @args}
function џлшз {Skip @args}
echo "Skip [номер] - пропустить накладную"

function Unskip {& $program\Unskip.ps1 @args}
function √тылшз {Unskip @args}
echo "Unskip [номер] - отменить пропуск накладной"

function Delete {& $program\Delete.ps1 @args}
function ¬удуеу {Delete @args}
echo "Delete [номер] - удалить накладную"

function Clip {& $program\Clip.ps1 @args}
function —дшз {Clip @args}
echo "Clip set - копировать имена сводных"
echo "Clip [`"+число`" `"=строка`" ...] - копировать числа по пор€дку и строки"
echo "Clip last - последнее использование clip"

function RenameSets {& $program\RenameSets.ps1 @args}
function  утфьуџуеы {RenameSets @args}
echo "RenameSets - переименовать сводные"

function Accept {& $program\Accept.ps1 @args}
function ‘ссузе {Accept @args}
echo "Accept [файл за€вки] - прин€ть за€вку"

#function Fix {& $program\Fix.ps1 @args}
#function јшч {Fix @args}
#echo "Fix - исправить CMR"

function Art {& $program\Art.ps1 @args}
function ‘ке {Art @args}
echo "Art - Єлка"

echo ""