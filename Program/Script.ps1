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
function ��� {[CmdletBinding()] param (
    [ArgumentCompleter({MovArgsComplete @args})]
    $Folder)
    & $program\Mov.ps1 $Folder
}
echo "Mov [�������] - ������� � �������"

function Log {& $program\Log.ps1 @args}
function ��� {Log @args}
echo "Log - ������"
echo "Log vet - ������ ��� ���-�����"

function Open {& $program\Open.ps1 @args}
function ���� {Open @args}
echo "Open [�����] - ������� ���������"
echo "Open - ������� ��������� ���������"
echo "Open bill - ������� ����-�������"
echo "Open cmr - ������� CMR"
echo "Open set - ������� �������"

function Skip {& $program\Skip.ps1 @args}
function ���� {Skip @args}
echo "Skip [�����] - ���������� ���������"

function Unskip {& $program\Unskip.ps1 @args}
function ������ {Unskip @args}
echo "Unskip [�����] - �������� ������� ���������"

function Delete {& $program\Delete.ps1 @args}
function ������ {Delete @args}
echo "Delete [�����] - ������� ���������"

function Clip {& $program\Clip.ps1 @args}
function ���� {Clip @args}
echo "Clip set - ���������� ����� �������"
echo "Clip [`"+�����`" `"=������`" ...] - ���������� ����� �� ������� � ������"
echo "Clip last - ��������� ������������� clip"

function RenameSets {& $program\RenameSets.ps1 @args}
function ���������� {RenameSets @args}
echo "RenameSets - ������������� �������"

function Accept {& $program\Accept.ps1 @args}
function ������ {Accept @args}
echo "Accept [���� ������] - ������� ������"

#function Fix {& $program\Fix.ps1 @args}
#function ��� {Fix @args}
#echo "Fix - ��������� CMR"

function Art {& $program\Art.ps1 @args}
function ��� {Art @args}
echo "Art - ����"

echo ""