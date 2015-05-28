#include <Date.au3>


$hTimer = TimerInit() ; Запускает таймер и хранит дескриптор в переменной
Sleep(3000)
$iDiff = TimerDiff($hTimer) ; Возвращает разницу во времени, от предыдущего запуска TimerInit, дескриптор которого указан в переменной

$iTicks = Int($iDiff / 1000)
$iTicks = Mod($iTicks, 86400)
$iHr = Int($iTicks / 3600)
$iTicks = Mod($iTicks, 3600)
$iMn = Int($iTicks / 60)
$iSc = Mod($iTicks, 60)
MsgBox(0, "Test", "Duration: " & StringFormat("%02d", $iHr) & ":" & StringFormat("%02d", $iMn) & ":" & StringFormat("%02d", $iSc) & "." & StringLeft(Mod($iDiff, 1000), 1) & @CR)