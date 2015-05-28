
$Text = WinGetText("[CLASS:IpAgent]")
$Fpos = StringInStr($Text, @LF) + 1
MsgBox(0, "Text", StringMid($Text, $Fpos, 8))
