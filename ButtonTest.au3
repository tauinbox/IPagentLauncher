#Include <GuiToolBar.au3>

While 1
$BAutoAnswer = BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2029), $TBSTATE_CHECKED) OR BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2029), $TBSTATE_PRESSED)
$BAutoAnswer2 = BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2029), $TBSTATE_PRESSED)
MsgBox(0, "TEST", $BAutoAnswer & " - " & $BAutoAnswer2, 1)
Sleep(2000)
WEnd