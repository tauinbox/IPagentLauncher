#Include <GuiToolBar.au3>
;$hWnd = ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]") ; <<<<< If it does not work check the INSTANCE
;$sText = _GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2029)
;MsgBox(64, 'SOLARIS', $sText & @CRLF)
;MsgBox(64, 'SOLARIS', "The button has the $TBSTYLE_CHECK style and is being clicked :" & $TBSTATE_CHECKED & @CRLF & _
;"The button is being clicked :" & $TBSTATE_PRESSED & @CRLF & _
;"The button accepts user input :" & $TBSTATE_ENABLED & @CRLF & _
;"The button is not visible and cannot receive user input :" & $TBSTATE_HIDDEN & @CRLF & _
;"The button is grayed :" & $TBSTATE_INDETERMINATE & @CRLF & _
;"The button is followed by a line break :" & $TBSTATE_WRAP & @CRLF & _
;"The button's text is cut off and an ellipsis is displayed :" & $TBSTATE_ELLIPSES & @CRLF & _
;"The button is marked :" & $TBSTATE_MARKED & @CRLF)

While 1
   $sText = _GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2029)
   If BitAND($sText, $TBSTATE_CHECKED) Then
	  MsgBox(64, 'SOLARIS', "Press detected!!!")
	  Exit
   EndIf
;MsgBox(64, 'SOLARIS', "Состояние: " & $sText & ", Ищем: " & $TBSTATE_PRESSED & ", Результат: " & BitAND($sText, $TBSTATE_PRESSED))
Sleep(1000)
WEnd