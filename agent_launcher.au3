#NoTrayIcon
#include <file.au3>
#include <Array.au3>
#Include <GuiToolBar.au3>
#include <Process.au3>
#include <Date.au3>
#include <WindowsConstants.au3>
#Include <StaticConstants.au3>
#include <GuiToolTip.au3> 
#Include <WinAPIEx.au3>

;OnAutoItExitRegister('_Quit')
;_WinAPI_EmptyWorkingSet()

Global Const $DEFAULT_YELLOW = 5
Global Const $DEFAULT_RED = 7

Global $AvayaID, $Exten=0, $TEXT, $hWnd, $FileData, $Mode=0, $ModeNow=0, $FirstStart=0, $Line1, $Line2, $Line3, $ActiveCall=0, $hTimer, $iDiff, $iHr, $iMn, $iSc, _
$BPause, $BHoldAfterCall, $BAutoAnswer, $BAutoAnswerCHK, $BAutoAnswerPRS, $BManualAnswer, $BOnHold, $Fpos, $HoldFlag=0, $Timer, $h_tooltip, $PauseTimer, $SplashFlag=0, $i, $CallerID[1][3], $HoldTimer, _
$CCounter=0, $CallInfo, $ModeTimer

$tYellow = $DEFAULT_YELLOW
$tRed = $DEFAULT_RED

$CallerID[0][0] = ""
$CallerID[0][1] = ""
$CallerID[0][2] = ""

;  Mode:
;1 - Вспомогательный рабочий режим
;2 - Режим работы после вызова
;3 - Режим автоматического приема
;4 - Режим ручного ответа
;5 - Активный вызов, ожидается переход в "Режим работы после вызова" по завершению звонка
;6 - Активный вызов, ожидается переход во "Вспомогательный рабочий режим" по завершению звонка

$aProcessList = ProcessList("Solaris.exe")
If $aProcessList[0][0] > 1 Then
   _WriteLog("Double launch detected. Exiting...")
   MsgBox(64, 'SOLARIS', 'Копия программы SOLARIS уже запущена. Выйдите из программы или дождитесь завершения её работы.', 4)
   Exit
EndIf

If ProcessExists("IpAgent.exe") Then
   _WriteLog("IpAgent.exe already runned. Exiting...")
   MsgBox(64, 'SOLARIS', 'Обнаружен активный Avaya IP Agent. Запуск Solaris невозможен. Приложение будет завершено.', 4)
   $hWnd = WinWait("[CLASS:IpAgent]", "", 5)
   _CloseIpAgent()
   Exit
EndIf

If FileGetSize(@TempDir & "\solaris.log") > 10485760 Then
   FileDelete(@TempDir & "\solaris.log")
EndIf

If FileExists (@TempDir & "\busy.flag") Then
   _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",12," & $Exten & ",'" & @ComputerName & "',DEFAULT,DEFAULT," & "'termination detected'", "event.id" )
   _WriteLog('[' & @ComputerName & '\' & @UserName & "] Abnormal program termination on previous run detected")
   MsgBox(64, 'SOLARIS', "Предыдущее завершение программы было некорректным!", 3)
Else
   FileClose (FileOpen (@TempDir & "\busy.flag", 2))
EndIf

If FileExists(@TempDir & "\agent.id") Then
   FileDelete(@TempDir & "\agent.id")
EndIf

_WriteLog('[' & @ComputerName & '\' & @UserName & "] -=-=-=-=-=-=-=-Running  Solaris-=-=-=-=-=-=-=-")
_SQLExec("exec SOLARIS.dbo.GET_AVAYA_ID " & "'" & @UserName & "',DEFAULT,'" & @ComputerName & "',DEFAULT,DEFAULT," & "'run and get id'", "agent.id" )

$file = FileOpen(@TempDir & "\agent.id", 0)
$AvayaID = StringStripWS (FileReadLine($file, 3), 1)

;MsgBox(64, 'SOLARIS', $AvayaID)

If StringLen($AvayaID) <> 4 and StringLeft($AvayaID, 1) <> '6' Then
   _WriteLog("Unable to get Agent ID")
   MsgBox(64, 'SOLARIS', "Не удалось получить идентификатор агента. Возможно, все идентификаторы заняты или для Вас нет назначенных кампаний.")
   _Exit()
EndIf

_WriteLog("Issued Agent ID: " & $AvayaID)

If Not _CalcExten() Then
   _WriteLog("Unable to determine extension")
   MsgBox(64, 'SOLARIS', 'Не удалось определить Extension. Проверьте имя компьютера.', 4)
   _Exit()
Else
   _WriteLog("Calculated Extension: " & $Exten)
   If IniRead(@AppDataDir& "\Avaya\Avaya IP Agent\Log Files\DEFINITY.ini", "DoLAN", "ACDSwitchHookMode", "0") = "0" Then
	  $FirstStart = 1
	  _WriteLog("First launch detected")
	  IniWrite(@AppDataDir& "\Avaya\Avaya IP Agent\Log Files\DEFINITY.ini", "DoLAN", "ACDSwitchHookMode", "1")
	  RegDelete('HKEY_CURRENT_USER\Software\Avaya')
   EndIf

   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "Extension", "REG_SZ", $Exten)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "ASN1Logging", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "AudioParametersTuned", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "AudioPortRangeHigh", "REG_DWORD", 0xffff)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "AudioPortRangeLow", "REG_DWORD", 2048)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "AutoLogin", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "BlockEmergencyCall", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "Debug", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "DialupID", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "DialupLocation", "REG_SZ", "Мое размещение")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "DontShowLoginStatus", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "E911DisclaimerShown", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "MicrophoneVolume", "REG_DWORD", 0xd400)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "NDT_enabled", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "NetworkBandwidth", "REG_DWORD", 4)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "NetworkRegion", "REG_DWORD", 0xffffffff)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "OtherTAPIAppDirOpt", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "PortHigh", "REG_DWORD", 0xffff)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "PortLow", "REG_DWORD", 1025)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "PrimaryServer", "REG_SZ", "192.168.161.6")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "RegistrationParametersTuned", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "RingerMuted", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "RingerVolume", "REG_DWORD", 33792)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "SavePassword", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "ShowTip", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "SpeakerMasterVolume", "REG_DWORD", 65535)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "SpeakerVolume", "REG_DWORD", 0xd400)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "UseOwnExtForEmergencyCall", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "WindowsPrevLocationID", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Options", "IMFeatureEnabled", "REG_DWORD", 0)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Parameters", "nNormalPlayVolume", "REG_DWORD", 0xd400)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Parameters", "fReceiveGain", "REG_SZ", "5,0000000000")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Parameters", "fTransmitGain", "REG_SZ", "1,5000000000")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\iClarity\Parameters", "bComfortNoise", "REG_DWORD", 1)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\PhoneFeatures\Логическая линия", "-1879048175", "REG_SZ", "Логическая линия " & $Exten)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\PhoneFeatures\Логическая линия", "-1879048177", "REG_SZ", "Логическая линия " & $Exten)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\PhoneFeatures\Логическая линия", "-1879048179", "REG_SZ", "Логическая линия " & $Exten)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "Enabled", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "IncomingCall", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "CallState", "REG_DWORD", 2)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "PopOnVdn", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "Vdn", "REG_SZ", "")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "DataExchange", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "Service", "REG_SZ", "")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "Topic", "REG_SZ", "")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "Function", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "Command", "REG_SZ", "")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "Item", "REG_SZ", "")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "ItemData", "REG_SZ", "")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "WindowsExplorer", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "Address", "REG_SZ", '"C:\Program Files\Internet Explorer\iexplore.exe" http://upiter:8089/Underway.aspx/StartAvaya/%v/%m')
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "NameFormatString", "REG_SZ", "%n")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "NumberFormatString", "REG_SZ", "%m")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "VdnFormatString", "REG_SZ", "%v")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "UuiFormatString", "REG_SZ", "%u")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "PromptedFormatString", "REG_SZ", "%p")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "StartTimeFormat", "REG_SZ", "%s")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "EndTimeFormat", "REG_SZ", "%e")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter", "DateFormat", "REG_SZ", "%d")
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-0", "Format", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-0", "Length", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-0", "StartAt", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-0", "Location", "REG_DWORD", 1)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-1", "Format", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-1", "Length", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-1", "StartAt", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-1", "Location", "REG_DWORD", 1)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-2", "Format", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-2", "Length", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-2", "StartAt", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-2", "Location", "REG_DWORD", 1)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-3", "Format", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-3", "Length", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-3", "StartAt", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-3", "Location", "REG_DWORD", 1)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-4", "Format", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-4", "Length", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-4", "StartAt", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\ScreenPops\upiter\Format-4", "Location", "REG_DWORD", 1)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Settings", "PopupButtonLimit", "REG_DWORD", 0x0e)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Settings", "LogMonths", "REG_DWORD", 3)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Settings", "RecentCalls", "REG_DWORD", 6)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Settings", "FontSize", "REG_DWORD", 0x0e)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Settings", "MinPhoneNumber", "REG_DWORD", 0x0a)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Settings", "MaxPhoneNumber", "REG_DWORD", 0x0a)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Settings", "MaxCityCode", "REG_DWORD", 3)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Settings", "NewContactWindow", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Settings", "ContactDirectoryDocked", "REG_DWORD", 0)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Settings", "Options", "REG_SZ", "0,1,-1,-1,-1,-1,444,244,996,627")
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\BarState\7.0.0\RoadWarrior-Bar4", "Visible", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\BarState\7.0.0\RoadWarrior-Bar10", "BarID", "REG_DWORD", 0x0132)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\BarState\7.0.0\RoadWarrior-Bar10", "Visible", "REG_DWORD", 1)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Settings", "MainFrame", "REG_SZ", "0,1,-1,-1,-1,-1,77,77,817,277")
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "Password", "REG_SZ", "A70C653CB0")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "AgentGreetings", "REG_DWORD", 4)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "CallHistory", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "ConfigAdmin", "REG_DWORD", 4)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "ExportSettings", "REG_DWORD", 4)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "ImportSettings", "REG_DWORD", 4)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "PhoneDirectory", "REG_DWORD", 4)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "PhoneFeatures", "REG_DWORD", 4)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "PublicSearch", "REG_DWORD", 4)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "ScreenPops", "REG_DWORD", 4)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "SpeedDials", "REG_DWORD", 4)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "VuStats", "REG_DWORD", 4)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\FeatureAccess", "ProgramOptions", "REG_DWORD", 4)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "AutoAnswer", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "AssistedDialing", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "EasDefinity", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "AgentLogin", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "SavePassword", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "AutoAgentLogin", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "BasicTransfer", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "ReplaceDisplay", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "ReplaceTitle", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "AlwaysOnTop", "REG_DWORD", 0)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "DialpadLetters", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "ToolTips", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "ShortcutIcon", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "FlashWindow", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Config\Options", "PopOnCall", "REG_DWORD", 0)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\PassageWay\Definity\DoLAN", "ACDSwitchHookMode", "REG_DWORD", 1)
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\PassageWay\Definity\DoLAN", "DebugGAPI", "REG_DWORD", 0)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya Web Dialer\Config\Options", "BrowserParse", "REG_DWORD", 1)
   
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Agent\Paster", "PASTEFAC", "REG_SZ", "")
   RegWrite("HKEY_CURRENT_USER\Software\Avaya\Avaya IP Agent\Agent\Paster", "Debug", "REG_DWORD", 0)
   
   _WriteLog("Reg params successfully imported")
EndIf

ReadINI()

_WriteLog("Running IpAgent.exe")
Run("C:\Program Files\Avaya\Avaya IP Agent\IpAgent.exe /lang rus")
;SplashTextOn("SOLARIS", "Стартуем, ждите..." & @CRLF & @CRLF & "Просьба не пользоваться мышкой и клавиатурой пока не погаснет это сообщение", 550, 210, 80, 80, 4, "", 24)
SplashTextOn("SOLARIS", "Стартуем, ждите..." & @CRLF & @CRLF & "Просьба не пользоваться мышкой и клавиатурой пока не погаснет это сообщение", 733, 184, 77, 77, 4, "", 24)

BlockInput (1)

While $FirstStart 
   If WinExists("[CLASS:#32770]", "Потеряна база данных программы") Then
	  WinActivate("[CLASS:#32770]")
	  Send("{ENTER}")
	  Sleep(500)
	  Send("{ENTER}")
	  Sleep(500)
	  Send("{ENTER}")
   EndIf

   If WinExists("[CLASS:#32770]", "Убедитесь, что закрыты все приложения") Then
	  WinActivate("[CLASS:#32770]")
	  Send("{ENTER}")
	  ControlClick ("Мастер настройки аудио", "", "[CLASS:Button; INSTANCE:8]", "left", 1)
	  WinWait("[CLASS:#32770]", "Выберите конфигурацию для вызовов VoIP")
   EndIf

   If WinExists("[CLASS:#32770]", "Выберите конфигурацию для вызовов VoIP") Then
	  WinActivate("[CLASS:#32770]")
	  Send("{ENTER}")
	  ControlClick ("Мастер настройки аудио", "", "[CLASS:Button; INSTANCE:8]", "left", 1)
	  Sleep(500)
   EndIf
   
   If WinExists("[CLASS:#32770]", "Подсоедините и включите динамик") Then
	  WinActivate("[CLASS:#32770]")
	  Send("{ENTER}")
	  ControlClick ("Мастер настройки аудио", "", "[CLASS:Button; INSTANCE:8]", "left", 1)
	  Sleep(500)
   EndIf
   
   If WinExists("[CLASS:#32770]", "Подсоедините и включите микрофон") Then
	  WinActivate("[CLASS:#32770]")
	  Send("{ENTER}")
	  ControlClick ("Мастер настройки аудио", "", "[CLASS:Button; INSTANCE:8]", "left", 1)
	  Sleep(500)
   EndIf
   
   If WinExists("[CLASS:#32770]", "чтобы определить уровень") Then
	  WinActivate("[CLASS:#32770]")
	  Sleep(500)
	  Send("{ENTER}")
	  Sleep(500)
	  Send("{ENTER}")
   EndIf

   If WinExists("[CLASS:#32770]", "Невозможно войти на сервер, пока не настроены параметры аудио.") Then
	  WinActivate("[CLASS:#32770]")
	  Send("{ENTER}")
	  Sleep(500)
	  Send("{ENTER}")
   EndIf
   
   If WinExists("[CLASS:#32770]", "Your login attempt was unsuccessful") Then
	  WinActivate("[CLASS:#32770]")
	  Send("{ENTER}")
   EndIf
   
   If WinExists("[CLASS:#32770]", "Добавочный номер:") Then
	  WinActivate("[CLASS:#32770]")
	  Send("{TAB}")
	  Send("0000")
	  Send("{ENTER}")
	  WinClose ("[CLASS:Afx:400000:8:10011:0:0]")
	  $FirstStart = 0
	  BlockInput (0)
	  $hWnd = WinWait("[CLASS:IpAgent]", "", 3)
	  WinActivate($hWnd)
	  SplashOff()
	  _CloseIpAgent()
	  FileDelete (@TempDir & "\busy.flag")
	  _WriteLog("First initialisation is complete. Exiting...")
	  MsgBox(4096, 'SOLARIS', "Настройка параметров завершена. Требуется повторный запуск приложения")
	  Exit
   EndIf
   
  Sleep(100)
WEnd

$hWnd = WinWait("[CLASS:IpAgent]", "", 45)

If Not $hWnd Then
   _WriteLog("Unable to start IpAgent.exe")
   _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",13," & $Exten & ",'" & @ComputerName & "',DEFAULT,DEFAULT," & _
   "'Unable to start IpAgent.exe'", "event.id" )
   SplashOff()
   ProcessClose("IpAgent.exe")
   MsgBox(64, 'SOLARIS', 'Не удалось запустить Avaya IP Agent', 4)
   _Exit()
EndIf

Sleep(3000)

_LogonAgent()

$ModeTimer = TimerInit()

;$TEXT =  WinGetText($hWnd)
;MsgBox(64, 'SOLARIS', $TEXT) 

While 1
   If WinExists($hWnd) = 0 Then
	  _Exit()
   Else
	  $Line1 =  ControlGetText ($hWnd, "", "[CLASS:Static; INSTANCE:5]")
	  $Line2 =  ControlGetText ($hWnd, "", "[CLASS:Static; INSTANCE:10]")
	  $Line3 =  ControlGetText ($hWnd, "", "[CLASS:Static; INSTANCE:15]")
	  
	  $BOnHold = BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:10]"), 2013), $TBSTATE_CHECKED) OR BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:10]"), 2013), $TBSTATE_PRESSED)
	  
	  If $ActiveCall = 1 And (StringInStr($Line1, "=") = 0 And StringInStr($Line2, "=") = 0 And StringInStr($Line3, "=") = 0) Then
		 $ActiveCall = 0
		 TimerDuration($hTimer) ;Вычисляем время длительности звонка
		 $Timer = ""
		 ToolTip($Timer)
		 _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",8," & $Exten & ",'" & @ComputerName & "'," & $iDiff & ",'" & $CallInfo & "'," & "'call ended'", "event.id")
		 _WriteLog("CALL ENDED. (" & StringFormat("%02d", $iHr) & ":" & StringFormat("%02d", $iMn) & ":" & StringFormat("%02d", $iSc) & ")")
		 $tYellow = $DEFAULT_YELLOW
		 $tRed = $DEFAULT_RED
	  EndIf
	  
	  CheckMode()

	  If $ActiveCall = 0 And (StringInStr($Line1, "=") <> 0 Or StringInStr($Line2, "=") <> 0 Or StringInStr($Line3, "=") <> 0) And Not (Not StringCompare($Line1, "a=") Or Not StringCompare($Line1, "b=")) Then
		 $ActiveCall = 1
		 $hTimer = TimerInit()
		 $CCounter = $CCounter + 1
		 SplashOff(); На всякий случий гасим всплывающее окно с сообщением
		 Sleep(500)
		 CheckMode();Ещё разок проверяем режим после паузы
		 $Line1 =  ControlGetText ($hWnd, "", "[CLASS:Static; INSTANCE:5]") ;Ещё разок считываем текст первой линии
		 $CallInfo = StringStripWS(StringReplace($Line1, "a=", ""), 5)
		 _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",7," & $Exten & ",'" & @ComputerName & "',DEFAULT,'" & $CallInfo & "'," & "'call detected'", "event.id" )
		 _WriteLog("CALL DETECTED [" & $CCounter & "] (" & $CallInfo & ")")
		 For $j=0 To $i ;Цикл поиска значений таймеров для разных кампаний
			$Len = StringLen($CallerID[$j][0])
			If StringCompare(StringRight($CallInfo, $Len), $CallerID[$j][0]) = 0 And Number($Len) > 0 Then
			   $tYellow = $CallerID[$j][1]
			   $tRed = $CallerID[$j][2]
			   _WriteLog("Apply timer settings to " & $CallerID[$j][0] & ". YellowZone: " & $CallerID[$j][1] & ", RedZone: " & $CallerID[$j][2])
			   ExitLoop
			EndIf
		 Next
	  ElseIf Not StringCompare($Line1, "a=") Or Not StringCompare($Line1, "b=") Then
;		 MsgBox(0, "SOLARIS", "Ay-ya-yay!")
		 _WriteLog("Line busy for outcoming. Force hang up!")
		 _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",13," & $Exten & ",'" & @ComputerName & "',DEFAULT,DEFAULT," & _
			"'Line busy for outcoming. Force hang up!'", "event.id" )
		 ControlClick ("[CLASS:IpAgent]", "", "[CLASS:Button; INSTANCE:1]", "left", 1)
	  EndIf
	  
	  If $BOnHold And Not $HoldFlag Then
		 $HoldFlag = 1
		 $HoldTimer = TimerInit()
		 _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",9," & $Exten & ",'" & @ComputerName & "',DEFAULT,'" & $CallInfo & "'," & "'putting a call on hold'", "event.id" )
		 _WriteLog("Putting a call on hold")
	  EndIf
	  
	  If $HoldFlag And Not $BOnHold And $ActiveCall Then
		 $HoldFlag = 0
		 TimerDuration($HoldTimer) ;Вычисляем время нахождения в режиме удержания вызова
		 _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",10," & $Exten & ",'" & @ComputerName & "'," & $iDiff & ",'" & $CallInfo & "'," & "'taking a call off hold'", "event.id" )
		 _WriteLog("Taking a call off hold (" & StringFormat("%02d", $iHr) & ":" & StringFormat("%02d", $iMn) & ":" & StringFormat("%02d", $iSc) & ")")
	  EndIf
	  
	  If $HoldFlag And Not $BOnHold And Not $ActiveCall Then
		 $HoldFlag = 0
		 TimerDuration($HoldTimer)  ;Вычисляем время нахождения в режиме удержания вызова
		 _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",11," & $Exten & ",'" & @ComputerName & "'," & $iDiff & ",'" & $CallInfo & "'," & "'hung up the call on hold'", "event.id" )
		 _WriteLog("Hung up the call on hold (" & StringFormat("%02d", $iHr) & ":" & StringFormat("%02d", $iMn) & ":" & StringFormat("%02d", $iSc) & ")")
	  EndIf

	  If $ActiveCall = 1 Then
		 $TEXT = WinGetText("[CLASS:IpAgent]")
		 $Fpos = StringInStr($Text, @LF) + 4
		 $Timer = StringMid($Text, $Fpos, 5)
		 ToolTip($Timer)
		 $h_tooltip = WinGetHandle ($Timer)
		 $CallMin = StringFormat("%u", StringMid ($Timer, 1, 2))
		 Select
		 Case Number($CallMin) >= Number($tYellow) And Number($CallMin) < Number($tRed) And Not $HoldFlag
			DllCall("user32.dll","int","SendMessage","hwnd",$h_tooltip,"int",$TTM_SETTIPBKCOLOR,"int",0x00DEFF,"int",0)
		 Case Number($CallMin) >= Number($tRed) And Not $HoldFlag
			DllCall("user32.dll","int","SendMessage","hwnd",$h_tooltip,"int",$TTM_SETTIPBKCOLOR,"int",0x7A7AFF,"int",0)
		 Case $HoldFlag
			DllCall("user32.dll","int","SendMessage","hwnd",$h_tooltip,"int",$TTM_SETTIPBKCOLOR,"int",0xDCDCC4,"int",0)
		 Case Else
			DllCall("user32.dll", "int", "SendMessage", "hwnd", $h_tooltip, "int", $TTM_SETTIPBKCOLOR, "int", 0xB8F2B1, "int", 0)
		 EndSelect
	  EndIf
	  
	  If $Mode = 1 And $ActiveCall = 0 And TimerDiff($PauseTimer) > 1800000 And $SplashFlag = 0 Then
		 $SplashFlag = 1
		 _WriteLog("Idle time expired")
		 SplashTextOn("SOLARIS", "Превышено максимальное время паузы", 590, 70, -1, -1, 4, "", 24)
	  EndIf

	  If _GUICtrlToolbar_GetButtonText(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2138) = "Вход в систему" Then
		_SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",5," & $Exten & ",'" & @ComputerName & "',DEFAULT,DEFAULT," & "'agent is disconnected'", "event.id" )
		_WriteLog("Agent is disconnected. Exiting...")
		_CloseIpAgent()
		_Exit()
	  EndIf
;	 MsgBox(64, 'SOLARIS', "Режим: " & $Mode & @CR & "Линия 1: " & $Line1 & @CR & "Линия 2: " & $Line2 & @CR & "Линия 3: " & $Line3, 1)
   EndIf

   If WinExists("[CLASS:IEFrame]") Then
	  $TEXT=ControlGetText("[CLASS:IEFrame]","","[CLASS:Edit;INSTANCE:1]")
	  If StringCompare($TEXT, "http://upiter:8089/Underway.aspx/EndScript", 0) = 0 Then
		 Sleep(1000)
		 WinClose("[CLASS:IEFrame]")
		 _WriteLog("Web-script is finished")
;		 MsgBox(64, "SOLARIS", "Сценарий завершён", 2)
	  ElseIf StringCompare($TEXT, "http://upiter:8089/mycustompage.htm?aspxerrorpath=/Underway.aspx/Error", 0) = 0 Then
		 Sleep(1000)
		 WinClose("[CLASS:IEFrame]")
		 _WriteLog("Web-script loading error")
		 MsgBox(64, "SOLARIS", "Ошибка загрузки страницы сценария", 2)
	  EndIf
   EndIf
   
Sleep(20)   
WEnd


;-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Func _Exit()
   _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",4," & $Exten & ",'" & @ComputerName & "',DEFAULT,DEFAULT," & "'exiting'", "event.id" )
   _WriteLog('[' & @ComputerName & '\' & @UserName & "] -=-=-=-=-=-=-=-=-=-=-EXIT-=-=-=-=-=-=-=-=-=-=-" & @CRLF)
   FileDelete (@TempDir & "\busy.flag")
;   MsgBox(64, "SOLARIS", "Завершение работы", 3)
   Exit
EndFunc   ;==>_Exit

Func _LogonAgent()
   BlockInput (1)
   WinActivate($hWnd)
   Sleep(500)
   Send("{CTRLDOWN}")
   Send("{INS}")
   Send("{CTRLUP}")
   Send("+{TAB}")
   Send($AvayaID)
   Send("{TAB}")
   Send($AvayaID)
   Send("{ENTER}")
   Sleep(7000)
   $TEXT=WinGetTitle($hWnd)
   If StringInStr ($TEXT, $AvayaID) = 0 Then
	  _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",3," & $Exten & ",'" & @ComputerName & "',DEFAULT,DEFAULT," & "'login failed'", "event.id" )
	  _WriteLog("Unable to login Agent " & $AvayaID)
	  SplashOff()
	  MsgBox(64, "SOLARIS", "Не удалось войти агентом " & $AvayaID & " в систему. Возможно, данный идентификатор уже занят.", 4)
	  _CloseIpAgent()
	  _Exit()
   Else
	  _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",2," & $Exten & ",'" & @ComputerName & "',DEFAULT,DEFAULT," & "'successfully logged on'", "event.id" )
	  _WriteLog("Agent " & $AvayaID & " successfully logged on")
	  SplashOff()
   EndIf
   BlockInput (0)
;   MsgBox (16,"Т",ControlGetText ($hWnd, "", "[CLASS:Static; INSTANCE:5]"))
EndFunc

Func _CloseIpAgent()
   BlockInput (1)
   WinActivate($hWnd)
   Sleep(500)
   WinClose($hWnd)
   Sleep(500)
   If WinExists("[CLASS:#32770]", "Хотите выйти из программы") Then
	  _WriteLog("Agent is on-line. Apply Yes-key to hang up")
	  ControlClick("[CLASS:#32770]", "", "[CLASS:Button; INSTANCE:1]")
   EndIf
   Sleep(500)
   If WinExists("[CLASS:#32770]", "Хотите закрыть программу") Then
	  _WriteLog("Agent is logged on. Apply Yes-key to exit")
	  ControlClick("[CLASS:#32770]", "", "[CLASS:Button; INSTANCE:1]")
   EndIf
;	  MsgBox(64, "SOLARIS", "!!!")
   WinWaitClose($hWnd, 6)
   If WinExists($hWnd) Then
	  _WriteLog("Killing process IpAgent.exe")
	  ProcessClose("IpAgent.exe")
   EndIf
   BlockInput (0)
EndFunc

Func _CalcExten()
   $CName = @ComputerName
   $Octet = StringMid($CName, 3, 3)

   Select 
   Case $Octet >= 100 AND $Octet <= 103 
	  $Numb = (StringMid($CName, 5, 1)*300 + StringMid($CName, 6, 3))
	  $Numb = StringFormat("%03i", $Numb)
	  $Exten = "5" & $Numb
	  Return  $Exten
   ; MsgBox(64, "SOLARIS", $Octet)
   Case $Octet = "020"
	  $Exten = StringMid($CName, 6, 3) + 1200
	  Return  $Exten
   Case Else
	  Return 0
   EndSelect
EndFunc

Func _SQLExec($CmdString, $OutFName)
   $file = FileOpen("C:\Program Files\Avaya\Scripts\template.sql", 0)
   $FileData = FileRead($file) & $CmdString
   FileClose($file)

   $file = FileOpen(@TempDir & "\template.sql", 2)
   FileWrite($file, $FileData)
   FileClose($file)

;MsgBox(64, 'SOLARIS', $FileData, 4)

   RunWait("C:\Program Files\Microsoft SQL Server\100\Tools\Binn\sqlcmd.exe -S mars -U task_executor -P Up1864 -i " & @TempDir & "\template.sql -o " & @TempDir & "\" & $OutFName, "C:\Program Files\Microsoft SQL Server\100\Tools\Binn", @SW_HIDE)
EndFunc

Func _WriteLog($Msg)
   $file = FileOpen (@TempDir & "\solaris.log", 1)
   FileWriteLine($file, @HOUR & ':' & @MIN & ':' & @SEC & ' ' & @MDAY & '/' & @MON & '/' & @YEAR & " " & $Msg & @CRLF)
   FileClose($file)
EndFunc

Func ReadINI()
   $file = FileOpen("\\pdc.fortline.org\NETLOGON\solaris.ini", 0)
   $i=0
   While 1
	  ReDim $CallerID[$i+1][3]
	  $line = FileReadLine($file)
	  If @error = -1 Then ExitLoop
	  If @error = 1 Then ExitLoop
	  $Params = StringSplit($line, "|")
	  $CallerID[$i][0] = $Params[1]
	  $CallerID[$i][1] = $Params[2]
	  $CallerID[$i][2] = $Params[3]
;	  MsgBox(4096, "Прочитанная строка:", $CallerID[$i][0] & ", " & $CallerID[$i][1] & ", " & $CallerID[$i][2])
	  $i=$i+1
   WEnd
   _WriteLog("Read " & $i & " lines from ini-file")
   FileClose($file)
EndFunc

Func CheckMode()
	  $BPauseCHK = BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2031), $TBSTATE_CHECKED)
	  $BPausePRS = BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2031), $TBSTATE_PRESSED)
	  $BHoldAfterCallCHK = BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2032), $TBSTATE_CHECKED)
	  $BHoldAfterCallPRS = BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2032), $TBSTATE_PRESSED)
	  $BAutoAnswerCHK = BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2029), $TBSTATE_CHECKED)
	  $BAutoAnswerPRS = BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2029), $TBSTATE_PRESSED)
	  $BManualAnswer = BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2030), $TBSTATE_CHECKED) OR BitAND(_GUICtrlToolbar_GetButtonState(ControlGetHandle("[CLASS:IpAgent]", "", "[CLASS:ToolbarWindow32; INSTANCE:3]"), 2030), $TBSTATE_PRESSED)
	  
	  $BPause = $BPauseCHK Or $BPausePRS
	  $BHoldAfterCall = $BHoldAfterCallCHK Or $BHoldAfterCallPRS
	  $BAutoAnswer = $BAutoAnswerCHK Or $BAutoAnswerPRS
	  
	  Select
	  Case $BPause And Not $ActiveCall
		 $ModeNow = 1
	  Case $BHoldAfterCall And Not $ActiveCall
		 $ModeNow = 2
	  Case $BAutoAnswer And Not $ActiveCall
		 $ModeNow = 3
	  Case $BManualAnswer And Not $ActiveCall
		 $ModeNow = 4
	  Case $BAutoAnswer And $BHoldAfterCall
		 $ModeNow = 5
	  Case $BAutoAnswer And $BPause
		 $ModeNow = 6
	  Case Else
		 If $BAutoAnswer And $ModeNow <> 3 And $ModeNow <> 5 And $ModeNow <> 6 Then
			$ModeNow = 3
;			$TEXT = ControlGetText ($hWnd, "", "[CLASS:Static; INSTANCE:5]")
			_SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",13," & $Exten & ",'" & @ComputerName & "',DEFAULT,DEFAULT," & _
			"'AutoAnswer:" & $BAutoAnswerCHK & "," & $BAutoAnswerPRS & "," & $BAutoAnswer & ". Mode fixed to 3'", "event.id" )
		 
			_WriteLog("[Missing mode detected] AutoAnswer:" & $BAutoAnswerCHK & "," & $BAutoAnswerPRS & "," & $BAutoAnswer & ". Fixed to 3")			
		 EndIf
;		 If Not $ActiveCall Then
;			$TEXT = ControlGetText ($hWnd, "", "[CLASS:Static; INSTANCE:5]")
;			_SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",13," & $Exten & ",'" & @ComputerName & "',DEFAULT,DEFAULT," & _
;			"'Pause:" & $BPauseCHK & "," & $BPausePRS & "," & $BPause & ", HoldAfterCall:" & $BHoldAfterCallCHK & _
;			"," & $BHoldAfterCallPRS & "," & $BHoldAfterCall & ", AutoAnswer:" & $BAutoAnswerCHK & "," & $BAutoAnswerPRS & "," & $BAutoAnswer & _
;			", Line:" & $TEXT & "'", "event.id" )
;		 
;			_WriteLog("[Can't detect mode] Pause:" & $BPauseCHK & "," & $BPausePRS & "," & $BPause & ", HoldAfterCall:" & $BHoldAfterCallCHK & _
;			"," & $BHoldAfterCallPRS & "," & $BHoldAfterCall & ", AutoAnswer:" & $BAutoAnswerCHK & "," & $BAutoAnswerPRS & "," & $BAutoAnswer & _
;			", Line:" & $TEXT)
;			Select
;			Case StringCompare($TEXT, "Режим автоматического приема") = 0
;			   $ModeNow = 3
;			Case StringCompare($TEXT, "Режим работы после вызова") = 0
;			   $ModeNow = 2
;			Case StringCompare($TEXT, "Вспомогательный рабочий режим") = 0
;			   $ModeNow = 1
;			EndSelect
;		 EndIf
	  EndSelect

	  If $Mode <> $ModeNow Then
		 TimerDuration($ModeTimer)
		 $ModeTimer = TimerInit()
		 _SQLExec("exec SOLARIS.dbo.EVENT_LOG_INSERT " & "'" & @UserName & "'" & ",6," & $Exten & ",'" & @ComputerName & "'," & $iDiff & ",DEFAULT," & "'mode changed'," & $ModeNow & "," & $Mode, "event.id" )
		 _WriteLog("Mode changed from " & $Mode & " to " & $ModeNow & " (" & StringFormat("%02d", $iHr) & ":" & StringFormat("%02d", $iMn) & ":" & StringFormat("%02d", $iSc) & ")")
		 $Mode = $ModeNow
		 If $Mode = 1 And $ActiveCall = 0 Then
			$PauseTimer = TimerInit()
		 EndIf
		 If $SplashFlag = 1 Then
			$SplashFlag = 0
			SplashOff()
		 EndIf
	  EndIf
EndFunc   

Func TimerDuration($TimerStr)
   $iDiff = TimerDiff($TimerStr)
   $iTicks = Int($iDiff / 1000)
   $iTicks = Mod($iTicks, 86400)
   $iHr = Int($iTicks / 3600)
   $iTicks = Mod($iTicks, 3600)
   $iMn = Int($iTicks / 60)
   $iSc = Mod($iTicks, 60)
EndFunc

;Func _Quit()
;   _CloseIpAgent()
;   _Exit()
;EndFunc
