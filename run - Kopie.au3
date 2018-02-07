#include <ButtonConstants.au3>
#include <DateTimeConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Form1", 806, 708, 192, 124)
$Label1 = GUICtrlCreateLabel("Label1", 8, 8, 388, 145)
$MonthCal1 = GUICtrlCreateMonthCal("2018/01/09", 8, 168, 180, 164)
$MonthCal2 = GUICtrlCreateMonthCal("2018/01/09", 216, 168, 180, 164)
$Button1 = GUICtrlCreateButton("Button1", 16, 344, 161, 49)
$Button2 = GUICtrlCreateButton("Button2", 224, 344, 161, 49)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

	EndSwitch
WEnd
