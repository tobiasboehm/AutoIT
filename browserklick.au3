#include <IE.au3>
#include <MsgBoxConstants.au3>

If WinExists("IEFrame") Then
        MsgBox($MB_SYSTEMMODAL, "", "Window exists")
    Else
        MsgBox($MB_SYSTEMMODAL, "", "Window does not exist")
		Local $oIE = _IECreate("www.google.com")
    EndIf
WinWait("Google - Internet Explorer","")
If Not WinActive("Google - Internet Explorer","") _
	Then WinActivate("Google - Internet Explorer","")
WinWaitActive("Google - Internet Explorer","")
;Sleep(2000)
send("^t")