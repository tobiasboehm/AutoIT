ShellExecute("notepad.exe")

Fenster("Unbenannt")
;~ WinWait("Unbenannt","")
;~ If Not WinActive("Unbenannt","") _
;~ 	Then WinActivate("Unbenannt","")
;~ WinWaitActive("Unbenannt","")

MsgBox(0,"Offen","OK")

WinClose("Unbenannt")

Func Fenster (Const $titel)
   WinWait($titel,"")
If Not WinActive($titel,"") _
	Then WinActivate($titel,"")
WinWaitActive($titel,"")
EndFunc