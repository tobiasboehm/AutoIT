#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>

ShellExecute("mspaint.exe")
WinWait("Unbenannt - Paint")
;MouseClick($MOUSE_CLICK_LEFT, 481, 70, 2, 0)
HotKeySet("{ESC}", "Terminate")
Func Terminate()
    Exit
EndFunc

MouseDown($MOUSE_CLICK_LEFT)
for $i=800 to 300 step -5
   for $j=350 to 800 step 5
;MouseMove ( 300, 200 , 10)
MouseMove ( $j, $i , 0)
MouseMove ( $i, $j , 0)
next
Next
MouseUp($MOUSE_CLICK_LEFT)



;for $i=300 to 1000 step 50
;MouseClickDrag($MOUSE_CLICK_LEFT, 200, 200, $i, 800, 0)
;next






;MouseClickDrag($MOUSE_CLICK_LEFT, 100, 100, 200, 100, 0)
;MouseClickDrag($MOUSE_CLICK_LEFT, 200, 200, 200, 100, 0)
;MouseClickDrag($MOUSE_CLICK_LEFT, 100, 100, 100, 100, 0)
;MouseClickDrag($MOUSE_CLICK_LEFT, 100, 100, 100, 100, 0)
;MouseClickDrag($MOUSE_CLICK_LEFT, 100, 100, 100, 100, 0)
;MouseClickDrag($MOUSE_CLICK_LEFT, 100, 100, 100, 100, 0)
;MouseUp($MOUSE_CLICK_LEFT,  1)
;MouseDown($MOUSE_CLICK_LEFT,  1)
;MouseUp($MOUSE_CLICK_LEFT, 550, 650, 1)


