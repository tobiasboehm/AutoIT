#include <GuiConstants.au3>
#include <EditConstants.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#region GLOBAL VARIABLES
Global $iW = 600, $iH = 400, $iT = 52, $iB = 52, $iLeftWidth = 150, $iGap = 10, $hMainGUI
#endregion GLOBAL VARIABLES

_MainGui()

Func _MainGui()
Local $hFooter, $nMsg, $aPos
Local $iLinks = 5
Local $sMainGuiTitle = "Auswertung"
Local $sHeader = "Bedienoberfläche des Auswertungsprogramms"
Local $sFooter = "Ordner 'Verbrauch' öffnen"
Local $aLink[$iLinks], $aPanel[$iLinks]
$aLink[0] = $iLinks - 1
$aPanel[0] = $iLinks - 1
$hMainGUI = GUICreate($sMainGuiTitle, $iW, $iH, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX, $WS_TABSTOP))
GUISetIcon("shell32.dll", -58, $hMainGUI)

GUICtrlCreateLabel($sHeader, 48, 8, $iW - 56, 32, $SS_CENTERIMAGE)
GUICtrlSetFont(-1, 14, 800, 0, "Arial", 5)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

GUICtrlCreateIcon("shell32.dll", -208, 8, 8, 32, 32)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlCreateLabel("", 0, $iT, $iW, 2, $SS_SUNKEN);separator
GUICtrlSetResizing(-1, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
GUICtrlCreateLabel("", $iLeftWidth, $iT + 2, 2, $iH - $iT - $iB - 2, $SS_SUNKEN);separator
GUICtrlSetResizing(-1, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH)
GUICtrlCreateLabel("", 0, $iH - $iB, $iW, 2, $SS_SUNKEN);separator
GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)

$hFooter = GUICtrlCreateLabel($sFooter, 10, $iH - 34, $iW - 20, 17, BitOR($SS_LEFT, $SS_CENTERIMAGE))
GUICtrlSetTip(-1, "Explorer", "Click to open...")
GUICtrlSetCursor(-1, 0)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT)

;add links to the left side
$aLink[1] = _AddNewLink("Auswertung Komplett", -167)
$aLink[2] = _AddNewLink("Letzte Liste Löschen", -132)

;and the corresponding GUI's
$aPanel[1] = _AddNewPanel("Auswertung Komplett")
$aPanel[2] = _AddNewPanel("Letzte Liste Löschen")
$aPanel[3] = _AddNewPanel("")


;add some controls to the panels
_AddControlsToPanel($aPanel[1])
GUICtrlCreateEdit("", 10, 37, $iW - $iLeftWidth + 2 - 20 - 5, 300 - $iT - $iB - 40, BitOR($ES_AUTOVSCROLL, $ES_NOHIDESEL, $ES_WANTRETURN, $WS_VSCROLL), $WS_EX_STATICEDGE)
Local $sTestTxt = ""
$sTestTxt &= "Dieses Programm erstellt die Jahres-(J_) und Monatsauswertungen(M_) aus SAP und fügt diese an die Gesamtlisten(Z_) hinzu." & @CRLF
$sTestTxt &= "Bitte beachten Sie:" & @CRLF
$sTestTxt &= @TAB & "-Alle SAP-Fenster werden geschlossen" & @CRLF
$sTestTxt &= @TAB & "-Alle Excel-Fenster werden geschlossen" & @CRLF
$sTestTxt &= @TAB & "-Ausweis muss im Kartenleser stecken" & @CRLF
$sTestTxt &= @TAB & "-Machen Sie KEINE Maus- oder Tastatureingaben" & @CRLF
$sTestTxt &= @TAB & "-Dauer ca. 5 Minuten" & @CRLF
$sTestTxt &= "Das Programm kann jederzeit mit F4 beendet werden!" & @CRLF
GUICtrlSetData(-1, $sTestTxt)
GUICtrlSetResizing(-1, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM)
Local $hButton1 = GUICtrlCreateButton("Start", 10, 200, 75, 25)
GUICtrlSetResizing(-1, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM)

_AddControlsToPanel($aPanel[2])
GUICtrlCreateEdit("", 10, 37, $iW - $iLeftWidth + 2 - 20 - 5, 300 - $iT - $iB - 40, BitOR($ES_AUTOVSCROLL, $ES_NOHIDESEL, $ES_WANTRETURN, $WS_VSCROLL), $WS_EX_STATICEDGE)
Local $sTestTxt = ""
$sTestTxt &= "Löscht das Tabellenblatt des letzten erstellten Monats in den Z_-Listen." & @CRLF
$sTestTxt &= "Ändert NICHT das JV-Blatt. Dieses kann nur durch erstellen der Auswertung geändert werden."  & @CRLF
$sTestTxt &= "Letzten erstellten Monat wirklich löschen?"
GUICtrlSetData(-1, $sTestTxt)
GUICtrlSetResizing(-1, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM)
Local $hButton2 = GUICtrlCreateButton("Start", 10, 200, 75, 25)
GUICtrlSetResizing(-1, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

;set default to Panel1
GUISwitch($aPanel[1])
;show the main GUI
GUISetState(@SW_SHOW, $hMainGUI)

While 1
Sleep(10)
$nMsg = GUIGetMsg(1)
Switch $nMsg[1]
Case $hMainGUI
Switch $nMsg[0]
Case $GUI_EVENT_CLOSE
Exit
Case $GUI_EVENT_MINIMIZE, $GUI_EVENT_MAXIMIZE, $GUI_EVENT_RESTORE
$aPos = WinGetPos($hMainGUI)
$iW = $aPos[2]
$iH = $aPos[3]
For $i = 0 To $aPanel[0]
WinMove($aPanel[$i], "", $iLeftWidth + 2, $iT, $iW - $iLeftWidth + 2, $iH - $iT - $iB - 20)
Next
Case $aLink[1], $aLink[2], $aLink[3], $aLink[4]
For $i = 1 To $aLink[0]
If $nMsg[0] = $aLink[$i] Then
GUISetState(@SW_SHOW, $aPanel[$i])
Else
GUISetState(@SW_HIDE, $aPanel[$i])
EndIf
Next
Case $hFooter
ShellExecute(@ScriptDir & "\Verbrauch")
EndSwitch
Case $aPanel[1]
Switch $nMsg[0]
Case $hButton1
MsgBox(32, "Test", "Auswertung Komplett starten")
EndSwitch
Case $aPanel[2]
Switch $nMsg[0]
Case $hButton2
MsgBox(32, "Test", "Letzte Liste Löschen")
EndSwitch
EndSwitch
WEnd
EndFunc ;==>_MainGui

Func _AddNewLink($sTxt, $iIcon = -44)
Local $hLink = GUICtrlCreateLabel($sTxt, 36, $iT + $iGap, $iLeftWidth - 46, 17)
GUICtrlSetCursor(-1, 0)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlCreateIcon("shell32.dll", $iIcon, 10, $iT + $iGap, 16, 16)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
$iGap += 22
Return $hLink
EndFunc ;==>_AddNewLink

Func _AddNewPanel($sTxt)
Local $gui = GUICreate("", $iW - $iLeftWidth + 2, $iH - $iT - $iB, $iLeftWidth + 2, $iT, $WS_CHILD + $WS_VISIBLE, -1, $hMainGUI)
GUICtrlCreateLabel($sTxt, 10, 10, $iW - $iLeftWidth - 20, 17, $SS_CENTERIMAGE)
GUICtrlSetFont(-1, 9, 800, 4, "Arial", 5)
GUICtrlSetResizing(-1, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
Return $gui
EndFunc ;==>_AddNewPanel

Func _AddControlsToPanel($hPanel)
GUISwitch($hPanel)
EndFunc ;==>_AddControlsToPanel