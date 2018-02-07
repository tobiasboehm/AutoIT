#include <MsgBoxConstants.au3>
#include <Misc.au3>

;~ Funktion zum Beenden mit F4
HotKeySet("{F4}", "ExitProg")
Func ExitProg()
    MsgBox($MB_ICONWARNING, "Beendet", "Programm beendet")
    Exit 0
EndFunc

;~ Konstanten, Arrays, Variablen
Const $Pfad_data	= "\\ww005.siemens.net\meddfsroot\1_MR\1_HQMR\Teams\Produktdaten\Tobi\Lieferzahlen\Auswertung_Tobi\Verbrauch"
Const $Angemeldet	= EnvGet("USERNAME")

;~ Autorisierte Benutzer
Dim $Benutzer[1]					;<--Bei neuem Benutzer das Array um 1 erhöhen
	$Benutzer[0] 	= "z002u7wh"	;Benutzer mit fortlaufender Nummer kopieren und Kennung einfügen


;~ Variantenauswahl in SAP
Dim $Variante_waehlen[12]			;<--Bei neuer Variante das Array um 1 erhöhen
    $Variante_waehlen[0]	=	"{RIGHT}{F2}"	;Variante mit fortlaufender Nummer einfügen
	$Variante_waehlen[1]	=	"{DOWN}{F2}"
	$Variante_waehlen[2]	=	"{DOWN 2}{F2}"
	$Variante_waehlen[3]	=	"{DOWN 3}{F2}"
	$Variante_waehlen[4]	=	"{DOWN 4}{F2}"
	$Variante_waehlen[5]	=	"{DOWN 5}{F2}"
	$Variante_waehlen[6]	=	"{DOWN 6}{F2}"
	$Variante_waehlen[7]	=	"{DOWN 7}{F2}"
	$Variante_waehlen[8]	=	"{DOWN 8}{F2}"
	$Variante_waehlen[9]	=	"{DOWN 9}{F2}"
	$Variante_waehlen[10]	=	"{DOWN 10}{F2}"
	$Variante_waehlen[11]	=	"{DOWN 11}{F2}"

;~ Dateinamen der gespeicherten Excel-Listen, Erweiterbar wie bei Variantenauswahl
Dim $Dateinamen[12]										;Variantennamen
    $Dateinamen[0]	=	"Lieferungen_2050_LC.xls"		;LIEFERUNG_LC
	$Dateinamen[1]	=	"Lieferungen_8er_LC.xls"		;LIEFERUNG_LC-8
	$Dateinamen[2]	=	"Lieferungen_END_LC.xls"		;LIEFERUNG_LC-E
	$Dateinamen[3]	=	"Lieferungen_2051_LC.xls"		;LIEFERUNG_LCJN
	$Dateinamen[4]	=	"Lieferungen_NW_2050_LC.xls"	;LIEF_NW2050
	$Dateinamen[5]	=	"Lieferungen_RW_2050_LC.xls"	;LIEF_RW2050
	$Dateinamen[6]	=	"Verbrauch_2050_gesamt.xls"		;VERBR2050_GES
	$Dateinamen[7]	=	"Verbrauch_2051_gesamt.xls"		;VERBR2051_GES
	$Dateinamen[8]	=	"Verbrauch_2050_BC.xls"			;VERBRAUCH_BC
	$Dateinamen[9]	=	"Verbrauch_2050_GC.xls"			;VERBRAUCH_GC
	$Dateinamen[10]	=	"Verbrauch_2050_LC.xls"			;VERBRAUCH_LC
	$Dateinamen[11]	=	"Verbrauch_2051_LC.xls"			;VERBRAUCH_LCJN

Dim $Var_Anzahl = Ubound($Variante_waehlen)		;Zählt die Anzahl der vorhandenen Varianten
Dim $i				;Zähler
Dim $j				;Zähler
Dim $abbrechen		;Abbrechbedingung
Dim $datum			;Platzhalter für beide Eingabevarianten
Dim $start_datum	;Startdatum für Monatseingabe
Dim $end_datum		;Enddatum für beide Eingabevarianten
Dim $zusatz_name

;AutoIt-Optionen
Opt("WinTitleMatchMode",4)
Opt("WinDetectHiddenText",1)

; Prüfen, ob der Angemeldete das Programm ausführen darf.
$abbrechen = True
For $i=0 to $Benutzer[0]
	If $Benutzer[$i] = $Angemeldet Then _
		$abbrechen = False
Next
If $abbrechen Then
    MsgBox($MB_ICONERROR,"Kein Zugriff","Keine Autorisierung für diesen Benutzer")
	Exit
 EndIf

;~ Alle SAP-Fenster schließen, SAP ausführen, Einloggen, Eingabevariante auswählen, Datum eingeben, mb5b starten
MsgBox($MB_ICONINFORMATION, "Programmstart", "Alle SAP-Fenster werden geschlossen!" & @LF & "Bitte stecken Sie Ihren Ausweis in den Kartenleser." & @LF & "Machen Sie keine Maus- oder Tastatureingaben während das Programm läuft wenn es nicht gefordert ist.")
If WinExists("[CLASS:SAP_FRONTEND_SESSION]","") _
	Then WinKill("[CLASS:SAP_FRONTEND_SESSION]","")
If WinExists("SAP Logon 740") _
	Then WinKill("SAP Logon 740")

ShellExecute("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

WinWait("SAP Logon 740","")
If Not WinActive("SAP Logon 740","") _
	Then WinActivate("SAP Logon 740","")
WinWaitActive("SAP Logon 740","")

ControlClick("SAP Logon 740", "", 1068)

WinWait("P41(1)/101 SAP Easy Access","")
If Not WinActive("P41(1)/101 SAP Easy Access","") _
	Then WinActivate("P41(1)/101 SAP Easy Access","")
WinWaitActive("P41(1)/101 SAP Easy Access","")

MsgBox($MB_ICONINFORMATION, "Auswahl", "F2 = Jahresauswertung"  & @LF & "F3 = Monatsauswertung"  & @LF & "F4 = Programm beenden (Gültig während des gesamten Programmablaufs)" & @LF & @LF & "Erst OK drücken, danach die gewünschte Auswahltaste")
While 1							;wartet bis eine Eingabe gemacht wurde
   if _IsPressed("71") Then
	  $end_datum = InputBox("Jahresauswertung","Geben Sie das Enddatum an." & @LF & "Beenden des Programms mit der Taste F4" & @LF & @LF & "Format: DD.MM.YYYY","DD.MM.YYYY")
	  $datum = "{DOWN 8}{TAB}" & $end_datum & "{f8}{ENTER}"
	  $zusatz_name = "G_"
	  ExitLoop
   ElseIf _IsPressed("72") Then
	  $start_datum = InputBox("Monatsauswertung","Geben Sie das Startdatum an." & @LF & @LF & "Format: DD.MM.YYYY","DD.MM.YYYY")
	  $end_datum = InputBox("Monatsauswertung","Geben Sie das Enddatum an." & @LF & @LF & "Format: DD.MM.YYYY","DD.MM.YYYY")
	  $datum = "{DOWN 7}{TAB}" & $start_datum & "{TAB}" & $end_datum & "{f8}{ENTER}"
	  $zusatz_name = "M_"
	  ExitLoop
   EndIf
Wend

Send("mb5b{ENTER}")

;~ Geht alle Varianten durch und Speichert die Tabellen als Excel-Listen ab
for $j = 0 to $Var_Anzahl - 1

WinWait("P41(1)/101 Bestand zum Buchungsdatum","")
If Not WinActive("P41(1)/101 Bestand zum Buchungsdatum","") _
	Then WinActivate("P41(1)/101 Bestand zum Buchungsdatum","")
WinWaitActive("P41(1)/101 Bestand zum Buchungsdatum","")

Send("+{f5}")

WinWait("P41(1)/101 Variante suchen","")
If Not WinActive("P41(1)/101 Variante suchen","") _
	Then WinActivate("P41(1)/101 Variante suchen","")
WinWaitActive("P41(1)/101 Variante suchen","")

Send("{TAB 4}foitrou9{f8}")

WinWait("P41(1)/101 ABAP: Variantenkatalog des Programms RM07MLBD","")
If Not WinActive("P41(1)/101 ABAP: Variantenkatalog des Programms RM07MLBD","") _
	Then WinActivate("P41(1)/101 ABAP: Variantenkatalog des Programms RM07MLBD","")
WinWaitActive("P41(1)/101 ABAP: Variantenkatalog des Programms RM07MLBD","")

Send($Variante_waehlen[$j])

WinWait("P41(1)/101 Bestand zum Buchungsdatum","")
If Not WinActive("P41(1)/101 Bestand zum Buchungsdatum","") _
	Then WinActivate("P41(1)/101 Bestand zum Buchungsdatum","")
WinWaitActive("P41(1)/101 Bestand zum Buchungsdatum","")

Send($datum)
Sleep(100)
Send("{ENTER}")

WinWait("P41(1)/101 Materialbestände zwischen","")
If Not WinActive("P41(1)/101 Materialbestände zwischen","") _
	Then WinActivate("P41(1)/101 Materialbestände zwischen","")
WinWaitActive("P41(1)/101 Materialbestände zwischen","")

Send("{LAlt Down}LSD{LAlt Up}")

WinWait("P41(1)/101 Liste sichern in Datei...","")
If Not WinActive("P41(1)/101 Liste sichern in Datei...","") _
	Then WinActivate("P41(1)/101 Liste sichern in Datei...","")
WinWaitActive("P41(1)/101 Liste sichern in Datei...","")

Send("{DOWN}{ENTER}")

WinWait("P41(1)/101 Materialbestände zwischen","")
If Not WinActive("P41(1)/101 Materialbestände zwischen","") _
	Then WinActivate("P41(1)/101 Materialbestände zwischen","")
WinWaitActive("P41(1)/101 Materialbestände zwischen","")

Send($zusatz_name & $Dateinamen[$j])
Send("+{Tab}+{END}" & $Pfad_data)
Send("^s")

WinWait("P41(1)/101 Materialbestände zwischen","")
If Not WinActive("P41(1)/101 Materialbestände zwischen","") _
	Then WinActivate("P41(1)/101 Materialbestände zwischen","")
WinWaitActive("P41(1)/101 Materialbestände zwischen","")

Send("{F3}")
Next

;~ Beendet alle offenen SAP-Fenster
If WinExists("[CLASS:SAP_FRONTEND_SESSION]","") _
	Then WinKill("[CLASS:SAP_FRONTEND_SESSION]","")
If WinExists("SAP Logon 740") _
	Then WinKill("SAP Logon 740")



