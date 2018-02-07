#include <Excel.au3>
HotKeySet("{F4}", "ExitProg")
Func ExitProg()
    MsgBox($MB_ICONWARNING, "Beendet", "Programm beendet")
    Exit 0
EndFunc

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

Dim $Var_Anzahl = Ubound($Dateinamen)		;Zählt die Anzahl der vorhandenen Varianten
Dim $l				;Zähler


for $l = 0 to $Var_Anzahl - 1
Dim $Datei_Z	= "Z_" & $Dateinamen[$l]
ShellExecute($Datei_Z, "", "\\ww005.siemens.net\meddfsroot\1_MR\1_HQMR\Teams\Produktdaten\Tobi\Lieferzahlen\Auswertung_Tobi\Verbrauch")
WinWait("Microsoft Excel",$Datei_Z)
If Not WinActive("Microsoft Excel",$Datei_Z) _
	Then WinActivate("Microsoft Excel",$Datei_Z)
WinWaitActive("Microsoft Excel",$Datei_Z)

Send("!{F8}")

WinWait("Makro")
If Not WinActive("Makro") _
	Then WinActivate("Makro")
WinWaitActive("Makro")

Send("Einspielen{Enter}")
while WinExists("Microsoft Excel")
   MsgBox(0,"t","Dn!")
sleep(250)

wend
Next

MsgBox(0,"Ende","Das Programm ist erfolgreich durchgelaufen!")