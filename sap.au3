;MsgBox(0, "SAP-Fenster", "SAP wird beendet, alle SAP-Fenster werden geschlossen!!!")
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

;WinWait("Windows-Sicherheit","")
;If Not WinActive("Windows-Sicherheit","") _
;	Then WinActivate("Windows-Sicherheit","")
;WinWaitActive("Windows-Sicherheit","")

WinWait("P41(1)/101 SAP Easy Access","")
If Not WinActive("P41(1)/101 SAP Easy Access","") _
	Then WinActivate("P41(1)/101 SAP Easy Access","")
WinWaitActive("P41(1)/101 SAP Easy Access","")

Send("mb5b{ENTER}")

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

Send("{RIGHT}{f2}")

WinWait("P41(1)/101 Bestand zum Buchungsdatum","")
If Not WinActive("P41(1)/101 Bestand zum Buchungsdatum","") _
	Then WinActivate("P41(1)/101 Bestand zum Buchungsdatum","")
WinWaitActive("P41(1)/101 Bestand zum Buchungsdatum","")

Send("{DOWN 8}{TAB}31.10.2017{f8}{ENTER}")
Sleep(200)
Send("{ENTER}")

WinWait("P41(1)/101 Materialbestände zwischen 01.10.2017 und 31.10.2017","")
If Not WinActive("P41(1)/101 Materialbestände zwischen 01.10.2017 und 31.10.2017","") _
	Then WinActivate("P41(1)/101 Materialbestände zwischen 01.10.2017 und 31.10.2017","")
WinWaitActive("P41(1)/101 Materialbestände zwischen 01.10.2017 und 31.10.2017","")

Send("{LAlt Down}LSD{LAlt Up}")

WinWait("P41(1)/101 Liste sichern in Datei...","")
If Not WinActive("P41(1)/101 Liste sichern in Datei...","") _
	Then WinActivate("P41(1)/101 Liste sichern in Datei...","")
WinWaitActive("P41(1)/101 Liste sichern in Datei...","")

Send("{DOWN}{ENTER}")

WinWait("P41(1)/101 Materialbestände zwischen 01.10.2017 und 31.10.2017","")
If Not WinActive("P41(1)/101 Materialbestände zwischen 01.10.2017 und 31.10.2017","") _
	Then WinActivate("P41(1)/101 Materialbestände zwischen 01.10.2017 und 31.10.2017","")
WinWaitActive("P41(1)/101 Materialbestände zwischen 01.10.2017 und 31.10.2017","")

Send("test.xls^s")
sleep(200)

If WinExists("[CLASS:SAP_FRONTEND_SESSION]","") _
	Then WinKill("[CLASS:SAP_FRONTEND_SESSION]","")
If WinExists("SAP Logon 740") _
	Then WinKill("SAP Logon 740")