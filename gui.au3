;~ WinActivate("Unbenannt - Paint")

hotkeyset("{f1}","stop")
hotkeyset("{f2}","pause")
$color=0x00FF00
$pause=false

while 1
;~    WinActivate("Unbenannt - Paint")
	$head=PixelSearch(600,300,1200,600,$color,0,1)
	if not @error then
	    MouseMove($head[0], $head[1])
		mouseclick ("left",$head[0],$head[1],1,0)
	EndIf
;~ 	sleep (5)
WEnd

func stop ()
	Exit
EndFunc

func pause()
	$pause=not $pause
	if $pause=true Then
		Do
			sleep (5)
		until $pause=False
	EndIf
EndFunc