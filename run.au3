;local $sDestination = "C:\Users\z002u7wh\Desktop\AutoIT - TestCode\Tulips.jpg"
HotKeySet("{F4}", "ExitProg")
Func ExitProg()
    Exit 0
EndFunc


SplashImageOn("bild", "C:\Users\z002u7wh\Desktop\AutoIT - TestCode\Tulips.jpg")

sleep(1000)
SplashOff()
$var = Run("notepad.exe")
;MsgBox(0, "Return value is", $var)


sleep(1000)
ProcessClose ( $var )

