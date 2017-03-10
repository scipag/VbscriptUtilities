Set objShell = CreateObject("WScript.Shell")

delay = Inputbox("In how many minutes shall we sleep?", "Delay to Sleep", "10")

WScript.Sleep (delay * 1000 * 60)

objShell.Run("cmd /c %windir%\system32\rundll32.exe powrprof.dll,SetSuspendState Sleep")
