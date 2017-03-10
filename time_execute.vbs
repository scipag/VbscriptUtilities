Set oShell = CreateObject("WScript.Shell")
Set oWmi = GetObject("winmgmts:")

actiontime = Inputbox("At which time shall the event happen?", "Time of Action", "0710")
command = Inputbox("Which command shall be executed after that?", "Command to Execute", "shutdown -s -t 600")

Do
	If actiontime > Replace(FormatDateTime(Time, 4), ":", "") Then
		Wscript.Sleep 30000
	Else
		shutdown = 1
	End If
Loop While shutdown <> 1

Set oShell = WScript.CreateObject("wscript.shell")
oShell.Run("cmd /c " & command)
