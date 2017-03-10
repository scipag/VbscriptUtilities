Set objShell = CreateObject("WScript.Shell")
Set objWmi = GetObject("winmgmts:")

process = Inputbox("Which application shall be observed?", "Application to Observe", "nmap.exe")
command = Inputbox("Which command shall be executed after close?", "Command to Execute", "shutdown -s -t 60")

Do
	strWmiq = "select * from Win32_Process where name='" & process & "'"
	Set objQResult = objWmi.Execquery(strWmiq)
	Wscript.Sleep 60000
Loop While objQResult.count > 0

objShell.Run("cmd /c " & command)
