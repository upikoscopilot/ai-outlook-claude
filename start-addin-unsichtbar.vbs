' Startet den Claude-Mail-Server unsichtbar im Hintergrund (kein Fenster).
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = dir
node = sh.ExpandEnvironmentStrings("%ProgramFiles%") & "\nodejs\node.exe"
If Not fso.FileExists(node) Then node = "node"
' Läuft schon ein Server auf Port 3000? Dann nichts tun.
Set ex = sh.Exec("cmd /c netstat -ano | findstr /R /C:"":3000 .*LISTENING""")
If Len(ex.StdOut.ReadAll()) > 0 Then WScript.Quit
sh.Run """" & node & """ """ & dir & "\serve.mjs""", 0, False
