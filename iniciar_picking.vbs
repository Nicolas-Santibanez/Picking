Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Obtener la carpeta actual donde esta guardado este script VBS
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = scriptPath

' Ejecutar el batch de forma 100% invisible (el 0 final)
WshShell.Run chr(34) & scriptPath & "\lanzar_picking.bat" & chr(34), 0, False