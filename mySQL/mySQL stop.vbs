strFilename =  Wscript.ScriptName
'Wscript.Echo "Script name: " & strFilename 

 i = InStrRev(strFilename , ".")
  If ( i > 0 ) Then
    Result = Mid(strFilename , 1, i - 1)
Else
Result = strFilename 
  End If
  strFilename = Result & ".bat"
Wscript.Echo "RUN BAT FILE: " & strFilename 

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objFile = objFSO.GetFile(strFilename )
'Wscript.Echo "File name: " & objFSO.GetFileName(objFile)
'Wscript.Echo "File name: " & objFSO.ShortName(objFile)

'WScript.Quit 1

Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & strFilename  & Chr(34), 0
Set WshShell = Nothing
