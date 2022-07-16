On Error Resume Next

strFile = WScript.Arguments(0)
vWaitTime = 300

SET objFSO = CREATEOBJECT("Scripting.FileSystemObject")
SET objFile = objFSO.GetFile(strFile)

Do
	Wscript.Sleep (vWaitTime*1000)
	If DateDiff("s", CDATE(objFile.DateLastModified), Now) >= vWaitTime Then
		Set objShell = WScript.CreateObject("WScript.Shell")
		objShell.Run "taskkill /f /im iexplore.exe", , True
		Set objShell = Nothing
	End If
Loop While objFSO.FileExists(strFile)

SET objFile = Nothing
SET objFSO = Nothing