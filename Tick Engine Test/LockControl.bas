Attribute VB_Name = "LockControl"
Public Sub setWoWLock(strLocked As Boolean)
Dim strFile As String
strFile = "<%" & vbCrLf & "Dim blnTickRunning" & vbCrLf & "blnTickRunning = cbool(" & strLocked & ")" & vbCrLf & "%>"
Dim FSO
Dim wFile
Set FSO = CreateObject("Scripting.FileSystemObject")
Set wFile = FSO.createTextFile(strLockFile)
wFile.Write (strFile)
wFile.Close

Set wFile = Nothing
Set FSO = Nothing
End Sub
