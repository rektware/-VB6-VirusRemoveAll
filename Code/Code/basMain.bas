Attribute VB_Name = "basMain"
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Function AppPath()
AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
End Function
Public Function FileExists(sFile) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

Public Function GetFileName(ByVal sPath As String) As String
On Error Resume Next
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function
Public Function GetFolderPath(ByVal sPath As String) As String
On Error Resume Next
GetFolderPath = Left(sPath, InStrRev(sPath, "\") - 1)
End Function

Public Function GetFileExt(ByVal sPath) As String
On Error Resume Next
GetFileExt = Mid(sPath, InStrRev(sPath, ".") + 1, Len(sPath) - InStr(1, sPath, "."))
End Function

Public Function GetKeyGoc(sKeyString)
On Error Resume Next
GetKeyGoc = Left(sKeyString, Len(sKeyString) - InStrRev(StrReverse(sKeyString), "\"))
End Function
Public Function GetKeyName(sKeyString)
On Error Resume Next
GetKeyName = Right(sKeyString, Len(sKeyString) - InStrRev(sKeyString, "\"))
End Function
Public Function GetKeyPath(sKeyString)
On Error Resume Next
GetKeyPath = Mid(sKeyString, 2 + Len(sKeyString) - InStrRev(StrReverse(sKeyString), "\"), InStrRev(sKeyString, "\") - (Len(sKeyString) - InStrRev(StrReverse(sKeyString), "\") + 2))
End Function


Public Function KillProcess(sID As Long) As Boolean
On Error Resume Next
DoEvents
KillProcess = True
KillProcessById sID
If CheckID(sID) <> 0 Then KillProcess = False
End Function

Public Function KillFile(sFilePath As String) As Boolean
On Error Resume Next
DoEvents
KillProcess CheckProcess(sFilePath)
KillFile = True
SetAttr sFilePath, vbNormal
DeleteFile sFilePath
If FileExists(sFilePath) = True Then KillFile = False
End Function

Public Function ReadFileUni(FileName As String) As String
On Error Resume Next
Dim FSO
   Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 1, , -2)
   ReadFileUni = FSO.Readall
   Set FSO = Nothing
End Function
Public Function WriteFileUni(FileName As String, Unistr As String)
On Error Resume Next
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject").CreateTextFile(FileName, True)
Set FSO = Nothing
Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 2, , -1)
    FSO.write Unistr
Set FSO = Nothing
End Function

