Attribute VB_Name = "modLietKeValue"
Option Explicit
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_ALL_ACCESS = &HF003F
Private Const REG_SZ = 1
Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Dim RetVal As Long
Dim hKey As Long
Dim NameKey As String
Dim lpType As Long
Dim LenName As Long
Dim Data(0 To 255) As Byte
Dim DataLen As Long
Dim DataString As String
Dim Index As Long
Dim i As Long
Dim KetQua As String
Public xTotalStartUp
Public Function GetKeyValue(FullKeyName)
xTotalStartUp = 0
Dim Key1, Key2, i, Ua
Ua = 10
For i = 1 To Len(FullKeyName)
    If Mid(FullKeyName, i, 1) = "\" Then
        Ua = Ua + 10
        If Ua = 20 Then
            Key1 = Left(FullKeyName, i - 1)
            Key2 = Right(FullKeyName, Len(FullKeyName) - i)
        End If
    End If
Next i
'frmMain.Cls
If Key1 = "HKEY_LOCAL_MACHINE" Then
RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Key2, 0, KEY_ALL_ACCESS, hKey)
ElseIf Key1 = "HKEY_CURRENT_USER" Then
RetVal = RegOpenKeyEx(HKEY_CURRENT_USER, Key2, 0, KEY_ALL_ACCESS, hKey)
End If

Index = 0
Do While RetVal = 0
    NameKey = Space(255)
    DataString = Space(255)
    LenName = 255
    DataLen = 255
    RetVal = RegEnumValue(hKey, Index, NameKey, LenName, ByVal 0, lpType, Data(0), DataLen)
    If RetVal = 0 Then
        NameKey = Left(NameKey, LenName) 'Rút b? kho?n tr?ng th?a
        DataString = ""
' X? lý thông tin theo ki?u c?a nó và ??a vào bi?n DataString
        Select Case lpType
             Case REG_SZ
                For i = 0 To DataLen - 1
                    DataString = DataString & Chr(Data(i)) ' N?i các ch? cái thành chu?i
                Next
             Case REG_BINARY
                For i = 0 To DataLen - 1
                    Dim temp As String
                    temp = Hex(Data(i))
                    If Len(temp) < 2 Then temp = String(2 - Len(temp), "0") & temp
                    DataString = DataString & temp & " "
 ' N?i các c?p s? nh? phân l?i v?i nhau
                Next
            Case REG_DWORD
                For i = DataLen - 1 To 0 Step -1
                    DataString = DataString & Hex(Data(i)) 'N?i các sô hexa v?i nhau
                Next
            Case REG_MULTI_SZ
                For i = 0 To DataLen - 1
                    DataString = DataString & Chr(Data(i))
    'N?i các ký t? bao g?m ký t? vbNullChar (?? cách dòng) thành m?t chu?i, b?n có th? s? d?ng m?t m?ng g?m nhi?u string thay vì là m?t
                Next
            Case REG_EXPAND_SZ
                For i = 0 To DataLen - 2
                    DataString = DataString & Chr(Data(i))
    'N?i các ký t? l?i v?i nhau, b? ký t? NULL cu?i cùng
                Next
            Case Else
                DataString = " Khong xac dinh duoc !"
        ' Trên ?ây là 5 ki?u có trên WinXP
        End Select
    End If
    If Left(Left(NameKey, LenName), 1) <> " " Then
    '///////////////////
    'Form1.List1.AddItem DataString
    If xCheckVirus(DataString) = True Then
        With frmMain.LVREG
            Dim h
            h = .Count + 1
            .ItemAdd h, Left(NameKey, LenName), 0, 0
            .SubItemSet h, 1, GetFileP(DataString), 0
            .SubItemSet h, 2, FullKeyName & "\" & Left(NameKey, LenName), 0
        End With
    End If
    '///////////////
    End If
    Index = Index + 1
    'frmMain.Print Left(NameKey, LenName) & "=" & DataString
Loop
RetVal = RegCloseKey(hKey)
End Function

Public Function GetFileName(ByVal sPath As String) As String
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function
Public Function GetFolderPath(ByVal sPath As String) As String
GetFolderPath = Left(sPath, InStrRev(sPath, "\") - 1)
End Function

Public Sub GetSystemKey()
   ' With frmMain.LV
   '     Dim iu
   '     iu = .ListItems.Count + 1
   '     .ListItems.Add iu, , ToUnicode("Shell [He65 Tho61ng]")
   '     .ListItems(iu).SubItems(1).Caption = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell")
   '     .ListItems(iu).SubItems(2).Caption = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Shell"
   '     iu = .ListItems.Count + 1
   '     .ListItems.Add iu, , ToUnicode("Userinit [He65 Tho61ng]")
   '     .ListItems(iu).SubItems(1).Caption = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit")
   '     .ListItems(iu).SubItems(2).Caption = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Userinit"
   ' End With
   With frmMain.LVREG
        Dim iu
        iu = .Count + 1
   
        If xCheckVirus("C:\WINDOWS\explorer.exe") = True Then
            .ItemAdd iu, "Shell", 0, 0
            .SubItemSet iu, 1, GetFileP(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell")), 0
            .SubItemSet iu, 2, "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Shell", 0
        End If
        iu = iu + 1
        If xCheckVirus("C:\WINDOWS\system32\userinit.exe") = True Then
            .ItemAdd iu, "Userinit", 0, 0
            .SubItemSet iu, 1, GetFileP(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit")), 0
            .SubItemSet iu, 2, "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Userinit", 0
        End If
    End With
End Sub

Function GetFileP(sFile)
GetFileP = sFile
If Left$(GetFileP, 1) = ChrW(34) Then
    GetFileP = Mid$(sFile, 2, InStr(2, sFile, ChrW(34)) - 2)
End If
End Function
