Attribute VB_Name = "basCheckVR"
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Public Function xCheckVirus(sFile As String) As Boolean
xCheckVirus = False
If FileExists(sFile) = True Then
    If FileLen(sFile) > 5000000 Then Exit Function
    Dim i
    For i = 0 To frmMain.lstFile.Count - 1
        If GetMD5(sFile) = frmMain.lstFile.SubItemText(i, 1) Then xCheckVirus = True
    Next i
End If
End Function

Public Sub xCreateReport()
On Error Resume Next
Dim xNoiDung As String
xNoiDung = UnicodeText("Ba3n pha6n ti1ch Virus." & vbCrLf _
& "Thu75c hie65n bo73i: Virus Remove All 1.0" & vbCrLf _
& "Tho72i gian: " & Time & " - " & Date & vbCrLf & "-------------------------" & vbCrLf & vbCrLf & vbCrLf & "[o] Tho6ng tin ve62 ma64u Virus: ") & vbCrLf

Dim U
For U = 0 To frmMain.lstFile.Count - 1
    xNoiDung = xNoiDung & "      [+] " & basMain.GetFileName(frmMain.lstFile.ItemText(U)) & vbCrLf & UnicodeText("            [-] Dung Lu7o75ng: ") & FileLen(frmMain.lstFile.ItemText(U)) & " Bytes" & vbCrLf & UnicodeText("            [-] Thuo65c Ti1nh: ") & GetAttribute(frmMain.lstFile.ItemText(U)) & vbCrLf
Next U

xNoiDung = xNoiDung & vbCrLf & vbCrLf & UnicodeText("[o] Tho6ng tin ve62 ta1c d9o65ng cu3a Virus:" & vbCrLf & "      [+] Ghi va2o Registry ca1c kho1a sau:") & vbCrLf
Dim Y
For Y = 0 To frmMain.LVREG.Count - 1
    xNoiDung = xNoiDung & "             [-] " & GetKeyGoc(frmMain.LVREG.SubItemText(Y, 2)) & "\" & GetKeyPath(frmMain.LVREG.SubItemText(Y, 2)) & UnicodeText(" | Gia1 Tri5: ") & frmMain.LVREG.ItemText(Y) & vbCrLf
Next Y

xNoiDung = xNoiDung & UnicodeText("      [+] Ca1c tie61n tri2nh d9u7o75c ki1ch hoa5t khi nhie64m Virus:") & vbCrLf
Dim O
For O = 0 To frmMain.LVPRO.Count - 1
    xNoiDung = xNoiDung & "            [-] " & frmMain.LVPRO.SubItemText(O, 1) & vbCrLf
Next O

xNoiDung = xNoiDung & UnicodeText("      [+] Vi5 tri1 ca1c File cu3a Virus:") & vbCrLf
Dim K
For K = 0 To frmMain.LVFILE.Count - 1
    xNoiDung = xNoiDung & "            [-] " & frmMain.LVFILE.ItemText(K) & vbCrLf
Next K

xNoiDung = xNoiDung & vbCrLf & vbCrLf & UnicodeText("-----------------------" & vbCrLf & "Chu7o7ng tri2nh All Virus Remove d9a4 xo1a/die65t he61t ta61t ca3 ca1c Key kho73i d9o65ng, ca1c file trong ma1y ti1nh d9a4 bi5 nhie64m." & vbCrLf & "Chu7o7ng tri2nh chi3 mo71i o73 phie6n ba3n d9a62u tie6n, ne6n ba3n pha6n ti1ch co2n so7 xa2i va2 chu7a d9a62y d9u3, ra61t mong nha65n d9u7o75c y1 kie61n cu3a ca1c ba5n d9e63 chu7o7ng tri2nh hoa5t d9o65ng to61t ho7n!" & vbCrLf & "Ba5n co1 the63 ta4i chu7o7ng tri2nh Virus Remove All ta5i d9i5a chi3: http://phanmemtiengviet.co.cc")
frmReport.Show
frmReport.txtRe.Text = xNoiDung
End Sub



Public Function GetAttribute(ByVal sFilePath As String) As String
    Select Case GetFileAttributes(sFilePath)
        Case 1: GetAttribute = "Read Only"
        Case 2: GetAttribute = "Hidden"
        Case 3: GetAttribute = "Read Only + Hidden"
        Case 4: GetAttribute = "System"
        Case 5: GetAttribute = "Read Only + System"
        Case 6: GetAttribute = "Hidden + System"
        Case 7: GetAttribute = "Read Only + Hidden + System"
        '-------------------------------------------------'
        Case 32: GetAttribute = "Archive"
        Case 33: GetAttribute = "Read Only + Archive"
        Case 34: GetAttribute = "Hidden + Archive"
        Case 35: GetAttribute = "Read Only + Hidden + Archive"
        Case 36: GetAttribute = "System + Archive"
        Case 37: GetAttribute = "Read Only + System + Archive"
        Case 38: GetAttribute = "HSA"
        Case 39: GetAttribute = "Read Only + Hidden + System + Archive"
        '-------------------------------------------------'
        Case 128: GetAttribute = "Normal"
        '-------------------------------------------------'
        Case Else: GetAttribute = "N/A"
    End Select
End Function
