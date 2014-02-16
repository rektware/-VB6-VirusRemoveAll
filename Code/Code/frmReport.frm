VERSION 5.00
Object = "{84B0A18C-0BF5-429A-953B-A6EACF525624}#17.0#0"; "UnicodeFullControl.ocx"
Begin VB.Form frmReport 
   BackColor       =   &H80000005&
   Caption         =   "Virus Reporter"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   8160
   StartUpPosition =   1  'CenterOwner
   Begin UnicodeControl.UniCommonDialog Dialog1 
      Left            =   3960
      Top             =   4800
      _ExtentX        =   423
      _ExtentY        =   423
      DialogTitle     =   "frmReport.frx":0A02
      Filename        =   "frmReport.frx":0A1A
      Filter          =   "frmReport.frx":0A32
      FilterIndex     =   1
      hDC             =   16842838
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UnicodeControl.UniTextBox txtRe 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11456
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Text            =   "frmReport.frx":0A6A
      BorderStyle     =   2
      UniToolTipText  =   "frmReport.frx":0A82
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Begin VB.Menu save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu close 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
Unload Me
End Sub

Private Sub Form_Resize()
'On Error Resume Next
If Me.WindowState <> 1 Then
    txtRe.Left = 120
    txtRe.Top = 120
    txtRe.Height = Me.Height - 1050
    txtRe.Width = Me.Width - 360
End If
End Sub



Private Sub Form_Unload(Cancel As Integer)

Cancel = 1
Me.Hide

End Sub

Private Sub save_Click()
Dialog1.FileName = ""
Dialog1.Filter = "Text File (*.txt)|*.txt"
Dialog1.ShowOpen
If Dialog1.FileName <> "" Then
    WriteFileUni Dialog1.FileName, txtRe.Text
    UnicodeMsgBox UnicodeText("D9a4 lu7u!"), vbOKOnly + vbInformation, "OK!"
End If
End Sub
