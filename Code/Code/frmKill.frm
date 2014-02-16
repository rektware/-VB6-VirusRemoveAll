VERSION 5.00
Object = "{84B0A18C-0BF5-429A-953B-A6EACF525624}#17.0#0"; "UnicodeFullControl.ocx"
Begin VB.Form frmKill 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kill Virus"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   120
   End
   Begin UnicodeControl.ProgressBar B1 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      Scrolling       =   2
   End
   Begin UnicodeControl.UniLabel UniLabel5 
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   3960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   873
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      BorderStyle     =   0
      WordWrap        =   -1  'True
      Caption         =   "frmKill.frx":058A
      Link            =   "myspecialbox@yahoo.com.vn"
      ForeColorWordEffect=   0
      SpeedOrtherColor=   0
      UniToolTipText  =   "frmKill.frx":0660
   End
   Begin UnicodeControl.UniLabel UniLabel4 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      BorderStyle     =   0
      Caption         =   "frmKill.frx":0678
      Link            =   "myspecialbox@yahoo.com.vn"
      ForeColorWordEffect=   0
      SpeedOrtherColor=   0
      UniToolTipText  =   "frmKill.frx":06A4
   End
   Begin UnicodeControl.UniLabel UniLabel3 
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1085
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      BorderStyle     =   0
      WordWrap        =   -1  'True
      Caption         =   "frmKill.frx":06BC
      Link            =   "myspecialbox@yahoo.com.vn"
      ForeColorWordEffect=   0
      SpeedOrtherColor=   0
      UniToolTipText  =   "frmKill.frx":07D4
   End
   Begin UnicodeControl.UniCheck cmdFix 
      Height          =   240
      Index           =   4
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   423
      Value           =   1
      Caption         =   "frmKill.frx":07EC
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   0
      AutoSize        =   -1  'True
      UniToolTipText  =   "frmKill.frx":0860
   End
   Begin UnicodeControl.UniCheck cmdFix 
      Height          =   240
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   423
      Value           =   1
      Caption         =   "frmKill.frx":0878
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   0
      AutoSize        =   -1  'True
      UniToolTipText  =   "frmKill.frx":08E8
   End
   Begin UnicodeControl.UniCheck cmdFix 
      Height          =   240
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   423
      Value           =   1
      Caption         =   "frmKill.frx":0900
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   0
      AutoSize        =   -1  'True
      UniToolTipText  =   "frmKill.frx":0966
   End
   Begin UnicodeControl.UniCheck cmdFix 
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   423
      Value           =   1
      Caption         =   "frmKill.frx":097E
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   0
      AutoSize        =   -1  'True
      UniToolTipText  =   "frmKill.frx":09E0
   End
   Begin UnicodeControl.UniCheck cmdFix 
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   423
      Value           =   1
      Caption         =   "frmKill.frx":09F8
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   0
      AutoSize        =   -1  'True
      UniToolTipText  =   "frmKill.frx":0A7A
   End
   Begin UnicodeControl.UniLabel UniLabel2 
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      BorderStyle     =   0
      WordWrap        =   -1  'True
      Caption         =   "frmKill.frx":0A92
      Link            =   "myspecialbox@yahoo.com.vn"
      ForeColorWordEffect=   0
      SpeedOrtherColor=   0
      UniToolTipText  =   "frmKill.frx":0B74
   End
   Begin UnicodeControl.UniLabel UniLabel1 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      BorderStyle     =   0
      Alignment       =   7
      Caption         =   "frmKill.frx":0B8C
      Link            =   "myspecialbox@yahoo.com.vn"
      ForeColorWordEffect=   0
      SpeedOrtherColor=   0
      UniToolTipText  =   "frmKill.frx":0BC0
   End
   Begin UnicodeControl.UniButton cmdReport 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      Caption         =   "frmKill.frx":0BD8
      ForecolorSelected=   0
      PictureNormal   =   "frmKill.frx":0C1A
      PictureAlignment=   4
      PictureSize     =   18
      UniToolTipText  =   "frmKill.frx":162C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableDoubleClick=   -1  'True
   End
   Begin UnicodeControl.UniButton cmdKill 
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   5040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      Caption         =   "frmKill.frx":1644
      ForeColor       =   255
      ForecolorSelected=   192
      PictureNormal   =   "frmKill.frx":1676
      PictureAlignment=   4
      PictureSize     =   18
      UniToolTipText  =   "frmKill.frx":1F50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableDoubleClick=   -1  'True
   End
   Begin UnicodeControl.UniButton cmdNoKill 
      Height          =   735
      Left            =   4560
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      Caption         =   "frmKill.frx":1F68
      ForecolorSelected=   0
      PictureNormal   =   "frmKill.frx":1FA2
      PictureAlignment=   4
      PictureSize     =   18
      UniToolTipText  =   "frmKill.frx":29B4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableDoubleClick=   -1  'True
   End
   Begin UnicodeControl.ImageXP ImageXP1 
      Height          =   1935
      Left            =   3960
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3413
      SImage          =   "frmKill.frx":29CC
   End
End
Attribute VB_Name = "frmKill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFix_Click(Index As Integer)
cmdFix(0).Checked = True
End Sub

Private Sub cmdKill_Click()
Timer1.Enabled = True
DisableAll
'Bu'o'c 1: Xóa các tiê'n trình
'     - Dong' Bang Neu Can Thiet
BatDauXoaPro:
Dim UyA As Integer
For UyA = 0 To frmMain.LVPRO.Count - 1
DoEvents
    'frmMain.LVPRO.SubItemText(Uy, 2)
    SuspendResumeProcess frmMain.LVPRO.SubItemText(UyA, 2), True
Next UyA

'        - Xoa'
Dim UyX As Integer
DoEvents
For UyX = 0 To frmMain.LVPRO.Count - 1
DoEvents
    'MsgBox LVPRO.SubItemText(0, 2)
    KillProcess frmMain.LVPRO.SubItemText(UyX, 2)
    KillFile (frmMain.LVPRO.SubItemText(UyX, 1))
Next UyX


Dim UyI As Integer
DoEvents
For UyI = 0 To frmMain.LVPRO.Count - 1
DoEvents
    If FileExists(frmMain.LVPRO.SubItemText(UyI, 1)) = False Then
    frmMain.LVPRO.ItemRemove UyI
    GoTo BatDauXoaPro
    End If
Next UyI
'-----------------------------------

'Bu'o'c 2: Xóa key trong Registry
Dim Ui As Integer
DoEvents
For Ui = 0 To frmMain.LVREG.Count - 1
DoEvents
    KillFile frmMain.LVREG.SubItemText(Ui, 1)
    If UCase(GetKeyGoc(frmMain.LVREG.SubItemText(Ui, 2))) = "HKEY_CURRENT_USER" Then
        DeleteValue HKEY_CURRENT_USER, GetKeyPath(frmMain.LVREG.SubItemText(Ui, 2)), GetKeyName(frmMain.LVREG.SubItemText(Ui, 2))
    ElseIf UCase(GetKeyGoc(frmMain.LVREG.SubItemText(Ui, 2))) = "HKEY_LOCAL_MACHINE" Then
        DeleteValue HKEY_LOCAL_MACHINE, GetKeyPath(frmMain.LVREG.SubItemText(Ui, 2)), GetKeyName(frmMain.LVREG.SubItemText(Ui, 2))
    End If
Next Ui
BatDauXoaReg:
Dim UiI As Integer
DoEvents
For UiI = 0 To frmMain.LVREG.Count - 1
DoEvents
    frmMain.LVREG.ItemRemove UiI
    GoTo BatDauXoaReg
Next UiI
'-----------------------------------

'Bu'o'c 3: Xóa file trong hê. thô'ng

Dim Uo As Integer

For Uo = 0 To frmMain.LVFILE.Count - 1
DoEvents
    KillFile frmMain.LVFILE.ItemText(Uo)
Next Uo

BaTdAuXoAfIleHaI:
Dim UoI As Integer

For UoI = 0 To frmMain.LVFILE.Count - 1
DoEvents
    If FileExists(frmMain.LVFILE.ItemText(UoI)) = False Then
        frmMain.LVFILE.ItemRemove UoI
        GoTo BaTdAuXoAfIleHaI
    End If
Next UoI

'------------------------------------


'Bu'o'c 4: Thu'c hiê.n phu.c hoooo`i hê. thô'ng
If cmdFix(1).Value = vbChecked Then RegistryClean

If cmdFix(2).Value = vbChecked Then
    KillFile "C:\WINDOWS\system32\drivers\etc\hosts"
    WriteFileUni "C:\WINDOWS\system32\drivers\etc\hosts", "127.0.0.1       localhost"
End If

If cmdFix(3).Value = vbChecked Then

Dim Ni As Integer
For Ni = 0 To frmMain.Drive1.ListCount - 1
DoEvents
    If UCase(Left(frmMain.Drive1.List(Ni), 1)) <> "A" Then
DoEvents
        KillFile Left(frmMain.Drive1.List(Ni), 2) & "\autorun.inf"
    End If
Next Ni

End If
EnableAll
Timer1.Enabled = False
B1.Value = 0

UnicodeMsgBox UnicodeText("D9a4 xo1a ca1c mu5c d9a4 d9a1nh da61u!"), vbOKOnly + vbInformation, UnicodeText("D9a4 die65t!")

End Sub

Private Sub cmdNoKill_Click()
With frmMain
    .cmdAdd.Enabled = True
    .cmdDel.Enabled = True
    .tmrReg.Enabled = False
    .cmdStart.Enabled = True
    .cmdOption.Enabled = False
End With
Unload frmReport
Unload Me
End Sub

Private Sub DisableAll()

Me.cmdKill.Enabled = False
Me.cmdNoKill.Enabled = False
Me.cmdReport.Enabled = False
End Sub
Private Sub EnableAll()

Me.cmdKill.Enabled = True
Me.cmdNoKill.Enabled = True
Me.cmdReport.Enabled = True
End Sub


Private Sub cmdReport_Click()
frmReport.Show
End Sub

Private Sub Timer1_Timer()
B1.Value = B1.Value + 3
If B1.Value > 99 Then B1.Value = 0
End Sub
