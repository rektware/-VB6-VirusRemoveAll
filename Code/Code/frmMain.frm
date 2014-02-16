VERSION 5.00
Object = "{84B0A18C-0BF5-429A-953B-A6EACF525624}#17.0#0"; "UnicodeFullControl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   7680
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox lstEXT 
      Height          =   2790
      Left            =   9600
      TabIndex        =   21
      Top             =   -480
      Visible         =   0   'False
      Width           =   855
   End
   Begin UnicodeControl.UniButton cmdOption 
      Height          =   975
      Left            =   9120
      TabIndex        =   20
      Top             =   2880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1720
      Caption         =   "frmMain.frx":617A
      ForeColor       =   255
      ForecolorSelected=   16711680
      PictureNormal   =   "frmMain.frx":61BA
      PictureHot      =   "frmMain.frx":7EC4
      PictureAlignment=   4
      PictureSize     =   32
      MouseIcon       =   "frmMain.frx":D6B6
      MousePointer    =   99
      Enabled         =   0   'False
      UniToolTipText  =   "frmMain.frx":12EA8
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
   Begin UnicodeControl.UniStatusBar Bar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      Top             =   8625
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Count           =   2
      PanelAlign1     =   0
      PanelText1      =   "frmMain.frx":12EE0
      PanelIconIndex1 =   -1
      PanelType1      =   0
      PanelAlign2     =   0
      PanelText2      =   "frmMain.frx":12F1A
      PanelIconIndex2 =   -1
      PanelType2      =   0
   End
   Begin VB.Timer tmrReg 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   6600
   End
   Begin UnicodeControl.ImageListXP ImageList2 
      Left            =   8160
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      Count           =   3
      FileSize1       =   1148
      ImageType1      =   1
      ListImage1      =   "frmMain.frx":12F4C
      Key1            =   ""
      Tag1            =   ""
      FileSize2       =   1148
      ImageType2      =   1
      ListImage2      =   "frmMain.frx":133E8
      Key2            =   ""
      Tag2            =   ""
      FileSize3       =   1148
      ImageType3      =   1
      ListImage3      =   "frmMain.frx":13884
      Key3            =   ""
      Tag3            =   ""
   End
   Begin UnicodeControl.ImageListXP ImageList1 
      Left            =   3000
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      Count           =   1
      FileSize1       =   1148
      ImageType1      =   1
      ListImage1      =   "frmMain.frx":13D20
      Key1            =   ""
      Tag1            =   ""
   End
   Begin UnicodeControl.UniFrame fm 
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8281
      TitleColor      =   14653050
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
      Picture         =   "frmMain.frx":141BC
      OldHeight       =   4695
      Caption         =   "frmMain.frx":14BCE
      ForeColor       =   -2147483635
      ForeHighlight   =   0
      ShowButton      =   0   'False
      Begin UnicodeControl.UniFrame fmREG 
         Height          =   3855
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6800
         TitleColor      =   14653050
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
         OldHeight       =   3855
         Caption         =   "frmMain.frx":14C08
         ForeColor       =   -2147483635
         ForeHighlight   =   0
         ShowButton      =   0   'False
         Begin UnicodeControl.ProgressBar pREG 
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1680
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            Scrolling       =   1
         End
         Begin UnicodeControl.UniLabel lblx1 
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            BorderStyle     =   0
            Caption         =   "frmMain.frx":14C20
            Link            =   "myspecialbox@yahoo.com.vn"
            ForeColorWordEffect=   0
            SpeedOrtherColor=   0
            UniToolTipText  =   "frmMain.frx":14C52
         End
      End
      Begin UnicodeControl.UniButton cmdStop 
         Height          =   375
         Left            =   8520
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "frmMain.frx":14C6A
         ForecolorSelected=   0
         PictureNormal   =   "frmMain.frx":14C92
         UniToolTipText  =   "frmMain.frx":14CAE
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
      Begin UnicodeControl.UniFrame fmPRO 
         Height          =   3855
         Left            =   3720
         TabIndex        =   13
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6800
         TitleColor      =   14653050
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
         OldHeight       =   3855
         Caption         =   "frmMain.frx":14CC6
         ForeColor       =   -2147483635
         ForeHighlight   =   0
         ShowButton      =   0   'False
         Begin UnicodeControl.UniLabel lblx2 
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            BorderStyle     =   0
            Caption         =   "frmMain.frx":14CDE
            Link            =   "myspecialbox@yahoo.com.vn"
            ForeColorWordEffect=   0
            SpeedOrtherColor=   0
            UniToolTipText  =   "frmMain.frx":14D10
         End
         Begin VB.Timer tmrPro 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   240
            Top             =   2160
         End
         Begin UnicodeControl.ProgressBar pPro 
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1680
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            Scrolling       =   1
         End
      End
      Begin UnicodeControl.UniListView LVFILE 
         Height          =   3855
         Left            =   7320
         TabIndex        =   11
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6800
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
         BorderStyle     =   1
         HideSelection   =   0   'False
         CheckBoxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
      End
      Begin UnicodeControl.UniLabel lblFile 
         Height          =   255
         Left            =   7320
         TabIndex        =   10
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
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
         Caption         =   "frmMain.frx":14D28
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "frmMain.frx":14D6A
      End
      Begin UnicodeControl.UniListView LVPRO 
         Height          =   3255
         Left            =   3720
         TabIndex        =   9
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5741
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
         BorderStyle     =   1
         HideSelection   =   0   'False
         CheckBoxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
      End
      Begin UnicodeControl.UniLabel lblPro 
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
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
         Caption         =   "frmMain.frx":14D82
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "frmMain.frx":14DC4
      End
      Begin UnicodeControl.UniListView LVREG 
         Height          =   3255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   5741
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
         BorderStyle     =   1
         HideSelection   =   0   'False
         CheckBoxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
      End
      Begin UnicodeControl.UniLabel lblReg 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
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
         Caption         =   "frmMain.frx":14DDC
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "frmMain.frx":14E24
      End
   End
   Begin UnicodeControl.UniCommonDialog Dialog1 
      Left            =   9240
      Top             =   1320
      _ExtentX        =   423
      _ExtentY        =   423
      CancelError     =   -1  'True
      DialogTitle     =   "frmMain.frx":14E3C
      Filename        =   "frmMain.frx":14E54
      Filter          =   "frmMain.frx":14E6C
      FilterIndex     =   1
      hDC             =   1191253729
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
   Begin UnicodeControl.UniButton cmdStart 
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   3240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      Caption         =   "frmMain.frx":14EA4
      ForecolorSelected=   0
      PictureNormal   =   "frmMain.frx":14F0E
      PictureAlignment=   1
      UniToolTipText  =   "frmMain.frx":154A8
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
   Begin UnicodeControl.UniButton cmdDel 
      Height          =   495
      Left            =   9120
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "frmMain.frx":154C0
      ForecolorSelected=   0
      PictureNormal   =   "frmMain.frx":154FA
      PictureAlignment=   1
      UniToolTipText  =   "frmMain.frx":15A94
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
   Begin UnicodeControl.UniButton cmdAdd 
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "frmMain.frx":15AAC
      ForecolorSelected=   0
      PictureNormal   =   "frmMain.frx":15AE8
      PictureAlignment=   1
      UniToolTipText  =   "frmMain.frx":16082
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
   Begin UnicodeControl.UniLabel UniLabel1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
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
      Caption         =   "frmMain.frx":1609A
      Link            =   "myspecialbox@yahoo.com.vn"
      ForeColorWordEffect=   0
      SpeedOrtherColor=   0
      UniToolTipText  =   "frmMain.frx":1614C
   End
   Begin UnicodeControl.UniLabel lbl1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2355
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   7
      Caption         =   "frmMain.frx":16164
      BackColor       =   -2147483643
      ThreeDColor1    =   16744576
      ForeColorWordEffect=   0
      SpeedOrtherColor=   0
      Gradient        =   -1  'True
      UniToolTipText  =   "frmMain.frx":161A4
   End
   Begin UnicodeControl.UniTitleBar UniTitleBar1 
      Left            =   8880
      Top             =   480
      _ExtentX        =   661
      _ExtentY        =   609
      Caption         =   "frmMain.frx":161BC
      Icon            =   "frmMain.frx":1621A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Logo            =   "frmMain.frx":1C3A4
   End
   Begin UnicodeControl.UniListView lstFile 
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2566
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const vbDot = 46
Private Const MAXDWORD As Long = &HFFFFFFFF
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type FILE_PARAMS
   bRecurse As Boolean
   sFileRoot As String
   sFileNameExt As String
   sResult As String
   sMatches As String
   Count As Long
End Type

Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim xStopScan As Boolean

Private Function SearchFile(Path, FileName)

   Dim FP As FILE_PARAMS  'holds search parameters
   Dim tstart As Single   'timer var for this routine only
   Dim tend As Single     'timer var for this routine only
   With FP
      .sFileRoot = Path       'start path
      .sFileNameExt = FileName    'file type of interest
      .bRecurse = 1 ' Check1.Value = 1  '1 = recursive search
   End With
   tstart = GetTickCount()
   Call SearchForFiles(FP)
   tend = GetTickCount()

   
End Function


Private Sub GetFileInformation(FP As FILE_PARAMS)
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   Dim SKetQua As String
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt
   hFile = FindFirstFile(sPath, WFD)
   If hFile <> INVALID_HANDLE_VALUE Then
      Do
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = _
                 FILE_ATTRIBUTE_DIRECTORY Then
            FP.Count = FP.Count + 1
            sTmp = TrimNull(WFD.cFileName)
            SKetQua = sRoot & sTmp
            
            'List1.AddItem SKetQua
            'MsgBox SKetQua
            'Text1.Text = Text1.Text & SKetQua & vbCrLf
            '*********************************************
            Bar1.PanelText(2) = SKetQua
            
            If xStopScan = True Then GoTo ThOaTkHoIvOnGfOr
            If FileExists(SKetQua) = True Then
                If xCheckVirus(SKetQua) = True Then
                    Dim Hu As Integer
                    Hu = LVFILE.Count + 1
                    LVFILE.ItemAdd Hu, SKetQua, 0, 2
                End If
            End If
         End If
    DoEvents
      Loop While FindNextFile(hFile, WFD)
ThOaTkHoIvOnGfOr:
      hFile = FindClose(hFile)
   End If

End Sub


Private Sub SearchForFiles(FP As FILE_PARAMS)
  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & "*.*"
   hFile = FindFirstFile(sPath, WFD)
   If hFile <> INVALID_HANDLE_VALUE Then
      Call GetFileInformation(FP)
      Do
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            If FP.bRecurse Then
               If Asc(WFD.cFileName) <> vbDot Then
                  FP.sFileRoot = sRoot & TrimNull(WFD.cFileName)
                  Call SearchForFiles(FP)
               End If
            End If
         End If
      Loop While FindNextFile(hFile, WFD)
      hFile = FindClose(hFile)
   End If
   

End Sub


Private Function QualifyPath(sPath As String) As String
   If Right$(sPath, 1) <> "\" Then
      QualifyPath = sPath & "\"
   Else
      QualifyPath = sPath
   End If
End Function


Private Function TrimNull(startstr As String) As String
   Dim pos As Integer
   pos = InStr(startstr, Chr$(0))
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
   TrimNull = startstr
End Function
Private Sub cmdAdd_Click()
On Error GoTo kEtThUcViLoI
Dialog1.FileName = ""
Dialog1.Filter = "All File (*.*)|*.*"
Dialog1.ShowOpen
If Dialog1.FileName <> "" Then
    If FileLen(Dialog1.FileName) = 0 Or FileLen(Dialog1.FileName) > 100000000 Then
        UnicodeMsgBox UnicodeText("File co1 dung lu7o7ng qua1 nho3 hoa85c qua1 lo71n!"), vbOKOnly + vbCritical, UnicodeText("Lo64i")
        Exit Sub
    End If
    
    Dim xDaCo As Boolean
    xDaCo = False
    Dim Ij As Integer
    For Ij = 0 To lstEXT.ListCount - 1
        If UCase(lstEXT.List(Ij)) = UCase(GetFileExt(Dialog1.FileName)) Then xDaCo = True
    Next Ij
    If xDaCo = False Then lstEXT.AddItem UCase(GetFileExt(Dialog1.FileName))


    
    
    lstFile.ItemAdd lstFile.Count + 1, Dialog1.FileName, 0, 0
    lstFile.SubItemSet lstFile.Count, 1, GetMD5(Dialog1.FileName), 0
End If
    
kEtThUcViLoI:
End Sub



Private Sub cmdDel_Click()

Dim i
BaTdAuXoA:
i = 0
For i = 0 To lstFile.Count - 1
    If lstFile.ItemSelected(i) = True Then
        lstFile.ItemRemove (i)
        GoTo BaTdAuXoA
    End If
Next i
End Sub


Private Sub cmdOption_Click()
frmKill.Show
End Sub





Private Sub cmdStart_Click()
If lstFile.Count > 0 Then
    xStopScan = False
    LVREG.Clear
    LVPRO.Clear
    LVFILE.Clear
    fmPRO.Visible = True
    fmREG.Visible = True
    cmdAdd.Enabled = False
    cmdDel.Enabled = False
    tmrReg.Enabled = True
    StartQuetReg
    StartQuetPro
    cmdStart.Enabled = False
    cmdOption.Enabled = False
    pREG.Value = 0
    pPro.Value = 0
    


Else
    UnicodeMsgBox UnicodeText("Ba5n chu7a cho5n ma64u Virus na2o!"), vbOKOnly + vbInformation, UnicodeText("Cho5n ma64u Virus")
End If
End Sub




Private Sub cmdStop_Click()
xStopScan = True
End Sub

















Private Sub Form_Load()
LVREG.ColumnAdd 0, UnicodeText("Te6n kho1a"), 100
LVREG.ColumnAdd 1, UnicodeText("Gia1 Tri5"), 100
LVREG.ColumnAdd 2, UnicodeText("D9u7o72ng da64n"), 100


LVPRO.ColumnAdd 0, UnicodeText("Tie61n Tri2nh"), 100
LVPRO.ColumnAdd 1, UnicodeText("D9u7o72ng da64n"), 100
LVPRO.ColumnAdd 2, UnicodeText("Chi3 So61"), 100

LVFILE.ColumnAdd 0, UnicodeText("Ta65p Tin"), 200

lstFile.ColumnAdd 0, UnicodeText("D9u7o72ng da64n"), 280
lstFile.ColumnAdd 1, UnicodeText("Ma4 Nha65n Da5ng"), 280


lstFile.ImageListSmalls ImageList1.hIml

LVPRO.ImageListSmalls ImageList2.hIml
LVFILE.ImageListSmalls ImageList2.hIml
LVREG.ImageListSmalls ImageList2.hIml

MaxFileSize = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Height < 7400 Then Me.Height = 7400
If Me.Width < 8200 Then Me.Width = 8200

If Me.WindowState <> 1 Then
lbl1.Width = Me.Width
fm.Width = Me.Width - 360
lstFile.Width = Me.Width - cmdAdd.Width - 480
cmdAdd.Left = Me.Width - cmdAdd.Width - 240
cmdDel.Left = Me.Width - cmdDel.Width - 240
cmdStart.Left = Me.Width / 2 - cmdStart.Width / 2

LVREG.Width = fm.Width / 3 - 240
LVPRO.Left = LVREG.Left + LVREG.Width + 120
LVPRO.Width = LVREG.Width
LVFILE.Left = LVPRO.Left + LVPRO.Width + 120
LVFILE.Width = LVPRO.Width
lblReg.Left = LVREG.Left
lblPro.Left = LVPRO.Left
lblFile.Left = LVFILE.Left
fm.Height = (Me.Height - fm.Top) - 700 - Bar1.Height

LVREG.Height = fm.Height - LVREG.Top - 120
LVPRO.Height = fm.Height - LVPRO.Top - 120
LVFILE.Height = fm.Height - LVFILE.Top - 120
cmdOption.Left = Me.Width - cmdOption.Width - 240
fmREG.Top = LVREG.Top
fmREG.Left = LVREG.Left
fmREG.Height = LVREG.Height
fmREG.Width = LVREG.Width

pREG.Left = 120
pREG.Width = fmREG.Width - 240
pREG.Top = fmREG.Height / 2 - pREG.Height / 2
lblx1.Top = pREG.Top - lblx1.Height - 60

fmPRO.Top = LVPRO.Top
fmPRO.Left = LVPRO.Left
fmPRO.Height = LVPRO.Height
fmPRO.Width = LVPRO.Width

pPro.Left = 120
pPro.Width = fmPRO.Width - 240
pPro.Top = fmPRO.Height / 2 - pPro.Height / 2
lblx2.Top = pPro.Top - lblx2.Height - 60


cmdStop.Top = LVFILE.Top + 800
cmdStop.Left = LVFILE.Left + 800

'Label1.Caption = Me.Height & " - " & Me.Width '7400 8200
End If
End Sub





Private Sub StartQuetFile()

cmdStop.Visible = True

Dim Ni As Integer
For Ni = 0 To Drive1.ListCount - 1
    If UCase(Left(Drive1.List(Ni), 1)) <> "A" Then
        Dim Nj As Integer
        For Nj = 0 To lstEXT.ListCount - 1
            SearchFile UCase(Left(Drive1.List(Ni), 2)) & "\", "*." & GetFileExt(lstEXT.List(Nj))
        Next Nj
    End If
Next Ni

Bar1.PanelText(2) = UnicodeText("D9a4 que1t xong!")
Dim Y
For Y = 0 To LVFILE.Count - 1
    LVFILE.ItemChecked(Y) = True
Next Y

'------- Quet Xong ---------
cmdStop.Visible = False
cmdOption.Enabled = True
xCreateReport
End Sub


Private Sub StartQuetReg()
GetSystemKey
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"

Dim Y
For Y = 0 To LVREG.Count - 1
    LVREG.ItemChecked(Y) = True
Next Y
End Sub
Private Sub StartQuetPro()
LVPRO.Clear
    On Error Resume Next
    Dim ColItems
    Dim ObjItem
    Set ColItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")
    For Each ObjItem In ColItems
        If ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System" And ObjItem.ExecutablePath <> "" Then
            'frmMain.lblStatus.Caption = ObjItem.ExecutablePath
            If xCheckVirus(ObjItem.ExecutablePath) = True Then
                Dim Ui
                Ui = LVPRO.Count + 1
                LVPRO.ItemAdd Ui, ObjItem.Caption, 0, 1
                LVPRO.SubItemSet Ui, 1, ObjItem.ExecutablePath, 1
                LVPRO.SubItemSet Ui, 2, ObjItem.ProcessID, 1
            End If
        End If 'ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System"

    Next
    Set ColItems = Nothing
    Set ObjItem = Nothing
    
Dim Y
For Y = 0 To LVPRO.Count - 1
    LVPRO.ItemChecked(Y) = True
Next Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub tmrPro_Timer()
lblx2.Caption = UnicodeText("D9ang Que1t Process...")
pPro.Value = pPro.Value + 1
If pPro.Value > 99 Then
    tmrPro.Enabled = False
    fmPRO.Visible = False
    lblx2.Caption = UnicodeText("Xong!")
    
    'Dong bang de dam bao Virus khong hoat dong trong khi quet
    Dim UyAX As Integer
    For UyAX = 0 To frmMain.LVPRO.Count - 1
    DoEvents
        'frmMain.LVPRO.SubItemText(Uy, 2)
        SuspendResumeProcess frmMain.LVPRO.SubItemText(UyAX, 2), True
    Next UyAX
    
    
    StartQuetFile
End If
End Sub

Private Sub tmrReg_Timer()
lblx1.Caption = UnicodeText("D9ang Que1t Registry...")
pREG.Value = pREG.Value + 1
If pREG.Value > 99 Then
    tmrReg.Enabled = False
    fmREG.Visible = False
    tmrPro.Enabled = True
    lblx1.Caption = UnicodeText("Xong!")
    
End If
End Sub
