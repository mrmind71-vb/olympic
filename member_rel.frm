VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form member_relfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÈíÇäÇÊ ÊæÇÈ ÇáÚÖæ"
   ClientHeight    =   4545
   ClientLeft      =   690
   ClientTop       =   1395
   ClientWidth     =   10860
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   10860
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   7845
      Begin VB.Label xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   270
         Width           =   2040
      End
      Begin VB.Label xDate_Begin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2295
         Width           =   5775
      End
      Begin VB.Label xAge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1890
         Width           =   5775
      End
      Begin VB.Label xDate_Birth 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1485
         Width           =   5775
      End
      Begin VB.Label xRel_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1080
         Width           =   5775
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   675
         Width           =   5775
      End
      Begin VB.Label xMember 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   270
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇáÓä"
         Height          =   330
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1935
         Width           =   1590
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÊÇÑíÎ ÇáãíáÇÏ"
         Height          =   330
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1575
         Width           =   1590
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÏÑÌÉ ÇáÞÑÇÈÉ"
         Height          =   330
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1170
         Width           =   960
      End
      Begin VB.Label xcode_zero 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   270
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÊÇÑíÎ ÈÏÇíÉ ÇáÚÖæíÉ"
         Height          =   330
         Left            =   5985
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2385
         Width           =   1590
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÑÞã ÇáÚÖæíÉ"
         Height          =   285
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   315
         Width           =   945
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇáÇÓã"
         Height          =   330
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   645
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   2670
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   510
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_rel.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   1350
         TabIndex        =   0
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_rel.frx":2323
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "member_rel.frx":46EE
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1800
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   -2295
      Top             =   405
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   741
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   375
      Left            =   7110
      Top             =   7515
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ImgXCtrl6.ImgXCtrl imgx1 
      DragIcon        =   "member_rel.frx":6797
      DragMode        =   1  'Automatic
      Height          =   2085
      Left            =   12330
      TabIndex        =   6
      Tag             =   "-1"
      Top             =   495
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   3678
      BorderStyle     =   1
      AutoZoom        =   -1  'True
      LicenseUserName =   "mrmind71"
      LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   4170
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   45
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   476
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   3465
         TabIndex        =   16
         Top             =   45
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   6975
         TabIndex        =   17
         Top             =   45
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   90
      Width           =   2625
      Begin VB.Image xMemberPhoto 
         Appearance      =   0  'Flat
         Height          =   3180
         Left            =   90
         Stretch         =   -1  'True
         Top             =   180
         Width           =   2430
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3510
      Width           =   3480
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2610
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_rel.frx":6BD9
         Caption         =   "Ãæá"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "member_rel.frx":8D80
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1710
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   180
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   741
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_rel.frx":ADC7
         Caption         =   "ÓÇÈÞ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "member_rel.frx":CEB2
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   855
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_rel.frx":EEAC
         Caption         =   "áÇÍÞ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "member_rel.frx":10FBD
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_rel.frx":12FB7
         Caption         =   "ÃÎíÑ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "member_rel.frx":151DB
      End
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   3690
      Top             =   585
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "member_relfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection, bCheck As Boolean
Dim fs As New FileSystemObject
Dim formMode As Byte
Dim oSearch As New Search_empty
Dim CardTable As ADODB.Recordset
Public sMember As String, sCode As String, ntype As Long
Dim pFile As String
Dim cFilter As String, cFilterLookup As String
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
bEditRecord = bEdit
cmdInform.Enabled = (nMode = LoadMode)

aRecords = retRecords(xCode.Caption)
nRecord = Val(retFlag(aRecords, "record") & "")
nRecords = Val(retFlag(aRecords, "records") & "")

If nMode = LoadMode Then
    panel1(0).Caption = ArbString("ÓÌá " & nRecord & " ãä " & nRecords)
Else
    panel1(0).Caption = ArbString("áÇ ÊæÌÏ ÓÌáÇÊ")
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2
End Sub
Sub mydefine()
xMember.Caption = sMember
xCode.Caption = ""
xDesca.Caption = ""
xRel_desca.Caption = ""
xDate_Begin.Caption = ""
xDate_Birth.Caption = ""
xAge.Caption = ""
Handlecontrols DefineMode
End Sub
Sub myProc()
If ActiveControl.Name = cmdInform.Name Then
    xCode.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
End If
myUndo
End Sub
Private Sub myload()
xCode.Caption = CardTable!code & ""
xMember.Caption = CardTable!MEMBER & ""
xDesca.Caption = CardTable!desca
xRel_desca.Caption = CardTable!REL_DESCA & ""
xDate_Begin.Caption = myFormat_p(CardTable!date_begin)
xDate_Birth.Caption = myFormat_p(CardTable!DATE_BIRTH)
If IsDate(xDate_Birth.Caption) Then
    xAge.Caption = AgeString(myFormat(xDate_Birth.Caption), myFormat(Date))
Else
    xAge.Caption = ""
End If
Handlecontrols LoadMode
xMemberPhoto.Picture = LoadPicture("")
If ntype = 1 Then
    If validPhoto(RetPhoto(xMember.Caption & "-" & xCode.Caption)) Then xMemberPhoto.Picture = LoadPicture(RetPhoto(xMember.Caption & "-" & xCode.Caption))
Else
    If validPhoto(RetPhoto_I(xMember.Caption & "-" & xCode.Caption)) Then xMemberPhoto.Picture = LoadPicture(RetPhoto_I(xMember.Caption & "-" & xCode.Caption))
End If
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdInform_Click()
If ntype = 1 Then
    relLookupAll2 Me, oSearch, "FILE1_11.MEMBER = " & sMember
Else
    relLookupAll_i2 Me, oSearch, "FILE2_11.MEMBER = " & sMember
End If
End Sub
Private Sub CmdNext_Click()
openCardTable xCode.Caption, ">"
If CardTable.EOF Then openCardTable xCode.Caption
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xCode.Caption, "<"
If CardTable.EOF Then openCardTable xCode.Caption
myload
End Sub
Private Sub CmdFirst_Click()
openCardTable , ">"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Private Sub CmdLast_Click()
openCardTable , "<"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()

openCon con
bEdit = True

pFile = IIf(ntype = 1, "FILE1_11", "FILE2_11")

If sCode <> "" Then xCode.Caption = sCode
myUndo
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
If Not ValidNum(xCode.Caption) Then
     If xCode.Tag = LoadMode Then
        mydefine
    Else
        xCode.Caption = ""
    End If
Else
    If (Not (CardTable.EOF)) And xCode.Tag = LoadMode Then
        If CardTable!code = xCode.Caption Then
            Exit Sub
        End If
    End If
    
    openCardTable xCode.Caption
    If Not CardTable.EOF Then
        myload
    ElseIf xCode.Tag = LoadMode Then
        mydefine
    Else
        'xcode.caption = ""
    End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Set member_relfrm = Nothing
End Sub
Private Function openCardTable(Optional pCode As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 " & pFile & ".*,RELATION_CODES.DESCA AS REL_DESCA FROM " & pFile & " LEFT JOIN RELATION_CODES ON " & pFile & ".RELATION = RELATION_CODES.CODE"
If pSign = "=" Then
    If pCode <> "" Then cWhere = pFile & ".CODE  " & pSign & addvalue(pCode)
Else
    If pCode <> "" Then cWhere = pFile & ".CODE " & pSign & addvalue(pCode)
End If

cFilter = ""
If sMember <> "" Then cFilter = "MEMBER = " & addvalue(sMember)
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by " & pFile & ".CODE desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by " & pFile & ".CODE ASC"
End If

CardTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
End Function
Private Sub myUndo()
On Error GoTo myerror
Dim cString As String, cWhere As String
If ValidNum(xCode.Caption) Then
    openCardTable xCode.Caption
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then mydefine Else myload
On Error Resume Next
'xdesca.SetFocus
Err.Clear
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function retRecords(pCode) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If Trim(pCode) <> "" Then
    cString = "SELECT SUM(1) AS records,SUM(CASE WHEN CODE <= " & addvalue(pCode) & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT SUM(1) AS records"
End If
cString = cString & " FROM " & pFile & turn(cFilter, " WHERE ") & cFilter
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If Trim(pCode) <> "" Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
End Function
