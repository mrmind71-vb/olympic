VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form ShowTaxInstallfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   " ›’Ì·Ì „ √Œ—«  «ﬁ”«ÿ ﬁÌ„… „÷«›…"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14325
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   14325
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   4635
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   9600
      Begin VB.Label xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   945
         Width           =   3525
      End
      Begin VB.Label xcode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   3525
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈”„ «·⁄„Ì·"
         Height          =   240
         Left            =   8235
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Left            =   8235
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1035
         Width           =   510
      End
      Begin VB.Label xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   3525
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
         Height          =   240
         Left            =   8235
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   1125
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5955
      Left            =   135
      TabIndex        =   0
      Top             =   1485
      Width           =   14100
      _cx             =   24871
      _cy             =   10504
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   7425
      Width           =   2535
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
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
         Picture         =   "Show_Tax_Install.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1260
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   900
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
         Picture         =   "Show_Tax_Install.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Show_Tax_Install.frx":4699
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2925
      Top             =   5310
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   582
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
End
Attribute VB_Name = "ShowTaxInstallfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pCode As String
Public pDate As String
Dim aHeader(1)
Dim con As New ADODB.Connection
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdPrint_Click()
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", grid1.rows - 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 2)
aRow(0) = AddFlag(aRow(0), "text", "«·≈Ã„«·Ì")

'grid1.ColHidden(4) = True
'grid1.ColHidden(5) = True
'grid1.ColHidden(6) = True

PrintGrdNew.doprint grid1, 1.1, -1, "", retHeader(aHeader, 0, 1, Space(6)), retHeader(aHeader, 1, 1, Space(6)), , False, False, 11, , aRow
PrintGrdNew.Show 1

'grid1.ColHidden(4) = False
'grid1.ColHidden(5) = False
'grid1.ColHidden(6) = False

End Sub
Private Sub Form_Load()
openCon con
Set grid1.DataSource = DATA1
xCode.Caption = pCode
xDesca.Caption = Member_Load(pCode, "desca") & ""
xDate.Caption = myFormat_p(pDate)
myload
End Sub
Sub Fixgrd()
With grid1
.RowHeight(0) = 800
.WordWrap = True
.FormatString = "—ﬁ„ «·«Ì’«·|" & "«· «—ÌŒ|" & "«·ﬁ”ÿ|" & "‰”»… «·÷—Ì»…|" & "≈Ã„«·Ì «·÷—Ì»…|" & "‘ÂÊ— «·€—«„…|" & "≈Ã„«·Ì «·€—«„…|" & "≈Ã„«·Ì «·„” Õﬁ"
.ColWidth(0) = 1000
.ColWidth(1) = 1300
.ColWidth(2) = 1300
.ColWidth(3) = 1200
.ColWidth(4) = 1200
.ColWidth(5) = 1000
.ColWidth(6) = 1200
.ColWidth(7) = 1200
.ColFormat(3) = "%##"

Dim nwidth As Long
For i = 0 To .Cols - 1
    nwidth = nwidth + IIf(.ColHidden(i), 0, .ColWidth(i))
    .ColAlignment(i) = flexAlignRightCenter
    
Next

If .rows > 1 Then
    .SubtotalPosition = flexSTBelow
    .ExplorerBar = flexExSortShow
    .Subtotal flexSTSum, -1, 2, "#0.00", &HC0FFC0, vbBlack, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0.00", &HC0FFC0, vbBlack, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0.00", &HC0FFC0, vbBlack, True, "  "
    .Subtotal flexSTSum, -1, 7, "#0.00", &HC0FFC0, vbBlack, True, "  "
    .TextMatrix(.rows - 1, 0) = "«·≈Ã„«·Ï"
    FixTotals grid1, .rows - 1, Array(2, 4, 6, 7)
End If
End With
End Sub
Private Sub myload()
Dim cString As String
Dim myCmd As New ADODB.Recordset
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "code", pCode)
aPrm = AddFlag(aPrm, "date", myFormat_p(pDate))

Set DATA1.Recordset = myCmdProc("dbo.sp_tax_install", con, aPrm)
Fixgrd
grid1.Select grid1.rows - 1, 1
grid1.ShowCell grid1.rows - 1, 1
aHeader(0) = ArbString("„ √Œ—«  ÷—Ì»… «·ﬁÌ„… «·„÷«›… ··√ﬁ”«ÿ ⁄·Ì «·⁄÷Ê : " & xDesca.Caption & "  (" & pCode & ")")
aHeader(1) = "» «—ÌŒ : " & myFormat_p(xDate.Caption)
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set ShowTaxLatefrm = Nothing
End Sub

