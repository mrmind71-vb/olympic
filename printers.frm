VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form printersfrm 
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9885
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
   ScaleHeight     =   1680
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand cmdExit 
      Height          =   510
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1125
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   900
      _Version        =   196610
      ForeColor       =   0
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "printers.frx":0000
      Alignment       =   4
      ButtonStyle     =   1
      BevelWidth      =   10
      ShapeSize       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   1050
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   9600
      _cx             =   16933
      _cy             =   1852
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
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
      Editable        =   2
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
End
Attribute VB_Name = "printersfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cFileSave As String

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
cFileSave = tempPath & turn(tempPath, "\") & "printers.txt"
myload
Fixgrd
End Sub
Private Sub Fixgrd()
With grid1
.ColWidth(0) = 1700
.ColWidth(1) = 7500
.TextMatrix(0, 0) = "≈”„ «·ÿ«»⁄…"
.TextMatrix(0, 1) = "‰Ê⁄ «·ÿ«»⁄…"
.TextMatrix(1, 0) = "PDF PRINTER"
.TextMatrix(2, 0) = "ÿ«»⁄… «·»«—ﬂÊœ"
.ColComboList(1) = "..."
.ColAlignment(0) = flexAlignRightCenter
.ColAlignment(1) = flexAlignRightCenter
.RowHidden(2) = True
End With
End Sub
Private Sub GRID1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Set sys_printersfrm.myForm = Me
sys_printersfrm.Show 1
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row > 0 And grid1.Col = 1 And grid1.TextMatrix(grid1.Row, 1) <> "" Then
    If MsgBox("Õ–› «·ÿ«»⁄…", vbOKCancel) = vbOK Then
        addSetting "printer" & grid1.Row, "", cFileSave
        If RetSetting("printer" & grid1.Row, cFileSave) = "" Then
            grid1.TextMatrix(grid1.Row, 1) = ""
            Inform " „  Õ–› «·ÿ«»⁄… »‰Ã«Õ"
        Else
            MsgBox "·„ Ì „ Õ–› «·ÿ«»⁄…"
        End If
    End If
End If
End Sub
Sub myProc()
Dim cPrinter As String
cPrinter = sys_printersfrm.grid1.TextMatrix(sys_printersfrm.grid1.Row, 0)
Unload sys_printersfrm
If cPrinter = "" Then Exit Sub
addSetting "printer" & grid1.Row, cPrinter, cFileSave
If RetSetting("printer" & grid1.Row, cFileSave) = cPrinter Then
    grid1.TextMatrix(grid1.Row, 1) = cPrinter
    Inform " „  «÷«›… «·ÿ«»⁄… »‰Ã«Õ"
Else
    MsgBox "·„ Ì „ «÷«›… «·ÿ«»⁄…"
End If
Exit Sub
End Sub
Private Sub myload()
For i = 1 To grid1.rows - 1
    grid1.TextMatrix(i, 1) = RetSetting("printer" & i, cFileSave)
Next
End Sub
