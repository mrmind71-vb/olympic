VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form ShowItem 
   Caption         =   " ⁄œÌ· »Ì«‰«  «·«’‰«›"
   ClientHeight    =   6750
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9675
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "file_2"
      Top             =   525
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   90
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   45
      Width           =   2115
   End
   Begin VSFlex7LCtl.VSFlexGrid VsItem 
      Height          =   9555
      Left            =   90
      TabIndex        =   0
      Top             =   495
      Width           =   18870
      _cx             =   33285
      _cy             =   16854
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
      AutoResize      =   -1  'True
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
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "ShowItem.frx":0000
      DataSource      =   "Data2"
      Height          =   315
      Left            =   14085
      TabIndex        =   2
      Top             =   135
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ã„Ê⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   18135
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   225
      Width           =   750
   End
End
Attribute VB_Name = "ShowItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CardTable As Recordset
Dim MyItemTable As Recordset
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Dim cStr1 As String
Set CardTable = mydb.OpenRecordset("SELECT * FROM FILE1_10")
Data2.DatabaseName = MdbPath
Data2.RecordSource = "FILE1_50"
Data2.Refresh
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"
With VsItem
    .WordWrap = True
    .ExplorerBar = flexExSortShow
    .Rows = 1
    .Cols = 10
    .RowHeight(0) = 500
    .FixedCols = 1
    .FixedRows = 1
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·’‰›"
    .TextMatrix(0, 2) = "»Ì«‰ »«—œﬂÊœ"
    .TextMatrix(0, 3) = "⁄œœ «·—›"
    .TextMatrix(0, 4) = "≈⁄«œ… «·ÿ·»"
    .TextMatrix(0, 5) = "Ã„·…"
    .TextMatrix(0, 6) = "Ã„·… „ Ê”ÿ"
    .TextMatrix(0, 7) = "„Õ·« "
    .TextMatrix(0, 8) = "„” Â·ﬂ"
    .TextMatrix(0, 9) = "⁄»Ê… ﬂ— Ê‰…"
    
    
    .ColWidth(0) = 800
    .ColWidth(1) = 6000
    .ColWidth(2) = 3000
    .ColWidth(3) = 808
    .ColWidth(4) = 800
    .ColWidth(5) = 800
    .ColWidth(6) = 800
    .ColWidth(7) = 800
    .ColWidth(8) = 800
    .ColWidth(9) = 800
    '.ColHidden(2) = True
    .ColHidden(6) = True
    .ColHidden(7) = True
    .ColHidden(9) = True

End With
End Sub
Private Sub VsItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo myerror
    With VsItem
        MyItemTable.FindFirst " ITEM = " & MyParn(.TextMatrix(Row, 0))
        If Not MyItemTable.NoMatch Then
            MyItemTable.Edit
            Select Case Col
                Case 1
                    If .EditText <> "" Then MyItemTable.desca = .EditText
                Case 2
                    If .EditText <> "" Then MyItemTable.desca2 = .EditText
                    
                Case 3
                    If .EditText <> "" Then MyItemTable.R1 = Val(.EditText)
                Case 4
                    If .EditText <> "" Then MyItemTable.R2 = Val(.EditText)
                Case 5
                    If .EditText <> "" Then MyItemTable.COST1 = Val(.EditText)
                Case 6
                    If .EditText <> "" Then MyItemTable.COST2 = Val(.EditText)
                Case 7
                    If .EditText <> "" Then MyItemTable.COST4 = Val(.EditText)
                Case 8
                    If .EditText <> "" Then MyItemTable.price = Val(.EditText)
                Case 9
                    If .EditText <> "" Then MyItemTable.Pack = Val(.EditText)
            End Select
            MyItemTable.Update
        End If
    End With
Exit Sub
myerror:
If Err.Number = 3022 Then
    MsgBox "ﬂÊœ «·ﬂ— Ê‰… „”Ã· „‰ ﬁ»·"
    CardTable.FindFirst " factcode = " & MyParn(VsItem.EditText)
    If Not CardTable.NoMatch Then MsgBox "„”Ã· „‰ ﬁ»· ··’‰› " & CardTable.desca
    Cancel = True
Else
    MsgBox Err.Description
End If
MsgBox "·„ Ì „ «·Õ›Ÿ"
Err.Clear
End Sub
Private Sub xGroup_Change()
    If xGroup.BoundText = "" Then Exit Sub
    cStr1 = " SELECT * FROM FILE1_10  WHERE FILE1_10.GROUP = " & MyParn(xGroup.BoundText) & " ORDER BY ITEM "
    Set MyItemTable = mydb.OpenRecordset(cStr1)
    With VsItem
    .Rows = 1
    If MyItemTable.RecordCount = 0 Then Exit Sub
    MyItemTable.MoveFirst
    Do While Not MyItemTable.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = TurnValue(MyItemTable.Item, Null, "")
        .TextMatrix(.Rows - 1, 1) = TurnValue(MyItemTable.desca, Null, "")
        .TextMatrix(.Rows - 1, 2) = TurnValue(MyItemTable!desca2, Null, "")
        .TextMatrix(.Rows - 1, 3) = Format(MyItemTable.R1, "#0")
        .TextMatrix(.Rows - 1, 4) = Format(MyItemTable.R2, "#0")
        .TextMatrix(.Rows - 1, 5) = Format(MyItemTable.COST1, "##0.00")
        .TextMatrix(.Rows - 1, 6) = Format(MyItemTable.COST2, "##0.00")
        .TextMatrix(.Rows - 1, 7) = Format(MyItemTable.COST4, "##0.00")
        .TextMatrix(.Rows - 1, 8) = Format(MyItemTable.price, "##0.00")
        .TextMatrix(.Rows - 1, 9) = Format(MyItemTable.Pack, "##0")
        MyItemTable.MoveNext
    Loop
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub

