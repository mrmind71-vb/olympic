VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Vs_Stock 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„” ‰œ Ã—œ"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdLast 
      BackColor       =   &H00E0E0E0&
      Caption         =   "√ŒÌ—"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1680
      Width           =   765
   End
   Begin VB.CommandButton CmdFirst 
      BackColor       =   &H00E0E0E0&
      Caption         =   "√Ê·"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2670
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1680
      Width           =   765
   End
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "·«Õﬁ"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1815
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1680
      Width           =   765
   End
   Begin VB.CommandButton CmdPrevious 
      BackColor       =   &H00E0E0E0&
      Caption         =   "”«»ﬁ"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   975
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1680
      Width           =   765
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   75
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "≈÷«›… ﬂ· «·√’‰«›"
      Height          =   390
      Left            =   7560
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ÿ»«⁄… «·Ã—œ"
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
      Left            =   10320
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1290
   End
   Begin VB.CommandButton Cmd_AddAll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "⁄—÷ —’Ìœ «·√’‰«›"
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
      Left            =   7080
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   75
      Width           =   1770
   End
   Begin VB.CommandButton cmdFix 
      BackColor       =   &H00E0E0E0&
      Caption         =   "«⁄«œ… ÷»ÿ «·Ã—œ"
      Height          =   390
      Left            =   240
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1500
   End
   Begin VB.CommandButton cmdunPost 
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·€«¡  —ÕÌ· «·„” ‰œ"
      Height          =   390
      Left            =   1875
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1500
   End
   Begin VB.CommandButton cmdPost 
      BackColor       =   &H00E0E0E0&
      Caption         =   " —ÕÌ· «·„” ‰œ"
      Height          =   390
      Left            =   3525
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1500
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   600
      Width           =   1140
   End
   Begin VB.CommandButton cmd_item 
      BackColor       =   &H00E0E0E0&
      Caption         =   "»Ì«‰«  «·√’‰«›"
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
      Left            =   9120
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   75
      Width           =   1290
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00E0E0E0&
      Caption         =   " —«Ã⁄"
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
      Left            =   1500
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1065
      Width           =   1365
   End
   Begin VB.CommandButton CmdDelInv 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Õ–› «·„” ‰œ"
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
      Left            =   5445
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   75
      Width           =   1290
   End
   Begin VB.CommandButton CmdAddItem 
      BackColor       =   &H00E0E0E0&
      Caption         =   "«÷«›… √’‰«›"
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
      Left            =   10545
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   75
      Width           =   1290
   End
   Begin VB.CommandButton cmdNewinv 
      BackColor       =   &H00E0E0E0&
      Caption         =   "„” ‰œ ÃœÌœ"
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
      Left            =   1500
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   435
      Width           =   1365
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Œ—ÊÃ"
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
      Left            =   75
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   435
      Width           =   1365
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Õ›Ÿ «·„” ‰œ"
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
      Left            =   75
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1065
      Width           =   1365
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   750
      Width           =   1290
   End
   Begin VB.TextBox xDoc_No 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9195
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   750
      Width           =   1290
   End
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   5460
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   11700
      _cx             =   20637
      _cy             =   9631
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      BackColor       =   14737632
      ForeColor       =   -2147483640
      BackColorFixed  =   12615808
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   14737632
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
      GridLineWidth   =   2
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Vs_Stock.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      OutlineCol      =   1
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
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   1
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   4
   End
   Begin MSDBCtls.DBCombo xStore 
      Bindings        =   "Vs_Stock.frx":0088
      DataSource      =   "Data1"
      Height          =   315
      Left            =   7200
      TabIndex        =   13
      Top             =   1200
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label xMItem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„Œ“‰"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10680
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1230
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ „” ‰œ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10680
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   705
      Width           =   840
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«· «—ÌŒ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4470
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   825
      Width           =   555
   End
End
Attribute VB_Name = "Vs_Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DocTable As Recordset, itemTable As Recordset
Dim movetable As Recordset
Dim ItemDateBal As Recordset
Dim ItemNotInv As Recordset
Dim formMode
Dim StockBalTable As Recordset
Dim C_InvTable As Recordset
Const NewInvMode = 4, applyMode = 5
Sub DocValid()
If xDoc_No.Text = "" Then Exit Sub
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.FindFirst " doc_no = " & MyParn(xDoc_No)
If DocTable.NoMatch Then
    Exit Sub
Else
    ApplyProc
End If
End Sub
Sub fillgrd()
ItemInv.Rows = 1
i = 1
With ItemInv
.FixedRows = 1
.ExplorerBar = flexExSortShow
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then Exit Sub
Do While True
   .AddItem ""
    .TextMatrix(i, 0) = DocTable.Item
    .TextMatrix(i, 1) = RetFind(itemTable, "item", "DescA", DocTable!Item)
    .TextMatrix(i, 2) = TurnValue(Format(DocTable.ComputerBal, "##0.00"), Null, "")
    .TextMatrix(i, 3) = TurnValue(Format(DocTable.RealBal, "##0.00"), Null, "")
    .TextMatrix(i, 4) = TurnValue(Format(DocTable.Differ, "##0.00"), Null, "")
    If bopt2 Then
    .TextMatrix(i, 5) = TurnValue(Format(DocTable.COST, "##0.00"), Null, "")
    .TextMatrix(i, 6) = TurnValue(Format(DocTable.total, "##0.00"), Null, "")
    End If
    
    DocTable.MoveNext
    If DocTable.EOF Then Exit Sub
    If DocTable.doc_no <> xDoc_No.Text Then Exit Sub
    i = i + 1
Loop
End With
End Sub
Sub ItemsLookup()
    ActiveControl.Text = ""
    Dim Generalarray(4)
    Dim GrdArray(2)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Item as «·’‰›,DescA as [«”„ «·’‰›] From file1_10 "
    Generalarray(3) = " Where DescA Like('*cFilter*')  OR  ITEM Like('*cFilter*')    "
    Generalarray(4) = "Order by Item"
    GrdArray(1) = 1000
    GrdArray(2) = 4500
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Function MyReplace() As Boolean
MyReplace = True
' „—«Ã⁄… «·„” ‰œ
With ItemInv
For i = 1 To .Rows - 1
    itemTable.FindFirst " ITEM = " & MyParn(.TextMatrix(i, 0))
    If itemTable.NoMatch And .TextMatrix(i, 0) <> "" Then
        .Select i, 0, i, 4
        cMess = "·‰ Ì „  ”ÃÌ· «·’‰› " & .TextMatrix(i, 0) & " ﬂÊœ «·’‰› €Ì— „”Ã· "
        MsgBox cMess
        MyReplace = False
    End If
Next i
End With

If MyReplace Then
    ' Õ–› «·„” ‰œ ﬁ»· «· ⁄œÌ·
    cString = " DELETE  *  FROM FILE0_10  WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    
    '  ”ÃÌ· «·„” ‰œ
    With ItemInv
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 0) <> "" Then
        itemTable.FindFirst " ITEM = " & MyParn(.TextMatrix(i, 0))
        DocTable.AddNew
        DocTable.doc_no = xDoc_No.Text
        DocTable.[Date] = xDate.Text
        DocTable.Store = xStore.BoundText
        DocTable.Item = .TextMatrix(i, 0)
        DocTable.ComputerBal = Val(.TextMatrix(i, 2))
        DocTable.RealBal = Val(.TextMatrix(i, 3))
        DocTable.Differ = Val(.TextMatrix(i, 4))
        DocTable.COST = Val(.TextMatrix(i, 5))
        DocTable.total = Val(.TextMatrix(i, 5)) * Val(.TextMatrix(i, 4))
        DocTable.Update
        End If
    Next i
    End With
    Set DocTable = mydb.OpenRecordset("select * from FILE0_10 order by doc_no ", dbOpenDynaset)
End If
End Function
Sub ApplyProc()
If Not DocTable.EOF Then
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then
    xDoc_No.Enabled = True
    xStore.BoundText = ""
    xDate.Text = Format(Date, "DD-MM-YYYY")
Else
    xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
    xStore.BoundText = TurnValue(DocTable.Store, Null, "")
    If DocTable.CLOSED Then
        cmdFix.Enabled = False
        cmdPost.Enabled = False
        cmdunPost.Enabled = True
    Else
        cmdFix.Enabled = True
        cmdPost.Enabled = True
        cmdunPost.Enabled = False
    End If
    fillgrd
    dispProc
    xDoc_No.Enabled = False
End If
End If
End Sub
Sub myProc()
If ActiveControl.Name = ItemInv.Name Then
    ItemInv.EditText = GrdText(Search.grid1, 0)
    ItemInv.TextMatrix(ItemInv.Row, 0) = GrdText(Search.grid1, 0)
    ItemInv.TextMatrix(ItemInv.Row, 1) = GrdText(Search.grid1, 1)
Else
    ActiveControl.Text = GrdText(Search.grid1, 0)
End If
Unload Search
End Sub
Function MYVALID()
MYVALID = True
If xDoc_No.Text = "" Then
    MsgBox " ”ÃÌ· —ﬁ„ «·„” ‰œ"
    MYVALID = False
End If
If xDate.Text = "" Or Not IsDate(xDate.Text) Then
    MsgBox " ”ÃÌ· «· «—ÌŒ"
    MYVALID = False
End If
If xStore.BoundText = "" Then
    MsgBox " ”ÃÌ· «·„Œ“‰ "
    MYVALID = False
End If
End Function
Private Sub Cmd_AddAll_Click()
If Not MsgBox("≈÷«›… √’‰«› ·Â« —’Ìœ Ê ·„  ”Ã· ›Ï «·Ã—œ Ê ÌﬂÊ‰ —’Ìœ Ã—œÂ« ’›—  : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    Exit Sub
End If
Dim nStock As Double

DocTable.FindFirst "doc_no = " & MyParn(xDoc_No)

'If Not DocTable.NoMatch Then Exit Sub

cString = " SELECT Sum(FILE1_11.[IN]) AS SumIN, FILE1_10.COST, Sum(FILE1_11.OUT) AS SumOUT, FILE1_10.ITEM  FROM FILE1_11 RIGHT JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM  " & _
          " WHERE STORE = " & MyParn(xStore.BoundText) & _
          " and Date < DateValue(" & MyParn(xDate.Text) & ")" & _
          " GROUP BY FILE1_10.ITEM , FILE1_10.COST "

mydb.Execute "drop table t1 "
mydb.CreateQueryDef "t1", cString

Set ItemNotInv = mydb.CreateSnapshot("t1")

With ItemNotInv
    If .RecordCount > 0 Then
        .MoveFirst
        Do While True
            DocTable.FindFirst "doc_no = " & MyParn(xDoc_No) & " and item = " & MyParn(.Item)
            If DocTable.NoMatch Then
                nBal = TurnValue(.SUMIN, Null, 0) - TurnValue(.SUMOUT, Null, 0)
                If nBal <> 0 Then
                    DocTable.AddNew
                    DocTable.doc_no = xDoc_No.Text
                    DocTable.[Date] = xDate.Text
                    DocTable.Store = xStore.BoundText
                    DocTable.Item = .Item
                    DocTable.ComputerBal = nBal
                    DocTable.RealBal = 0
                    DocTable.Differ = 0 - nBal
                    DocTable.COST = .COST
                    DocTable.total = .COST * nBal * -1
                    DocTable.Update
                End If
            End If
            .MoveNext
            If .EOF Then Exit Do
        Loop
    End If
End With
fillgrd
End Sub
Private Sub CMD_ITEM_Click()
items.Show 1
itemTable.Requery
End Sub
Private Sub Cmd_Print_Click()
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
With ItemInv
For i = 1 To .Rows - 1
    TargetTable.AddNew
    TargetTable.str1 = xDoc_No.Text
    TargetTable.str2 = xStore.Text
    TargetTable.Date1 = xDate.Text
    
    TargetTable.str3 = .TextMatrix(i, 0)
    TargetTable.str4 = .TextMatrix(i, 1)
    TargetTable.VAL1 = Val(.TextMatrix(i, 2))
    TargetTable.VAL2 = Val(.TextMatrix(i, 3))
    TargetTable.VAL3 = Val(.TextMatrix(i, 4))
    If bopt2 Then
    TargetTable.VAL4 = Val(.TextMatrix(i, 5))
    If Val(.TextMatrix(i, 6)) > 0 Then
        TargetTable.VAL6 = Val(.TextMatrix(i, 6))
    Else
        TargetTable.VAL5 = Val(.TextMatrix(i, 6))
    End If
    End If
    TargetTable.str9 = firsttitle & "     " & Secondtitle
    TargetTable.Update
Next i
End With
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\INVENT.RPT"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
Private Sub cmdAdditem_Click()
    With ItemInv
    .Rows = .Rows + 1
    For i = ItemInv.Rows - 1 To 2 Step -1
        .TextMatrix(i, 0) = .TextMatrix(i - 1, 0)
        .TextMatrix(i, 1) = .TextMatrix(i - 1, 1)
        .TextMatrix(i, 2) = .TextMatrix(i - 1, 2)
        .TextMatrix(i, 3) = .TextMatrix(i - 1, 3)
        .TextMatrix(i, 4) = .TextMatrix(i - 1, 4)
        .TextMatrix(i, 5) = .TextMatrix(i - 1, 5)
        .TextMatrix(i, 6) = .TextMatrix(i - 1, 6)
    Next i
    .TextMatrix(1, 0) = ""
    .TextMatrix(1, 1) = ""
    .TextMatrix(1, 2) = ""
    .TextMatrix(1, 3) = ""
    .TextMatrix(1, 4) = ""
    .TextMatrix(1, 5) = ""
    .TextMatrix(1, 6) = ""
    End With
End Sub
Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        myDelete
        xDoc_No.Text = ""
        fillgrd
        xDoc_No.Enabled = True
        ItemInv.Enabled = False
        Set DocTable = mydb.OpenRecordset("select * from FILE0_10 order by doc_no ", dbOpenDynaset)
    End If
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdFix_Click()
Dim cString As String
cString = "SELECT FILE1_11.ITEM, Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOut " & _
          " FROM FILE1_11 " & _
          " WHERE Date < DateValue(" & MyParn(xDate.Text) & ")" & _
          " AND STORE = " & MyParn(xStore.BoundText) & _
          " GROUP BY ITEM "
Set movetable = mydb.OpenRecordset(cString, dbOpenDynaset)
    
    Dim nBal As Double
    With ItemInv
    For i = 1 To .Rows - 1
        nBal = BalItemDateStore(.TextMatrix(i, 0))
        xMItem.Caption = Val(.TextMatrix(i, 0))
        .TextMatrix(i, 2) = nBal
        .TextMatrix(i, 4) = Format(Val(.TextMatrix(i, 3)) - nBal, "##0.00")
        .TextMatrix(i, 6) = Format(Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 5)), "##0.00")
    Next i
    End With
If Not MyReplace Then Exit Sub
End Sub
Private Sub CmdNewInv_Click()
'If Not MyReplace Then Exit Sub
ItemInv.Rows = 1
ItemInv.AddItem ""
xStore.BoundText = ""
xDate.Text = Date
xDoc_No.Enabled = True
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.doc_no)
Else
    xDoc_No.Text = "000001"
End If
xDoc_No.SetFocus
End Sub
Private Sub cmdPost_Click()
    cString = "insert into File1_11(type,item,[date],store,desca,Doc_Id,[In])" & _
        " Select 'z',item,[date],store,' Ã—œ ›Ï ' & Format(Date,'dd-mm-yyyy'),Doc_NO,differ From File0_10 where DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
  
    cString = " UPDATE FILE0_10 SET FILE0_10.closed = " & True & " WHERE FILE0_10.DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString

    Set DocTable = mydb.OpenRecordset("select * from FILE0_10 order by doc_no ", dbOpenDynaset)

    cmdFix.Enabled = False
    cmdPost.Enabled = False
    cmdunPost.Enabled = True

End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
fillgrd
End Sub
Private Sub CmdUndo_Click()
    DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
    If Not DocTable.NoMatch Then
        xStore.BoundText = DocTable.Store
        xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
        fillgrd
    End If
End Sub
Private Sub cmdunPost_Click()
    cString = " DELETE  * FROM FILE1_11  WHERE [type] = 'Z' AND DOC_id = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    
    cString = " UPDATE FILE0_10 SET FILE0_10.closed = " & False & " WHERE FILE0_10.DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString

    Set DocTable = mydb.OpenRecordset("select * from FILE0_10 order by doc_no ", dbOpenDynaset)

    cmdFix.Enabled = True
    cmdPost.Enabled = True
    cmdunPost.Enabled = False
End Sub

Private Sub Command1_Click()
With DocTable
itemTable.MoveFirst
Do While True
    .FindFirst " ITEM = " & MyParn(itemTable.Item) & " AND  DOC_NO = " & MyParn(xDoc_No.Text)
    If .NoMatch Then
        .AddNew
        .doc_no = xDoc_No.Text
        .[Date] = xDate.Text
        .Store = xStore.BoundText
        .Item = itemTable.Item
        DocTable.Update
    End If
    itemTable.MoveNext
    If itemTable.EOF Then Exit Do
Loop
fillgrd
End With
End Sub
Private Sub Command2_Click()
    mydb.Execute " UPDATE FILE0_10 LEFT JOIN FILE1_10 ON FILE0_10.item = FILE1_10.ITEM SET FILE0_10.COST =  [FILE1_10].[COST], FILE1_10.PRICE = [FILE1_10].[COST]*[FILE0_10].[Differ] "
    
    Dim ITEMcost As Recordset
    Set ITEMcost = mydb.OpenRecordset("ITEMCOST")
    With DocTable
        .Requery
        .MoveFirst
        Do While True And Not .EOF
            .Edit
            ITEMcost.FindLast " ITEM = " & MyParn(.Item) & " and date < DateValue(" & MyParn(.Date) & ")"
            If Not ITEMcost.NoMatch Then
                .COST = ITEMcost.price
            End If
            .total = .COST * .Differ
            Me.Command2.Caption = .Item & "    " & .doc_no
            .Update
            .MoveNext
        Loop
    End With
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set DocTable = mydb.OpenRecordset("select * from FILE0_10 order by doc_no ", dbOpenDynaset)
Set itemTable = mydb.OpenRecordset("file1_10", dbOpenDynaset)
DATA1.DatabaseName = MdbPath
DATA1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
xStore.ListField = "Desca"
xStore.BoundColumn = "code"
Cmd_Print.Visible = bopt3
xDate.Text = Format(Date, "dd-mm-yyyy")
xStore.BoundText = ""
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.doc_no)
Else
    xDoc_No.Text = "000001"
End If
With ItemInv
    .Cols = 7
    .Rows = 1
    .Editable = flexEDKbdMouse
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·’‰‹‹‹‹‹‹›"
    .TextMatrix(0, 2) = "—’Ìœ ﬂÊ„»ÌÊ —"
    .TextMatrix(0, 3) = "—’Ìœ Ã—œ"
    .TextMatrix(0, 4) = "«·›—ﬁ"
    .TextMatrix(0, 5) = " ﬂ·›…"
    .TextMatrix(0, 6) = "«· ﬁÌÌ„"
    
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 4000
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(0) = flexDTString
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(5) = flexAlignRightCenter
    .ColAlignment(6) = flexAlignRightCenter
End With
End Sub
Sub dispProc()
formMode = dispMode
End Sub
Private Sub ItemInv_AfterMoveColumn(ByVal Col As Long, Position As Long)
    If ItemInv.Row + 1 = ItemInv.Rows And ItemInv.Col + 1 = ItemInv.Cols Then
        ItemInv.Rows = ItemInv.Rows + 1
    End If
End Sub
Private Sub ItemInv_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If ItemInv.Col = 2 Or ItemInv.Col = 3 Or ItemInv.Col = 4 Then
    KeyAscii = RetNumber(KeyAscii, True)
End If
End Sub
Private Sub ItemInv_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        ItemInv.RemoveItem ItemInv.Row
    End If
End If
End Sub
Private Sub ItemInv_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case ItemInv.Col
    Case 0
        If KeyCode = 27 Then
            Exit Sub
        End If
        If KeyCode = 112 Then
            ItemsLookup
        End If
End Select
End Sub
Private Sub ItemInv_SelChange()
    If ItemInv.Row + 1 = ItemInv.Rows And ItemInv.Col + 1 = ItemInv.Cols Then
        ItemInv.Rows = ItemInv.Rows + 1
    End If
End Sub
Private Sub ItemInv_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = False
    Select Case ItemInv.Col
        Case 0
            itemTable.FindFirst " ITEM = " & MyParn(ItemInv.EditText)
            If itemTable.NoMatch Or ItemInv.EditText = "" Then
                Cancel = True
            Else
                ItemInv.TextMatrix(ItemInv.Row, 1) = itemTable.desca
'                ItemInv.TextMatrix(ItemInv.Row, 2) = BalItemDateStore(ItemInv.EditText)
            End If
        Case 3
'            ItemInv.TextMatrix(ItemInv.Row, 4) = Val(ItemInv.EditText) - Val(ItemInv.TextMatrix(ItemInv.Row, 2))
    End Select
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xDoc_No.Text = ""
    Dim Generalarray(4)
    Dim GrdArray(4)
    Set Generalarray(1) = Me
    Generalarray(2) = "SELECT FILE0_10.DOC_NO, file0_10.[date] ,First(FILE1_70.DESCA) AS StoreDesc , iif(closed, '„—Õ·','€Ì— „—Õ·')" & _
                      " FROM FILE0_10 LEFT JOIN FILE1_70 ON FILE0_10.Store = FILE1_70.CODE WHERE FILE1_70.FLAG = 1 "
    Generalarray(3) = "AND FILE1_70.DESCA Like '*cFilter*'"
    Generalarray(4) = " Group by FILE0_10.DOC_NO , file0_10.[date],closed"
    GrdArray(1) = 1000
    GrdArray(2) = 1500
    GrdArray(3) = 3000
    GrdArray(4) = 1000
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Show 1
End If
End Sub
Private Sub xDoc_No_LostFocus()
DocValid
End Sub
Function myDelete()
    ' Õ–›  «·„” ‰œ
    cString = " DELETE  * FROM FILE0_10  WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
End Function
Private Function BalItemDateStore(cItem)
    With movetable
    BalItemDateStore = 0
   .FindFirst " ITEM = " & MyParn(cItem)
    If Not movetable.NoMatch Then
        BalItemDateStore = TurnValue(.[SUMIN], Null, 0) - TurnValue(.SUMOUT, Null, 0)
    End If
    End With
End Function
Private Sub ItemInv_LeaveCell()
    If ItemInv.Col = 1 Or ItemInv.Col = 2 Or ItemInv.Col = 4 Then
        ItemInv.Editable = flexEDKbdMouse
    End If
End Sub
Private Sub ItemInv_EnterCell()
    If ItemInv.Col = 1 Or ItemInv.Col = 2 Or ItemInv.Col = 4 Then
        ItemInv.Editable = flexEDNone
    End If
End Sub

Private Sub CmdFirst_Click()
DocTable.MoveFirst
xDoc_No.Text = DocTable!doc_no
DocValid
End Sub
Private Sub CmdLast_Click()
DocTable.MoveLast
xDoc_No.Text = DocTable!doc_no
DocValid
End Sub
Private Sub CmdNext_Click()
DocTable.FindLast " DOC_NO = " & MyParn(xDoc_No)
DocTable.MoveNext
If Not DocTable.EOF Then
    xDoc_No.Text = DocTable!doc_no
    DocValid
End If
End Sub
Private Sub CmdPrevious_Click()
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No)
DocTable.MovePrevious
If Not DocTable.BOF Then
    xDoc_No.Text = DocTable!doc_no
    DocValid
End If
End Sub

