VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Vs_Trans 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕÊÌ·«  »Ì‰ «·„Œ«“‰"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
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
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Left            =   3525
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   675
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox xDescA 
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
      Left            =   3525
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1425
      Width           =   4440
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   360
      Left            =   1537
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1425
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
      Left            =   2925
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
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
      Left            =   8025
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1290
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   1290
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   " ”ÃÌ· «·„” ‰œ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   75
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1425
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
      Left            =   1275
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   525
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
      Left            =   6675
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   525
      Width           =   1290
   End
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   4665
      Left            =   150
      TabIndex        =   6
      Top             =   2025
      Width           =   9135
      _cx             =   16113
      _cy             =   8229
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
      FormatString    =   $"Vs_Trans.frx":0000
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
   Begin MSDBCtls.DBCombo xStore1 
      Bindings        =   "Vs_Trans.frx":0088
      Height          =   315
      Left            =   5325
      TabIndex        =   2
      Top             =   975
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDBCtls.DBCombo xStore2 
      Bindings        =   "Vs_Trans.frx":009C
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì‹‹‹‹‹‹‹‹«‰"
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
      Left            =   8175
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1425
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï „Œ“‰"
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
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1080
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰ „Œ“‰"
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
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
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
      Height          =   315
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   555
      Width           =   1080
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
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   555
   End
End
Attribute VB_Name = "Vs_Trans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DocTable As Recordset, itemTable As Recordset
Dim formMode
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
Sub Fillgrd()
ItemInv.Rows = 1
I = 1
With ItemInv
.FixedRows = 1
.ExplorerBar = flexExSortShow
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then Exit Sub
Do While True
   .AddItem ""
    .TextMatrix(I, 0) = DocTable.Item
    .TextMatrix(I, 1) = RetFind(itemTable, "ITEM", "DescA", DocTable!Item)
    .TextMatrix(I, 2) = TurnValue(Format(DocTable.Quant, "###0.00"), Null, "")
    .TextMatrix(I, 3) = Format(DocTable.cost, "##0.00")
    .TextMatrix(I, 4) = Format(DocTable.total, "##0.00")
    
    DocTable.MoveNext
    If DocTable.EOF Then Exit Sub
    If DocTable.DOC_NO <> xDoc_No.Text Then Exit Sub
    I = I + 1
Loop
End With
End Sub
Sub ItemsLookup()
    ActiveControl.Text = ""
    Dim Generalarray(4)
    Dim GrdArray(2)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Item as «·’‰›,DescA as [«”„ «·’‰›] From file1_10 "
    Generalarray(3) = " Where DescA Like('*cFilter*')"
    Generalarray(4) = "Order by Item"
    GrdArray(1) = 1000
    GrdArray(2) = 4500
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Function MyReplace()
MyReplace = True
' „—«Ã⁄… «·„” ‰œ
With ItemInv
For I = 1 To .Rows - 1
    itemTable.FindFirst " ITEM = " & MyParn(.TextMatrix(I, 0))
    If itemTable.NoMatch And .TextMatrix(I, 0) <> "" Then
        .Select I, 0, I, 2
        cMess = "·‰ Ì „  ”ÃÌ· «·’‰› " & .TextMatrix(I, 0) & " ﬂÊœ «·’‰› €Ì— „”Ã· "
        MsgBox cMess
        MyReplace = False
    End If
Next I
End With

If MyReplace Then
    ' Õ–› «·„” ‰œ ﬁ»· «· ⁄œÌ·
    cString = " DELETE  FILE1_60.*  FROM FILE1_60 WHERE FILE1_60.DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    DocTable.Requery
    '  ”ÃÌ· «·„” ‰œ
    With ItemInv
    For I = 1 To .Rows - 1
        If .TextMatrix(I, 0) <> "" Then
        itemTable.FindFirst " ITEM = " & MyParn(.TextMatrix(I, 0))
        DocTable.AddNew
        DocTable.DOC_NO = xDoc_No.Text
        DocTable.[Date] = DateValue(xDate.Text)
        DocTable.store1 = xStore1.BoundText
        DocTable.Store2 = xStore2.BoundText
        DocTable.Item = .TextMatrix(I, 0)
        DocTable.Quant = Val(.TextMatrix(I, 2))
        DocTable.cost = Val(.TextMatrix(I, 3))
        DocTable.total = Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 2))
        DocTable.DESCA = TurnValue(xDescA.Text, "", Null)
        DocTable.Update
        End If
    Next I
    End With
    
    ' Õ–› Õ—ﬂ… √’‰«› «·„” ‰œ
    cString = " DELETE  FILE1_11.* FROM FILE1_11 WHERE FILE1_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND FILE1_11.[TYPE] = 'F'"
    mydb.Execute cString
    cString = " DELETE  FILE1_11.* FROM FILE1_11 WHERE FILE1_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND FILE1_11.[TYPE] = 'T'"
    mydb.Execute cString
    
    cString = "INSERT INTO FILE1_11( " & _
               "[Date],Doc_Id,[Type],item,[IN],PRICE,DescA,STORE)" & _
               " Select [Date],Doc_No,'T'," & _
               "item,Quant,COST," & _
               MyParn(xDescA.Text) & " &   '  ÕÊÌ· „‰ ' & " & MyParn(xStore1.Text) & " , " & xStore2.BoundText & _
               " From File1_60 " & _
               " WHERE FILE1_60.DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    
    cString = "INSERT INTO FILE1_11( " & _
               "[Date],Doc_Id,[Type],item,[OUT],PRICE,DescA,STORE)" & _
               " Select [Date],Doc_No,'T'," & _
               "item,Quant,COST," & _
               MyParn(xDescA.Text) & " &   '  ŒÊÌ· ≈·Ï ' & " & MyParn(xStore2.Text) & " , " & xStore1.BoundText & _
               " From File1_60 " & _
               " WHERE FILE1_60.DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    DocTable.Requery
End If
End Function
Sub ApplyProc()
If Not DocTable.EOF Then
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then
    
    xDoc_No.Enabled = True
Else
    xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
    xStore1.BoundText = TurnValue(DocTable.store1, Null, "")
    xStore2.BoundText = TurnValue(DocTable.Store2, Null, "")
    xDescA.Text = TurnValue(DocTable.DESCA, Null, "")
    Fillgrd
    dispProc
    xDoc_No.Enabled = False
End If
End If
End Sub
Sub myProc()
If ActiveControl.Name = ItemInv.Name Then
    ItemInv.EditText = GrdText(Search.Grid1, 0)
    ItemInv.TextMatrix(ItemInv.Row, 0) = GrdText(Search.Grid1, 0)
    itemTable.FindFirst " ITEM = " & MyParn(ItemInv.TextMatrix(ItemInv.Row, 0))
    If Not itemTable.NoMatch Then
        ItemInv.TextMatrix(ItemInv.Row, 1) = itemTable.DESCA
    End If
Else
    ActiveControl.Text = GrdText(Search.Grid1, 0)
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
If xStore1.BoundText = "" Or xStore2.BoundText = "" Then
    MsgBox " ”ÃÌ· «·„Œ“‰ "
    MYVALID = False
End If
End Function
Private Sub Cmd_Inv_Click()
xDoc_No.Enabled = False
ItemInv.Enabled = True
ItemInv.SetFocus
ItemInv.Rows = 2
End Sub
Private Sub cmdAdditem_Click()
    With ItemInv
    .Rows = .Rows + 1
    .Select .Rows - 1, 0, .Rows - 1, 4
    .ShowCell .Rows - 1, 0
    End With
End Sub
Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        myDelete
        xDoc_No.Text = ""
        xStore1.BoundText = ""
        xStore2.BoundText = ""
        xDescA.Text = ""
        Fillgrd
        xDoc_No.Enabled = True
        ItemInv.Enabled = False
    End If
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdNewInv_Click()
ItemInv.Rows = 1
ItemInv.AddItem ""
xStore1.BoundText = ""
xStore2.BoundText = ""
xDate.Text = Date
xDescA.Text = ""
xDoc_No.Enabled = True
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.DOC_NO)
Else
    xDoc_No.Text = "000001"
End If
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Fillgrd
End Sub
Private Sub CmdUndo_Click()
    DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
    If Not DocTable.NoMatch Then
        xStore1.BoundText = DocTable.store1
        xStore2.BoundText = DocTable.Store2
        xDescA.Text = DocTable.DESCA
        xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
        Fillgrd
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set DocTable = mydb.OpenRecordset("SELECT * FROM FILE1_60 ORDER BY DOC_NO ,ITEM", dbOpenDynaset)
Set itemTable = mydb.OpenRecordset("file1_10", dbOpenDynaset)

Data1.DatabaseName = MdbPath
Data1.RecordSource = "Stores"
xStore1.BoundColumn = "Code"
xStore1.ListField = "DescA"
xStore2.BoundColumn = "Code"
xStore2.ListField = "DescA"

xDate.Text = Format(Date, "dd-mm-yyyy")
xStore1.BoundText = ""
xStore2.BoundText = ""
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.DOC_NO)
Else
    xDoc_No.Text = "000001"
End If
With ItemInv
    .Cols = 5
    .Rows = 1
    .Editable = flexEDKbdMouse
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·’‰›"
    .TextMatrix(0, 2) = "«·ﬂ„Ì…"
    .TextMatrix(0, 3) = " ﬂ·›…"
    .TextMatrix(0, 4) = "≈Ã„«·Ï"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 3000
    .ColWidth(2) = 1100
    .ColWidth(3) = 1100
    .ColWidth(4) = 1100
    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignRightCenter
End With
End Sub
Sub dispProc()
formMode = dispMode
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
                ItemInv.TextMatrix(ItemInv.Row, 1) = itemTable.DESCA
                ItemInv.TextMatrix(ItemInv.Row, 3) = itemTable.cost
            End If
        Case 2
            Cancel = False
            ItemInv.TextMatrix(ItemInv.Row, 4) = Val(ItemInv.TextMatrix(ItemInv.Row, 2)) * Val(ItemInv.TextMatrix(ItemInv.Row, 3))
    End Select
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(4)
    Dim GrdArray(5)
    Set Generalarray(1) = Me
    Generalarray(2) = "SELECT FILE1_60.DOC_NO as [„” ‰œ], format(FILE1_60.DATE ,'dd-mm-yyyy') as [ «—ÌŒ] , FILE1_60.DESCA as [«·»Ì«‰], first(STORES.DESCA) as [„‰ „Œ“‰], first(STORES_1.DESCA) as [≈·Ï „Œ“‰] " & _
                      " FROM (FILE1_60 LEFT JOIN STORES ON FILE1_60.STORE1 = STORES.CODE) LEFT JOIN STORES AS  " & _
                      " STORES_1 ON FILE1_60.STORE2 = STORES_1.CODE  "
    Generalarray(3) = " Where FILE1_60.DESCA Like '*cFilter*'"
    Generalarray(4) = " Group by Doc_No,FILE1_60.DESCA,[date] ORDER BY [DATE]"
    GrdArray(1) = 1000
    GrdArray(2) = 1200
    GrdArray(3) = 3000
    GrdArray(4) = 1200
    GrdArray(5) = 1200
    
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
    cString = " DELETE  * FROM FILE1_60 WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    ' Õ–› Õ—ﬂ… √’‰«› «·›« Ê—…
    cString = " DELETE  FILE1_11.* FROM FILE1_11 WHERE FILE1_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = 'F'"
    mydb.Execute cString
    
    cString = " DELETE  FILE1_11.* FROM FILE1_11 WHERE FILE1_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = 'T'"
    mydb.Execute cString
    DocTable.Requery

End Function
Private Sub ItemInv_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If ItemInv.Col = 2 Then
    KeyAscii = RetNumber(KeyAscii, True)
End If
End Sub
Private Sub Command1_Click()
    mydb.Execute " UPDATE FILE1_60 LEFT JOIN FILE1_10 ON FILE1_60.item = FILE1_10.ITEM SET FILE1_60.COST =  [FILE1_10].[COST], FILE1_60.total = [FILE1_10].[COST]*[FILE1_60].[QUANT] "
    
    Dim ITEMcost As Recordset
    Set ITEMcost = mydb.OpenRecordset("ITEMCOST")
    With DocTable
        .Requery
        .MoveFirst
        Do While True And Not .EOF
            .Edit
            ITEMcost.FindLast " ITEM = " & MyParn(.Item) & " and date < DateValue(" & MyParn(.Date) & ")"
            If Not ITEMcost.NoMatch Then
                .cost = ITEMcost.price
            End If
            .total = .cost * .Quant
            Me.Command1.Caption = .Item & "    " & .DOC_NO
            
            .Update
            .MoveNext
        Loop
    End With
End Sub
