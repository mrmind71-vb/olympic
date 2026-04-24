VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Vs_Input 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ê«—œ"
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
      Height          =   540
      Left            =   3675
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1500
      Visible         =   0   'False
      Width           =   2565
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
      Left            =   6600
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   75
      Width           =   1290
   End
   Begin VB.TextBox xDescDoc 
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
      Left            =   4650
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1125
      Width           =   3390
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      Left            =   8025
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   75
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
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Left            =   600
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
      Left            =   6675
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   750
      Width           =   1290
   End
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   3540
      Left            =   150
      TabIndex        =   3
      Top             =   2175
      Width           =   8940
      _cx             =   15769
      _cy             =   6244
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
      FormatString    =   $"Vs_Input.frx":0000
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
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰"
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
      TabIndex        =   12
      Top             =   1230
      Width           =   1080
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
      TabIndex        =   7
      Top             =   705
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
      Left            =   1950
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   825
      Width           =   555
   End
End
Attribute VB_Name = "Vs_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DocTable As Recordset, itemTable As Recordset
Dim itemMoveTable As Recordset
Dim storeTable As Recordset
Dim formMode
Dim myFileName As String, nDiscount As Byte, sStore As String, nPrice As Double
Dim cCost As String
Dim itemMoveType As String
Dim cStrStore As String
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
Sub editProc()
formMode = Editmode
End Sub
Sub EmptyProc()
formMode = EmptyMode
ItemInv.Rows = 1
End Sub
Sub AddProc()
formMode = addmode
ItemInv.AddItem ""
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
    .TextMatrix(I, 0) = DocTable.Store
    .TextMatrix(I, 1) = DocTable.Item
    .TextMatrix(I, 2) = TurnValue(DocTable.DESCA, Null, "")
    .TextMatrix(I, 3) = TurnValue(Format(DocTable.Quant, "###0.00"), Null, "")
    .TextMatrix(I, 4) = Format(DocTable.COST, "##0.00")
    .TextMatrix(I, 5) = Format(DocTable.total, "##0.00")
    
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
Function MyReplace() As Boolean
MyReplace = True
' „—«Ã⁄… «·„” ‰œ
With ItemInv
For I = 1 To .Rows - 1
    itemTable.FindFirst " ITEM = " & MyParn(.TextMatrix(I, 1))
    If itemTable.NoMatch And .TextMatrix(I, 1) <> "" Then
        .Select I, 0, I, 3
        cMess = "·‰ Ì „  ”ÃÌ· «·’‰› " & .TextMatrix(I, 1) & " ﬂÊœ «·’‰› €Ì— „”Ã· "
        MsgBox cMess
        MyReplace = False
    End If
Next I
End With

If MyReplace Then
    ' Õ–› «·„” ‰œ ﬁ»· «· ⁄œÌ·
    cString = " DELETE  " & myFileName & ".*  FROM " & myFileName & "  WHERE " & myFileName & ".DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    
    '  ”ÃÌ· «·„” ‰œ
    With ItemInv
    For I = 1 To .Rows - 1
        If .TextMatrix(I, 1) <> "" Then
        itemTable.FindFirst " ITEM = " & MyParn(.TextMatrix(I, 1))
        DocTable.AddNew
        DocTable.DOC_NO = xDoc_No.Text
        DocTable.[Date] = xDate.Text
        DocTable.Store = .TextMatrix(I, 0)
        DocTable.Item = .TextMatrix(I, 1)
        DocTable.Quant = .TextMatrix(I, 3)
        DocTable.COST = .TextMatrix(I, 4)
        DocTable.total = Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4))
        DocTable.DESCA = TurnValue(.TextMatrix(I, 2), Null, "")
        DocTable.DESCDOC = TurnValue(xDescDoc.Text, Null, "")
        DocTable.Update
        End If
    Next I
    End With
    
    ' Õ–› Õ—ﬂ… √’‰«› «·„” ‰œ
    cString = " DELETE  FILE1_11.* FROM FILE1_11 WHERE FILE1_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND FILE1_11.[TYPE] = " & MyParn(itemMoveType)
    mydb.Execute cString
    
    ' ≈‰‘«¡ Õ—ﬂ… √’‰«› «·„” ‰œ
    Select Case publicFlag
        Case 0
            cString = "INSERT INTO FILE1_11( " & _
                       "[Date],Doc_Id,[Type],item,[IN],PRICE,DescA,STORE)" & _
                       " Select [Date],Doc_No,'4'," & _
                       "item,Quant,COST," & _
                       " 'Ê«—œ ›Ï ' & Format([Date], 'dd-mm-yy'), " & _
                       " Store" & _
                       " From File1_80 " & _
                       " WHERE FILE1_80.DOC_NO = " & MyParn(xDoc_No.Text)
        Case 1
            cString = "INSERT INTO FILE1_11( " & _
                       "[Date],Doc_Id,[Type],item,[OUT],PRICE,DescA,STORE)" & _
                       " Select [Date],Doc_No,'8'," & _
                       "item,Quant,COST," & _
                       " '’«œ— ›Ï ' & Format([Date], 'dd-mm-yy'), " & _
                       " Store" & _
                       " From File1_81 " & _
                       " WHERE FILE1_81.DOC_NO = " & MyParn(xDoc_No.Text)
        
        Case 2
            cString = "INSERT INTO FILE1_11( " & _
                       "[Date],Doc_Id,[Type],item,[OUT],PRICE,DescA,STORE)" & _
                       " Select [Date],Doc_No,'9'," & _
                       "item,Quant,COST," & _
                       " 'Â«·ﬂ ›Ï ' & Format([Date], 'dd-mm-yy'), " & _
                       " Store" & _
                       " From File1_82 " & _
                       " WHERE FILE1_82.DOC_NO = " & MyParn(xDoc_No.Text)
    End Select
    mydb.Execute cString
    Set DocTable = mydb.OpenRecordset("SELECT * FROM " & myFileName & " ORDER BY DOC_NO ,STORE , ITEM", dbOpenDynaset)
End If
End Function
Sub ApplyProc()
If Not DocTable.EOF Then
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then
    EmptyProc
    xDoc_No.Enabled = True
Else
    xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
    xDescDoc.Text = TurnValue(DocTable.DESCDOC, Null, "")
    Fillgrd
    dispProc
    xDoc_No.Enabled = False
End If
End If
End Sub
Sub myProc()
If ActiveControl.Name = ItemInv.Name Then
    ItemInv.EditText = GrdText(Search.Grid1, 0)
    ItemInv.TextMatrix(ItemInv.Row, 1) = GrdText(Search.Grid1, 0)
    itemTable.FindFirst " ITEM = " & MyParn(ItemInv.TextMatrix(ItemInv.Row, 1))
    If Not itemTable.NoMatch Then
        ItemInv.TextMatrix(ItemInv.Row, 2) = itemTable.DESCA
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
End Function
Private Sub Cmd_Inv_Click()
xDoc_No.Enabled = False
ItemInv.Enabled = True
ItemInv.SetFocus
ItemInv.Rows = 2
End Sub
Private Sub cmd_item_Click()
Load items
items.Show 1
End Sub
Private Sub cmdAdditem_Click()
    With ItemInv
    .AddItem ""
    .Select .Rows - 1, 0, .Rows - 1, 4
    .ShowCell .Rows - 1, 0
End With
End Sub
Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        myDelete
        xDoc_No.Text = ""
        xDescDoc.Text = ""
        Fillgrd
        xDoc_No.Enabled = True
        ItemInv.Enabled = False
    End If
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdNewInv_Click()
If Not MyReplace Then Exit Sub
ItemInv.Rows = 1
ItemInv.AddItem ""
xDescDoc.Text = ""
xDate.Text = Date
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
        xDescDoc.Text = DocTable.DESCDOC
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
Select Case publicFlag
    Case 0 ' Ê«—œ
        itemMoveType = "4"
        myFileName = "File1_80"
        Vs_Input.Caption = "Ê«—œ"
    Case 1 ' ’«œ—
        itemMoveType = "8"
        myFileName = "File1_81"
        Vs_Input.Caption = "’«œ— "
    
    Case 2 ' Â«·ﬂ
        itemMoveType = "9"
        myFileName = "File1_82"
        Vs_Input.Caption = "Â«·ﬂ "
End Select
Set DocTable = mydb.OpenRecordset("SELECT * FROM " & myFileName & " ORDER BY DOC_NO ,STORE , ITEM", dbOpenDynaset)
Set storeTable = mydb.OpenRecordset("Stores", dbOpenDynaset)
Set itemTable = mydb.OpenRecordset("file1_10", dbOpenDynaset)
cStrStore = StrStore
xDate.Text = Format(Date, "dd-mm-yyyy")
xDescDoc.Text = ""
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.DOC_NO)
Else
    xDoc_No.Text = "000001"
End If
With ItemInv
    .Cols = 6
    .Rows = 1
    .Editable = flexEDKbdMouse

    .TextMatrix(0, 0) = "„Œ“‰"
    .TextMatrix(0, 1) = "ﬂÊœ"
    .TextMatrix(0, 2) = "«·’‰›"
    .TextMatrix(0, 3) = "«·ﬂ„Ì…"
    .TextMatrix(0, 4) = " ﬂ·›…"
    .TextMatrix(0, 5) = "≈Ã„«·Ï"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1200
    .ColWidth(2) = 2500
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(4) = 1000
    
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(1) = flexDTString
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .ColComboList(0) = cStrStore
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
Private Sub ItemInv_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With ItemInv
If .Row > 0 Then
    Select Case .Col
        Case 1
            If .TextMatrix(ItemInv.Row, 0) = "" Then MsgBox "ÌÃ»  ”Ã· «·„Œ“‰ "
    End Select
End If
End With
End Sub
Private Sub ItemInv_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If ItemInv.Row + 1 = ItemInv.Rows And ItemInv.Col = 0 And (ItemInv.TextMatrix(ItemInv.Rows - 1, 0) = "") Then
        ItemInv.TextMatrix(ItemInv.Rows - 1, 0) = ItemInv.TextMatrix(ItemInv.Rows - 2, 0)
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
    Case 1
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
        Case 1
            itemTable.FindFirst " ITEM = " & MyParn(ItemInv.EditText)
            If itemTable.NoMatch Or ItemInv.EditText = "" Then
                Cancel = True
            Else
                ItemInv.TextMatrix(ItemInv.Row, 2) = itemTable.DESCA
                ItemInv.TextMatrix(ItemInv.Row, 4) = itemTable.COST
            End If
        Case 3
            Cancel = False
            ItemInv.TextMatrix(ItemInv.Row, 5) = Val(ItemInv.TextMatrix(ItemInv.Row, 4)) * Val(ItemInv.TextMatrix(ItemInv.Row, 3))
    End Select
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(4)
    Dim GrdArray(3)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Doc_No as [«·„”·”·], [Date] as [ «—ÌŒ ],DescDOC as [»Ì«‰] " & _
                      " From " & myFileName
    Generalarray(3) = "Where DescDOC Like '*cFilter*'"
    Generalarray(4) = " Group by Doc_No,[date], descDOC ORDER BY [DATE]"
    GrdArray(1) = 1000
    GrdArray(2) = 1500
    GrdArray(3) = 3000
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Show 1
End If
End Sub
Private Sub xDoc_No_LostFocus()
DocValid
End Sub
Private Function StrStore()
If storeTable.RecordCount > 0 Then
    storeTable.MoveFirst
    I = 1
    StrStore = "#" & storeTable!CODE & ";" & storeTable.DESCA
    storeTable.MoveNext
    Do While True
        I = I + 1
        If storeTable.EOF Then Exit Do
        StrStore = StrStore & "|#" & storeTable!CODE & ";" & storeTable.DESCA
        storeTable.MoveNext
    Loop
End If
End Function
Function myDelete()
    ' Õ–›  «·„” ‰œ
    cString = " DELETE  " & myFileName & " .* FROM " & myFileName & " WHERE " & myFileName & ".DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    ' Õ–› Õ—ﬂ… √’‰«› «·›« Ê—…
    cString = " DELETE  FILE1_11.* FROM FILE1_11 WHERE FILE1_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = " & MyParn(itemMoveType)
    mydb.Execute cString
End Function
Private Sub ItemInv_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If ItemInv.Col = 3 Then
    KeyAscii = RetNumber(KeyAscii, True)
End If
End Sub
Private Sub Command1_Click()
    Select Case publicFlag
        Case 0 ' Ê«—œ
            mydb.Execute " UPDATE FILE1_80 LEFT JOIN FILE1_10 ON FILE1_80.item = FILE1_10.ITEM SET FILE1_80.COST =  [FILE1_10].[COST], FILE1_80.total = [FILE1_10].[COST]*[FILE1_80].[QUANT] "
        Case 1 ' ’«œ—
            mydb.Execute " UPDATE FILE1_81 LEFT JOIN FILE1_10 ON FILE1_81.item = FILE1_10.ITEM SET FILE1_81.COST =  [FILE1_10].[COST], FILE1_81.total = [FILE1_10].[COST]*[FILE1_81].[QUANT] "
        Case 2 ' Â«·ﬂ
            mydb.Execute " UPDATE FILE1_82 LEFT JOIN FILE1_10 ON FILE1_82.item = FILE1_10.ITEM SET FILE1_82.COST =  [FILE1_10].[COST], FILE1_82.total = [FILE1_10].[COST]*[FILE1_82].[QUANT] "
    End Select
    
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
            .total = .COST * .Quant
            Me.Command1.Caption = .Item & "    " & .DOC_NO
            
            .Update
            .MoveNext
        Loop
    End With
End Sub
