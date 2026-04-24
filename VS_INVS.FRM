VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Vs_InvS 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›Ê« Ì— «·„»Ì⁄« "
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
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
   ScaleHeight     =   7530
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   555
      ScaleWidth      =   11700
      TabIndex        =   16
      Top             =   0
      Width           =   11760
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0FFFF&
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
         Left            =   75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   75
         Width           =   1290
      End
      Begin VB.CommandButton cmdNewinv 
         BackColor       =   &H00C0FFFF&
         Caption         =   "›« Ê—… ÃœÌœ…"
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
         Left            =   2775
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   75
         Width           =   1290
      End
      Begin VB.CommandButton CmdDelInv 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Õ–› «·›« Ê—… "
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
         Left            =   1425
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   75
         Width           =   1290
      End
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00C0FFFF&
      Caption         =   " —«Ã⁄"
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
      Left            =   225
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2250
      Width           =   1440
   End
   Begin VB.CommandButton CmdAddItem 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3225
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2250
      Width           =   1440
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Õ›Ÿ «·›« Ê—…"
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
      Left            =   1725
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2250
      Width           =   1440
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox xCode 
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
      Left            =   8940
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1275
      Width           =   1500
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
      Left            =   450
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   825
      Width           =   1665
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
      Left            =   8940
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   900
      Width           =   1500
   End
   Begin VB.CheckBox xCash 
      BackColor       =   &H00FFFFFF&
      Caption         =   "‰ﬁœÌ"
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
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1350
      Width           =   840
   End
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   3615
      Left            =   450
      TabIndex        =   15
      Top             =   2925
      Width           =   11040
      _cx             =   19473
      _cy             =   6376
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   12615680
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
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
      FormatString    =   $"Vs_invs.frx":0000
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
      WallPaperAlignment=   4
   End
   Begin VB.Label xBalItem 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4875
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2280
      Width           =   3900
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      Height          =   465
      Left            =   300
      Shape           =   4  'Rounded Rectangle
      Top             =   6825
      Width           =   2490
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   1440
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   675
      Width           =   11415
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   3915
      Left            =   300
      Top             =   2775
      Width           =   11340
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "«·≈Ã„«·Ì :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   6975
      Width           =   810
   End
   Begin VB.Label LblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   375
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   6900
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "—’Ì‹‹‹‹‹‹‹œ"
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
      Left            =   10650
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1725
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ‹‹‹‹‹‹‹‹Êœ"
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
      Left            =   10545
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1365
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·›« Ê—…"
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
      Left            =   10545
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   930
      Width           =   915
   End
   Begin VB.Label xclientDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4950
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1275
      Width           =   3930
   End
   Begin VB.Label xClientBalance 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8940
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1650
      Width           =   1500
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
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
      Left            =   2325
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   900
      Width           =   555
   End
End
Attribute VB_Name = "Vs_InvS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DocTable As Recordset, ClientTable As Recordset, itemTable As Recordset
Dim itemBalTable As Recordset
Dim SerInvTable As Recordset, cStore As String
Dim BalClient As Recordset, SerFileName As String
Dim storeTable As Recordset, COUNTINVTOTAL As Double
Dim ChargesTable As Recordset
Dim formMode, cSerItem As String
Dim itemMoveType As String, ClientMoveType As String
Dim myFileName As String, nDiscount As Byte, sStore As String, nPrice As Double
Dim cCost As String
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
xCode.Text = ""
xDoc_No.Text = ""
xCash.Value = False
xDate.Text = Format(Date, "dd-mm-yyyy")
xDoc_No.Enabled = True
ItemInv.Rows = 1
'        Ser_Item.Rows = 1

End Sub
Sub AddProc()

formMode = addmode
ItemInv.AddItem ""
End Sub
Sub Fillgrd()
COUNTINVTOTAL = 0
ItemInv.Rows = 1
i = 1
With ItemInv
'.ExplorerBar = flexExSortAndMove
.FixedRows = 1
.ExplorerBar = flexExSortShow
LblTotal.Caption = Format(COUNTINVTOTAL, "##0.00")
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then Exit Sub
Do While True
   .AddItem ""
    .TextMatrix(i, 0) = DocTable.Store
    .TextMatrix(i, 1) = DocTable.Item
    .TextMatrix(i, 2) = IIf(DocTable.Store = "SS", RetFind(ChargesTable, "Code", "DescA", DocTable!Item), RetFind(itemTable, "ITEM", "DescA", DocTable!Item))
    .TextMatrix(i, 3) = Format(DocTable.Quant)
    .TextMatrix(i, 4) = TurnValue(Format(DocTable.price, "###0.00"), Null, "")
    .TextMatrix(i, 5) = TurnValue(Format(DocTable.total, "###0.00"), Null, "")
     COUNTINVTOTAL = COUNTINVTOTAL + DocTable.total
    LblTotal.Caption = Format(COUNTINVTOTAL, "##0.00")
    DocTable.MoveNext
    .Cell(flexcpForeColor, 1, 5, .Rows - 1) = RGB(255, 0, 0)
    If DocTable.EOF Then Exit Sub
    If DocTable.Doc_No <> xDoc_No.Text Then Exit Sub
    i = i + 1
Loop
End With
End Sub
Sub ItemsLookup()
'    ActiveControl.Text = ""
    Dim Generalarray(4)
    Dim GrdArray(3)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Item as «·’‰›,DescA as [«”„ «·’‰›] ,price as [«·”⁄—] From file1_10 "
    Generalarray(3) = " Where DescA Like('*cFilter*')"
    Generalarray(4) = "Order by Item"
    GrdArray(1) = 1000
    GrdArray(2) = 3500
    GrdArray(3) = 1000
    
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Function MyReplace()
MyReplace = True
' „—«Ã⁄… «·›« Ê—…
With ItemInv
For i = 1 To .Rows - 1
    If .TextMatrix(i, 0) <> "" Or .TextMatrix(i, 1) <> "" Then
        If .TextMatrix(i, 0) = "SS" Then
            ChargesTable.FindFirst " CODE = " & MyParn(.TextMatrix(i, 1))
            If ChargesTable.NoMatch And .TextMatrix(i, 1) = "" Then
                .Select i, 0, i, 5
                cMess = "·‰ Ì „  ”ÃÌ· «·„’—Ê›" & .TextMatrix(i, 1) & " ﬂÊœ «·„’—Ê› €Ì— „”Ã· "
                MsgBox cMess
                MyReplace = False
            End If
        Else
            itemTable.FindFirst " ITEM = " & MyParn(.TextMatrix(i, 1))
            If itemTable.NoMatch And .TextMatrix(i, 1) = "" Then
                .Select i, 0, i, 5
                cMess = "·‰ Ì „  ”ÃÌ· «·’‰› " & .TextMatrix(i, 1) & " ﬂÊœ «·’‰› €Ì— „”Ã· "
                MsgBox cMess
                MyReplace = False
            End If
        End If
    End If
Next i
End With

If MyReplace Then
    ' Õ–› «·›« Ê—… ﬁ»· «· ⁄œÌ·
    cString = " DELETE  " & myFileName & ".*  FROM " & myFileName & "  WHERE " & myFileName & ".DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    
    '  ”ÃÌ· «·›« Ê—…
    With ItemInv
    
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 1) <> "" Then
            itemTable.FindFirst " ITEM = " & MyParn(.TextMatrix(i, 1))
            DocTable.AddNew
            DocTable.Doc_No = xDoc_No.Text
            DocTable.[Date] = xDate.Text
            DocTable.CODE = xCode.Text
            DocTable.Store = .TextMatrix(i, 0)
            DocTable.Item = .TextMatrix(i, 1)
            DocTable.price = TurnValue(.TextMatrix(i, 4), "", 0)
            DocTable.Quant = TurnValue(.TextMatrix(i, 3), "", 0)
            DocTable.CASH = xCash.Value
            If .TextMatrix(i, 0) <> "SS" Then
                DocTable.total = Val(.TextMatrix(i, 4)) * Val(.TextMatrix(i, 3))
            Else
                DocTable.total = Val(.TextMatrix(i, 5))
            End If
            DocTable.DescA = TurnValue(.TextMatrix(i, 2), Null, "")
            If publicFlag = 1 Or publicFlag = 3 Then
    '            DocTable.COST = itemTable.COST
            End If
            DocTable.Update
         End If
    Next i
    End With
    
    ' Õ–› Õ—ﬂ… √’‰«› «·›« Ê—…
    cString = " DELETE  FILE1_11.* FROM FILE1_11 WHERE FILE1_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND FILE1_11.[TYPE] = " & MyParn(itemMoveType)
    mydb.Execute cString
    
    ' ≈‰‘«¡ Õ—ﬂ… √’‰«› «·›« Ê—…
    Select Case publicFlag
        Case 1
            cString = "INSERT INTO FILE1_11( " & _
                       "[Date],Doc_Id,Code,[Type],item,Out,DescA,Price,Total,STORE)" & _
                       " Select [Date],Doc_No,Code,'6'," & _
                       "item,Quant," & _
                       " '„»Ì⁄«  ›Ï ' & Format([Date], 'dd-mm-yy'), " & _
                       " Price,PRICE * Quant ,Store" & _
                       " From File6_20 " & _
                       " WHERE FILE6_20.STORE <> 'SS' AND FILE6_20.DOC_NO = " & MyParn(xDoc_No.Text)
        Case 2
            cString = "INSERT INTO FILE1_11( " & _
                       "[Date],Doc_Id,Code,[Type],item,[IN],DescA,Price,Total,STORE)" & _
                       " Select [Date],Doc_No,Code,'3'," & _
                       "item,Quant," & _
                       " '„—œÊœ „»Ì⁄«  ›Ï ' & Format([Date], 'dd-mm-yy'), " & _
                       " Price,PRICE * Quant ,Store" & _
                       " From File6_10 " & _
                       " WHERE FILE6_10.STORE <> 'SS' AND FILE6_10.DOC_NO = " & MyParn(xDoc_No.Text)
        
        Case 3
            cString = "INSERT INTO FILE1_11( " & _
                       "[Date],Doc_Id,Code,[Type],item,[IN],DescA,Price,Total,STORE)" & _
                       " Select [Date],Doc_No,Code,'2'," & _
                       "item,Quant," & _
                       " '„‘ —Ì«  ›Ï ' & Format([Date], 'dd-mm-yy'), " & _
                       " Price,PRICE * Quant ,Store" & _
                       " From File7_20 " & _
                       " WHERE FILE7_20.STORE <> 'SS' AND FILE7_20.DOC_NO = " & MyParn(xDoc_No.Text)
        Case 4
            cString = "INSERT INTO FILE1_11( " & _
                       "[Date],Doc_Id,Code,[Type],item,[OUT],DescA,Price,Total,STORE)" & _
                       " Select [Date],Doc_No,Code,'7'," & _
                       "item,Quant," & _
                       " '„—œÊœ „‘ —Ì«  ›Ï ' & Format([Date], 'dd-mm-yy'), " & _
                       " Price,PRICE * Quant ,Store" & _
                       " From File6_11 " & _
                       " WHERE FILE6_11.STORE <> 'SS' AND FILE6_11.DOC_NO = " & MyParn(xDoc_No.Text)
    End Select
    mydb.Execute cString
    
    ' Õ–› Õ—ﬂ… ⁄„Ì· «·›« Ê—…
    If publicFlag = 1 Or publicFlag = 2 Then
        cString = " DELETE  FILE3_11.* FROM FILE3_11 WHERE FILE3_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = " & MyParn(ClientMoveType)
    Else
        cString = " DELETE  FILE4_11.* FROM FILE4_11 WHERE FILE4_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = " & MyParn(ClientMoveType)
    End If
    mydb.Execute cString
    
    ' ≈‰‘«¡ Õ—ﬂ… ⁄„Ì· ··›« Ê—…
    Select Case publicFlag
        Case 1
            cString = "Insert Into File3_11(" & _
                      "[Type],Doc_Id,Code,[Date],Sal,DescA)" & _
                      " Select '3',Doc_No,Code,[Date],Sum(total),'„»Ì⁄« ' From File6_20 " & _
                      " WHERE FILE6_20.DOC_NO = " & MyParn(xDoc_No.Text) & _
                      " Group by Doc_No,Code,[Date]"
        Case 2
            cString = "Insert Into File3_11(" & _
                      "[Type],Doc_Id,Code,[Date],PAY,DescA)" & _
                      " Select '5',Doc_No,Code,[Date],Sum(total),'„—œÊœ „»Ì⁄« ' From File6_10 " & _
                      " WHERE FILE6_10.DOC_NO = " & MyParn(xDoc_No.Text) & _
                      " Group by Doc_No,Code,[Date]"
        Case 3
            cString = "Insert Into File4_11(" & _
                      "[Type],Doc_Id,Code,[Date],Sal,DescA)" & _
                      " Select '4',Doc_No,Code,[Date],Sum(total),'„‘ —Ì« ' From File7_20 " & _
                      " WHERE FILE7_20.DOC_NO = " & MyParn(xDoc_No.Text) & _
                      " Group by Doc_No,Code,[Date]"
        Case 4
            cString = "Insert Into File4_11(" & _
                      "[Type],Doc_Id,Code,[Date],PAY,DescA)" & _
                      " Select '5',Doc_No,Code,[Date],Sum(total),'„—œÊœ „‘ —Ì«  ' From File6_11 " & _
                      " WHERE FILE6_11.DOC_NO = " & MyParn(xDoc_No.Text) & _
                      " Group by Doc_No,Code,[Date]"
    End Select
    mydb.Execute cString
    
    If publicFlag = 1 Or publicFlag = 2 Then
        Set BalClient = mydb.OpenRecordset("SELECT FILE3_11.CODE, Sum(FILE3_11.PAY) AS SumPay, Sum(FILE3_11.SAL) AS SumSal FROM FILE3_11 GROUP BY FILE3_11.CODE ")
    Else
        Set BalClient = mydb.OpenRecordset("SELECT FILE4_11.CODE, Sum(FILE4_11.PAY) AS SumPay, Sum(FILE4_11.SAL) AS SumSal FROM FILE4_11 GROUP BY FILE4_11.CODE ")
    End If

    BalClient.FindFirst " CODE = " & MyParn(xCode.Text)
    If Not BalClient.NoMatch Then
        xClientBalance.Caption = Format(TurnValue(BalClient.SUMSAL, Null, 0) - TurnValue(BalClient.SUMPAY, Null, 0), "###0.00")
    End If
    Set DocTable = mydb.OpenRecordset("SELECT * FROM " & myFileName & " ORDER BY DOC_NO ,STORE , ITEM", dbOpenDynaset)
Set itemBalTable = mydb.CreateSnapshot("SELECT ITEM, STORE, Sum(OUT) AS SumOUT, Sum([IN]) AS SumIN FROM FILE1_11 GROUP BY ITEM, STORE ORDER BY ITEM , STORE  ", dbOpenDynaset)
End If
End Function
Sub ApplyProc()
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
xDoc_No.Enabled = DocTable.NoMatch
If Not DocTable.EOF Then
If DocTable.NoMatch Then
    EmptyProc
Else
    xCode.Text = DocTable.CODE
    xCash.Value = IIf(DocTable.CASH, 1, 0)
    xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
    If publicFlag = 1 Or publicFlag = 2 Then
        Set BalClient = mydb.OpenRecordset("SELECT FILE3_11.CODE, Sum(FILE3_11.PAY) AS SumPay, Sum(FILE3_11.SAL) AS SumSal FROM FILE3_11 GROUP BY FILE3_11.CODE ")
    Else
        Set BalClient = mydb.OpenRecordset("SELECT FILE4_11.CODE, Sum(FILE4_11.PAY) AS SumPay, Sum(FILE4_11.SAL) AS SumSal FROM FILE4_11 GROUP BY FILE4_11.CODE ")
    End If
    BalClient.FindFirst " CODE = " & MyParn(xCode.Text)
    If Not BalClient.NoMatch Then
        xClientBalance.Caption = Format(TurnValue(BalClient.SUMSAL, Null, 0) - TurnValue(BalClient.SUMPAY, Null, 0), "###0.00")
    End If
    Fillgrd
    dispProc
    xDoc_No.Enabled = False
End If
End If
End Sub
Sub myProc()
If ActiveControl.Name = ItemInv.Name Then
    If ItemInv.TextMatrix(ItemInv.Row, 0) = "SS" Then
        ItemInv.EditText = GrdText(Search.Grid1, 0)
        ItemInv.TextMatrix(ItemInv.Row, 1) = GrdText(Search.Grid1, 0)
        ChargesTable.FindFirst " CODE = " & MyParn(ItemInv.TextMatrix(ItemInv.Row, 1))
        If Not ChargesTable.NoMatch Then
'           ItemInv.TextMatrix(ItemInv.Row, 2) = ChargesTable.DESCA
        End If
    Else
        ItemInv.EditText = GrdText(Search.Grid1, 0)
        ItemInv.TextMatrix(ItemInv.Row, 1) = GrdText(Search.Grid1, 0)
        itemTable.FindFirst " ITEM = " & MyParn(ItemInv.TextMatrix(ItemInv.Row, 1))
        If Not itemTable.NoMatch Then
'           ItemInv.TextMatrix(ItemInv.Row, 2) = itemTable.DESCA
'           If ItemInv.TextMatrix(ItemInv.Row, 4) <> "" Then ItemInv.TextMatrix(ItemInv.Row, 4) = itemTable.price
        End If
    End If
Else
    ActiveControl.Text = GrdText(Search.Grid1, 0)
End If
Unload Search
End Sub
Function MYVALID()
MYVALID = True
If xDoc_No.Text = "" Then
    MsgBox " ”ÃÌ· —ﬁ„ «·›« Ê—…"
    MYVALID = False
End If
If xDate.Text = "" Or Not IsDate(xDate.Text) Then
    MsgBox " ”ÃÌ· «· «—ÌŒ"
    MYVALID = False
End If
If xCode.Text = "" Then
    MsgBox " ”ÃÌ· «·⁄„Ì·"
    MYVALID = False
End If

ClientTable.FindFirst "code = " & MyParn(xCode.Text)
If ClientTable.NoMatch Then
    MsgBox "ﬂÊœ «·⁄„Ì· €Ì— „”Ã·"
    MYVALID = False
End If

End Function
Function turnPrice()
turnPrice = ""
If priceTable.RecordCount = 0 Then Exit Function
If IsNull(ClientTable.PriceList) Then Exit Function
priceTable.Seek ">=", ClientTable.PriceList, xItem.Text
If priceTable.NoMatch Then Exit Function
turnPrice = IIf(priceTable.NoMatch, TurnValue(priceTable.price, 0, ""), TurnValue(itemTable.price, 0, ""))
End Function
Sub Undoinv()
Select Case formMode
Case addmode
    invGrid.Rows = invGrid.Rows - 1
    dispProc
Case Editmode
    dispProc
Case EmptyMode
    
End Select
End Sub
Private Sub Cmd_Inv_Click()
xDoc_No.Enabled = False
ItemInv.Enabled = True
ItemInv.SetFocus
ItemInv.Rows = 2
End Sub
Private Sub cmdAdditem_Click()
ItemInv.AddItem ""
ItemInv.ShowCell ItemInv.Rows - 1, 0
ItemInv.TextMatrix(ItemInv.Rows - 1, 0) = cStore
ItemInv.Row = ItemInv.Rows - 1
ItemInv.Col = 0

'    Dim nTest As Double
'    Ser_Item.Rows = 1
'    With ItemInv
'    .Rows = .Rows + 1
'    For i = ItemInv.Rows - 1 To 2 Step -1
'        .TextMatrix(i, 0) = .TextMatrix(i - 1, 0)
'        .TextMatrix(i, 1) = .TextMatrix(i - 1, 1)
'        .TextMatrix(i, 2) = .TextMatrix(i - 1, 2)
'        .TextMatrix(i, 3) = .TextMatrix(i - 1, 3)
'        .TextMatrix(i, 4) = .TextMatrix(i - 1, 4)
'        .TextMatrix(i, 5) = .TextMatrix(i - 1, 5)
'    Next i
'    .TextMatrix(1, 0) = ""
'    .TextMatrix(1, 1) = ""
'    .TextMatrix(1, 2) = ""
'    .TextMatrix(1, 3) = ""
'    .TextMatrix(1, 4) = ""
'    .TextMatrix(1, 5) = ""
'    .Cell(flexcpForeColor, 1, 5, .Rows - 1) = RGB(255, 0, 0)
'    End With
End Sub
Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·›« Ê—… »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        myDelete
        EmptyProc
'        Ser_Item.Rows = 1
    End If
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdNewInv_Click()
'If Not MyReplace Then Exit Sub
ItemInv.Rows = 1
ItemInv.AddItem ""
'Ser_Item.Rows = 1
xCode.Text = ""
'xDate.Text = Date
xClientBalance.Caption = ""
xclientDescA.Caption = ""
xCash.Value = 0
xDoc_No.Enabled = True
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.Doc_No)
Else
    xDoc_No.Text = "000001"
End If
xDoc_No.SetFocus
LblTotal.Caption = Format(0, "##0.00")
'Handlecontrols NewInvMode
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
'SendKeys "{tab}"
If Not MyReplace Then Exit Sub
Fillgrd
End Sub
Private Sub cmdundo_Click()
    DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
    If Not DocTable.NoMatch Then
        xCode.Text = DocTable.CODE
        xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
        If publicFlag = 1 Or publicFlag = 2 Then
            Set BalClient = mydb.OpenRecordset("SELECT FILE3_11.CODE, Sum(FILE3_11.PAY) AS SumPay, Sum(FILE3_11.SAL) AS SumSal FROM FILE3_11 GROUP BY FILE3_11.CODE ")
        Else
            Set BalClient = mydb.OpenRecordset("SELECT FILE4_11.CODE, Sum(FILE4_11.PAY) AS SumPay, Sum(FILE4_11.SAL) AS SumSal FROM FILE4_11 GROUP BY FILE4_11.CODE ")
        End If
        BalClient.FindFirst " CODE = " & MyParn(xCode.Text)
        If Not BalClient.NoMatch Then
            xClientBalance.Caption = Format(TurnValue(BalClient.SUMSAL, Null, 0) - TurnValue(BalClient.SUMPAY, Null, 0), "###0.00")
        End If
        Fillgrd
    End If
End Sub
Private Sub Command1_Click()
cString = "Insert into FILE6_20(Doc_No,[Date],Code,Store,Item,Total,Quant,price,DAY_DOC) " & _
          " Select  doc_no ,date,Code,Store,Item,Total,quant,price,DOC_NO" & _
          " From File6_30 "
mydb.Execute cString
End Sub
Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
Command1.Caption = KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Select Case publicFlag
    Case 1 ' „»Ì⁄« 
        itemMoveType = "6"
        myFileName = "File6_20"
        ClientMoveType = "4"
        Set ClientTable = mydb.OpenRecordset("File3_10", dbOpenDynaset)
        Set BalClient = mydb.OpenRecordset("SELECT FILE3_11.CODE, Sum(FILE3_11.PAY) AS SumPay, Sum(FILE3_11.SAL) AS SumSal FROM FILE3_11 GROUP BY FILE3_11.CODE ")
        Me.Caption = "›« Ê—… „»Ì⁄« "
    Case 2 ' „—œÊœ „»Ì⁄« 
        itemMoveType = "3"
        myFileName = "File6_10"
        ClientMoveType = "5"
        Set ClientTable = mydb.OpenRecordset("File3_10", dbOpenDynaset)
        Set BalClient = mydb.OpenRecordset("SELECT FILE3_11.CODE, Sum(FILE3_11.PAY) AS SumPay, Sum(FILE3_11.SAL) AS SumSal FROM FILE3_11 GROUP BY FILE3_11.CODE ")
        Me.Caption = "›« Ê—… „—œÊœ „»Ì⁄« "
    Case 3 ' „‘ —Ì« 
        itemMoveType = "2"
        myFileName = "File7_20"
        ClientMoveType = "4"
        Set ClientTable = mydb.OpenRecordset("File4_10", dbOpenDynaset)
        Set BalClient = mydb.OpenRecordset("SELECT FILE4_11.CODE, Sum(FILE4_11.PAY) AS SumPay, Sum(FILE4_11.SAL) AS SumSal FROM FILE4_11 GROUP BY FILE4_11.CODE ")
        Me.Caption = "›« Ê—… „‘ —Ì« "
    Case 4
        itemMoveType = "7"
        myFileName = "File6_11"
        ClientMoveType = "5"
        Set ClientTable = mydb.OpenRecordset("File4_10", dbOpenDynaset)
        Set BalClient = mydb.OpenRecordset("SELECT FILE4_11.CODE, Sum(FILE4_11.PAY) AS SumPay, Sum(FILE4_11.SAL) AS SumSal FROM FILE4_11 GROUP BY FILE4_11.CODE ")
        Me.Caption = "›« Ê—… „—œÊœ «·„‘ —Ì«  "
End Select

Set DocTable = mydb.OpenRecordset("SELECT * FROM " & myFileName & " ORDER BY DOC_NO ,STORE , ITEM", dbOpenDynaset)
Set storeTable = mydb.OpenRecordset("Stores", dbOpenDynaset)
Set itemTable = mydb.OpenRecordset("file1_10", dbOpenDynaset)
Set ChargesTable = mydb.CreateSnapshot("Select * From File8_70 ", dbOpenDynaset)
Set itemBalTable = mydb.CreateSnapshot("SELECT ITEM, STORE, Sum(OUT) AS SumOUT, Sum([IN]) AS SumIN FROM FILE1_11 GROUP BY ITEM, STORE ORDER BY ITEM , STORE  ", dbOpenDynaset)
xDate.Text = Format(Date, "dd-mm-yyyy")
cStrStore = StrStore
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.Doc_No)
Else
    xDoc_No.Text = "000001"
End If
With ItemInv
    .Cols = 6
    .Rows = 1
'    .ColEditMask(5) = ""
'    .Editable = flexEDKbdMouse
    .FormatString = "„Œ“‰|" & "ﬂÊœ|" & "«·’‰‹‹‹‹‹‹›|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "«·≈Ã„«·‹‹‹Ï"
    .ColWidth(0) = 1800
    .ColWidth(1) = 1200
    .ColWidth(2) = 5000
    .ColWidth(3) = 900
    .ColWidth(4) = 900
    .ColWidth(5) = 1100
    
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(1) = flexDTString
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(5) = flexAlignRightCenter
    .ColComboList(0) = cStrStore
End With

'With Ser_Item
'    .Cols = 1
'    .Rows = 1
'    .Editable = flexEDKbdMouse
'    .FormatString = "—ﬁ„ «·„”·”·"
'    .ColWidth(0) = 1500
'    .ColDataType(0) = flexDTString
'End With

'xDoc_No.SetFocus
End Sub
Sub dispProc()
formMode = dispMode
'Handlecontrols dispMode
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
    Case 5
       If .TextMatrix(.Row, 4) <> "" And .TextMatrix(.Row, 3) <> "" And .TextMatrix(.Row, 0) <> "SS" Then
            .TextMatrix(.Row, 5) = .TextMatrix(.Row, 4) * .TextMatrix(.Row, 3)
        End If
End Select
End If
End With
End Sub
Private Sub ItemInv_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If ItemInv.Row + 1 = ItemInv.Rows And ItemInv.Col = 0 And (ItemInv.TextMatrix(ItemInv.Rows - 1, 0) = "") Then
        ItemInv.TextMatrix(ItemInv.Rows - 1, 0) = ItemInv.TextMatrix(ItemInv.Rows - 2, 0)
        ItemInv.Cell(flexcpForeColor, ItemInv.Rows - 1, 5, ItemInv.Rows - 1) = RGB(255, 0, 0)
    End If
End Sub
Private Sub ItemInv_CellChanged(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then
cStore = ItemInv.TextMatrix(Row, 0)
End If
End Sub
Private Sub ItemInv_EnterCell()
If ItemInv.TextMatrix(ItemInv.Row, 0) = "SS" Then
    ItemInv.Editable = IIf(ItemInv.Col = 5 Or ItemInv.Col = 0 Or ItemInv.Col = 1, flexEDKbdMouse, flexEDNone)
Else
    ItemInv.Editable = IIf(ItemInv.Col = 2 Or ItemInv.Col = 5, flexEDNone, flexEDKbdMouse)
End If
End Sub
Private Sub ItemInv_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If ItemInv.Col = 3 Or ItemInv.Col = 4 Or ItemInv.Col = 5 Then
    KeyAscii = RetNumber(KeyAscii, True)
End If
End Sub
Private Sub ItemInv_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 46
    If MsgBox("Õ–› «·’‰› „‰ «·›« Ê—… ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        ItemInv.RemoveItem ItemInv.Row
    End If
Case 45
    cmdAdditem_Click
Case 112
    If ItemInv.Col = 1 Then
        If ItemInv.TextMatrix(ItemInv.Row, 0) = "SS" Then
            ChargesLookup
        Else
            ItemsLookup
        End If
    End If
End Select
End Sub
Private Sub ItemInv_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case KeyCode
Case 112
    If ItemInv.Col = 1 Then
        If ItemInv.TextMatrix(ItemInv.Row, 0) = "SS" Then
            ChargesLookup
        Else
            ItemsLookup
        End If
    End If
End Select
End Sub

Private Sub ItemInv_SelChange()
    If ItemInv.Row + 1 = ItemInv.Rows And ItemInv.Col + 1 = ItemInv.Cols Then
        ItemInv.Rows = ItemInv.Rows + 1
    End If
End Sub
Private Sub ItemInv_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Select Case ItemInv.Col
    Case 1
        If ItemInv.TextMatrix(ItemInv.Row, 0) <> "SS" Then
            itemTable.FindFirst " ITEM = " & MyParn(ItemInv.EditText)
            If itemTable.NoMatch Then
                 Cancel = True
'                     cMess = "·‰ Ì „  ”ÃÌ· «·’‰› " & TurnValue(ItemInv.exitText, Null, "") & " ﬂÊœ «·’‰› €Ì— „”Ã· "
            Else
                ItemInv.TextMatrix(ItemInv.Row, 2) = itemTable.DescA
                If ItemInv.EditText <> ItemInv.TextMatrix(ItemInv.Row, 1) Then
                    Select Case publicFlag
                        Case 1
                            ItemInv.TextMatrix(ItemInv.Row, 4) = itemTable.price
                        Case 2
                            ItemInv.TextMatrix(ItemInv.Row, 4) = itemTable.price
                        Case 3, 4
                            ItemInv.TextMatrix(ItemInv.Row, 4) = itemTable.cost
                    End Select
                    End If
                    
                    itemBalTable.FindFirst " ITEM = " & MyParn(ItemInv.EditText) & " AND STORE = " & MyParn(ItemInv.TextMatrix(ItemInv.Row, 0))
                    If itemBalTable.NoMatch Then
                        xBalItem.Caption = "·« ÌÊÃœ —’Ìœ ··’‰› ›Ï «·„Œ“‰ "
                    Else
                        If TurnValue(itemBalTable.SumIn, Null, 0) - TurnValue(itemBalTable.SUMout, Null, 0) <> 0 Then
                            xBalItem.Caption = "—’Ìœ «·’‰› ›Ï «·„Œ“‰ " & Format(TurnValue(itemBalTable.SumIn, Null, 0) - TurnValue(itemBalTable.SUMout, Null, 0), "##0.00")
                        Else
                            xBalItem.Caption = "·« ÌÊÃœ —’Ìœ ··’‰› ›Ï «·„Œ“‰ "
                        End If
                    End If
            End If
        Else
            ChargesTable.FindFirst "CODE = " & MyParn(ItemInv.EditText)
            If ChargesTable.NoMatch Then
                cMess = "·‰ Ì „  ”ÃÌ· «·„’—Ê›" & ItemInv.EditText & " ﬂÊœ «·„’—Ê› €Ì— „”Ã· "
                Cancel = True
            Else
                ItemInv.TextMatrix(ItemInv.Row, 2) = ChargesTable.DescA
            End If
        End If
End Select
End Sub
Private Sub xCode_Change()
If xCode.Text = "" Then
    xclientDescA.Caption = ""
    xClientBalance.Caption = ""
End If
ClientTable.FindFirst "Code = " & MyParn(xCode.Text)
xclientDescA.Caption = IIf(ClientTable.NoMatch, "", ClientTable.DescA)
If ClientTable.NoMatch Then Exit Sub
BalClient.FindFirst " CODE = " & MyParn(xCode.Text)
xClientBalance.Caption = IIf(BalClient.NoMatch, "", Format(TurnValue(BalClient.SUMSAL, Null, 0) - TurnValue(BalClient.SUMPAY, Null, 0), "###0.00"))
End Sub

Private Sub xCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xCode.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    If publicFlag = 1 Or publicFlag = 2 Then
        Generalarray(2) = "Select Code As ﬂÊœ , DescA As ≈”„ From File3_10"
    Else
        Generalarray(2) = "Select Code As ﬂÊœ,DescA As ≈”„ From File4_10"
    End If
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 3000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(4)
    Dim GrdArray(3)
    Set Generalarray(1) = Me
    If publicFlag = 1 Or publicFlag = 2 Then
        Generalarray(2) = "Select " & myFileName & ".Doc_No as [«·„”·”·],format([Date],'dd-mm-yyyy') as [ «—ÌŒ «·›« Ê—…],File3_10.DescA as [«”„ «·„Ê—œ] " & _
                          " From " & myFileName & " Inner join file3_10 on " & myFileName & ".Code = file3_10.code "
        Generalarray(3) = " Where File3_10.DescA Like '*cFilter*'"
        Generalarray(4) = " Group by " & myFileName & ".Doc_No,[date],File3_10.desca ORDER BY [DATE]"
    Else
        Generalarray(2) = "Select " & myFileName & ".Doc_No as [«·„”·”·],format([Date],'dd-mm-yyyy') as [ «—ÌŒ «·›« Ê—…],File4_10.DescA as [«”„ «·„Ê—œ] " & _
                          " From " & myFileName & " Inner join file4_10 on " & myFileName & ".Code = file4_10.code "
        Generalarray(3) = "Where File4_10.DescA Like '*cFilter*'"
        Generalarray(4) = " Group by " & myFileName & ".Doc_No,[date],File4_10.desca ORDER BY [DATE]"
    End If
    GrdArray(1) = 1000
    GrdArray(2) = 1500
    GrdArray(3) = 3000
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Show 1
End If
End Sub
Sub ChargesLookup()
    Dim Generalarray(3)
    Dim GrdArray(2)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code as [«·„”·”·],DescA as [«·«”„] From file8_70 "
    Generalarray(3) = "WHERE DescA Like '*cFilter*'"
    GrdArray(1) = 1000
    GrdArray(2) = 4000
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Private Function StrStore()
If storeTable.RecordCount > 0 Then
    storeTable.MoveFirst
    i = 1
    StrStore = "#" & storeTable!CODE & ";" & storeTable.DescA
    storeTable.MoveNext
    Do While True
        i = i + 1
        If storeTable.EOF Then Exit Do
        StrStore = StrStore & "|#" & storeTable!CODE & ";" & storeTable.DescA
        storeTable.MoveNext
    Loop
    StrStore = StrStore & "|#SS" & ";„’«—Ì›"
End If
End Function
Function myDelete()
    ' Õ–›  «·›« Ê—…
    cString = " DELETE  " & myFileName & " .* FROM " & myFileName & " WHERE " & myFileName & ".DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
       
    ' Õ–› Õ—ﬂ… √’‰«› «·›« Ê—…
    cString = " DELETE  FILE1_11.* FROM FILE1_11 WHERE FILE1_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = " & MyParn(itemMoveType)
    mydb.Execute cString
        
    ' Õ–› Õ—ﬂ… ⁄„Ì· «·›« Ê—…
    Select Case publicFlag
        Case 1
            cString = " DELETE  FILE3_11.* FROM FILE3_11 WHERE FILE3_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '4' "
        Case 2
            cString = " DELETE  FILE3_11.* FROM FILE3_11 WHERE FILE3_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '5' "
        Case 3
            cString = " DELETE  FILE4_11.* FROM FILE3_11 WHERE FILE4_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '4' "
        Case 4
            cString = " DELETE  FILE4_11.* FROM FILE3_11 WHERE FILE4_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '5' "
    End Select
    mydb.Execute cString
    Set itemBalTable = mydb.CreateSnapshot("SELECT ITEM, STORE, Sum(OUT) AS SumOUT, Sum([IN]) AS SumIN FROM FILE1_11 GROUP BY ITEM, STORE ORDER BY ITEM , STORE  ", dbOpenDynaset)

    Set DocTable = mydb.OpenRecordset("SELECT * FROM " & myFileName & " ORDER BY DOC_NO ,STORE , ITEM", dbOpenDynaset)
End Function
Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.Text = "" Then Cancel = True
DocValid
End Sub


