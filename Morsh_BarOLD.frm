VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form barcodefrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13185
   BeginProperty Font 
      Name            =   "Simplified Arabic"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   13185
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   915
      Left            =   6390
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   -90
      Width           =   6525
      Begin VB.CommandButton CMD_Bar56 
         Caption         =   "ÿ»«⁄… 56"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   300
         Width           =   1650
      End
      Begin VB.CommandButton CMD_Bar144 
         Caption         =   "ÿ»«⁄… 144"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   300
         Width           =   1650
      End
      Begin VB.Label xTotal144 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   3450
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   345
         Width           =   765
      End
      Begin VB.Label xTotal56 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   345
         Width           =   765
      End
   End
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   645
      Left            =   2160
      Top             =   8100
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   1138
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
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   1710
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   4200
      Begin VB.CommandButton Command6 
         Caption         =   "ÿ»«⁄…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   270
         Width           =   1650
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄œœ «·’›Õ«  :"
         Height          =   360
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   1245
      End
      Begin VB.Label xPageCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   810
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6405
      Left            =   180
      TabIndex        =   12
      Top             =   900
      Width           =   12750
      _cx             =   22490
      _cy             =   11298
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
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
   Begin VB.Frame Frame3 
      Height          =   870
      Left            =   9315
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   7695
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "Õ–› «·„ÿ»Ê⁄"
         Height          =   510
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   270
         Width           =   1230
      End
      Begin VB.CommandButton cmdDelAll 
         Caption         =   "Õ–› «·ﬂ·"
         Height          =   510
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Œ—ÊÃ"
         Height          =   510
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   5220
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   7695
      Width           =   4080
      Begin VB.TextBox xCol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox xRow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "«·⁄„Êœ :"
         Height          =   315
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "«·’›:"
         Height          =   390
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   315
         Width           =   615
      End
   End
End
Attribute VB_Name = "barcodefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NCOLS As Double
Dim NROWS As Double
Dim con As New ADODB.Connection
Private Sub cmduno_Click()
MyLoad
End Sub
Private Sub Command3_Click()
frmReturn.Show 1
MyLoad
End Sub
Private Sub CmdDelAll_Click()
If MsgBox("Õ–› ﬂ· «·”Ã·« ", vbYesNo + vbDefaultButton1) = vbYes Then
    con.BeginTrans
    con.Execute "delete  from addprint"
    con.CommitTrans
    myloadgrd
End If
End Sub

Private Sub Command1_Click()
If MsgBox("Õ–› «·„ÿ»Ê⁄ ›ﬁÿ", vbYesNo + vbDefaultButton1) = vbYes Then
    con.BeginTrans
    con.Execute "delete  from addprint where isprint = 0"
    con.CommitTrans
    myloadgrd
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
    BalBar.Show 1
    tAddPrint.Requery
    MyLoad
End Sub

Private Sub Command6_Click()
    If Val(xRow.Text) > 24 Then
        MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
        Exit Sub
    End If
    
    If Val(xCol.Text) > 6 Then
        MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
        Exit Sub
    End If
    
    doprint
    Set myForm = Me
    CardPrintNew.Show 1
    If MsgBox(" „  «·ÿ»«⁄…", vbYesNo, "«·ÿ»«⁄…") = vbYes Then
        con.BeginTrans
        con.Execute "update ADDPRINT SET ISPRINT = 0"
        con.CommitTrans
        myloadgrd
    End If
End Sub
Private Sub Form_Load()
openCon con
Set grid1.DataSource = DATA10
DATA10.ConnectionString = strCon
myloadgrd
End Sub
Sub MyLoad()
Dim nTotal144 As Double
Dim nTotal56 As Double

With grid1
.Rows = 1
tAddPrint.Requery

Do Until tAddPrint.EOF
   .AddItem ""
    ItemTable.Find "item = " & MyParn(tAddPrint!Item & ""), , adSearchForward, adBookmarkFirst
    If Not ItemTable.EOF Then
        .TextMatrix(grid1.Rows - 1, 0) = TurnValue(tAddPrint!Item, Null, "")
        If Val(ItemTable!PRICE & "") = O Then MsgBox " ·« ÌÊÃœ ”⁄— „” Â·ﬂ  " & ItemTable!desca & " ===>  " & ItemTable!Item
        .TextMatrix(grid1.Rows - 1, 1) = ItemTable!desca & ""
        .TextMatrix(grid1.Rows - 1, 3) = tAddPrint!Quant & ""
        .TextMatrix(grid1.Rows - 1, 2) = ItemTable!PRICE & ""
        .TextMatrix(grid1.Rows - 1, 4) = IIf(tAddPrint!isPrint, "-1", "0")
        .TextMatrix(grid1.Rows - 1, 5) = tAddPrint!T_BAR & ""
        .TextMatrix(grid1.Rows - 1, 6) = tAddPrint!doc_no & ""
        If .TextMatrix(grid1.Rows - 1, 5) = "1" Then
            nTotal56 = nTotal56 + Val(.TextMatrix(grid1.Rows - 1, 3))
        Else
            nTotal144 = nTotal144 + Val(.TextMatrix(grid1.Rows - 1, 3))
        End If
    End If
    tAddPrint.MoveNext
Loop
.AddItem ""
End With
xTotal144.Caption = nTotal144
xTotal56.Caption = nTotal56
xTot144.Caption = nTotal144
xTot56.Caption = nTotal56

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Search3
Unload Me
Set Morsh_Bar = Nothing
End Sub
Private Sub Grid1_EnterCell()
'If grid1.col = 0 Or grid1.col = 2 Or grid1.col = 4 Then
'    grid1.Editable = flexEDKbdMouse
'Else
'    grid1.Editable = flexEDNone
'End If
 grid1.Editable = flexEDKbdMouse
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        grid1.RemoveItem grid1.Row
    End If
End If
If KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, Search3
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    'ItemsLookup
End If
End Sub
Private Function MYVALID() As Boolean
With grid1
For i = 1 To grid1.Rows - 2
    If Val(.TextMatrix(i, 3)) = 0 Then
        MsgBox "«·ﬂ„Ì… €Ì— „”Ã·…"
        Exit Function
    End If
Next
MYVALID = True
End With
End Function
Private Function myvalidRowCol() As Boolean
If Val(xRow.Text) > NROWS Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Function
End If

If Val(xCol.Text) > NCOLS Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Function
End If
myvalidRowCol = True
End Function
Private Sub doprint()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim nCost As Double
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(0.2)
SettingArray(cLeftMargin) = MyMeasure(0.3)
SettingArray(cCardWidth) = MyMeasure(3.48)
SettingArray(cCardHeight) = MyMeasure(1.73)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 17
SettingArray(cCols) = 6
SettingArray(cPageWidth) = MyMeasure(21)

contemp.Execute "delete * From Card"
temptable.Open "Select * From card", contemp, adOpenKeyset, adLockOptimistic, adCmdText

'cString = "Select ADDPRINT.ITEM,FILE1_10.DESCA,FILE1_10.PRICE3,SUM(ADDPRINT.QUANT) AS sumofQuant" & _
'          " FROM ADDPRINT INNER JOIN FILE1_10 ON ADDPRINT.ITEM = FILE1_10.ITEM"
'cString = cString & turn(cString) & " ADDPRINT.ISPRINT = 1"
'cString = cString & " GROUP BY ADDPRINT.ITEM,FILE1_10.DESCA,FILE1_10.PRICE3"
'cString = cString & " ORDER BY ADDPRINT.ITEM"

cString = "Select ADDPRINT.ITEM,ADDPRINT.DESCA,FILE1_10.PRICE3,SUM(ADDPRINT.QUANT) AS sumofQuant" & _
          " FROM ADDPRINT LEFT JOIN FILE1_10 ON ADDPRINT.ITEM = FILE1_10.ITEM"
cString = cString & turn(cString) & " ADDPRINT.ISPRINT = 1"
cString = cString & " GROUP BY ADDPRINT.ITEM,ADDPRINT.DESCA,FILE1_10.PRICE3"
cString = cString & " ORDER BY ADDPRINT.ITEM"


sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
            

With temptable
nCard = 0
nRow = 0
nCard = 0
nCol = 0
NCOLS = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * NCOLS) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = NCOLS, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next
'«‰ Â«¡

Do Until sourcetable.EOF
'************
    For i = 1 To Val(sourcetable!sumOfQuant & "")
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = (nCol = NCOLS)
        nFirst = MyMeasure(0.13)
        nHeight = 0
        nLast = MyMeasure(0)
        nLastCol = MyMeasure(0)
        For nCount = 1 To 1
            '≈”„ «·„Õ·
            temptable.AddNew
            temptable!Top = MyMeasure(0) + IIf(nRow = 1, nFirst, 0)
            temptable!Left = MyMeasure(0.4)
            temptable!Width = SettingArray(cCardWidth) - (Val(temptable!Left & "") * 2)
            temptable!Height = MyMeasure(0)
            temptable!FontName = "Shurooq 07"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!fontsize = 8
            temptable!TextAlign = taCenterTop
            temptable!Text = sourcetable!Item
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
                        
' BARCODE
            temptable.AddNew
           'temptable!Top = MyMeasure(0.3) + nHeight + IIf(nRow = 1, nFirst, 0)
            temptable!Top = MyMeasure(0.4) + nHeight + IIf(nRow = 1, nFirst, 0)
            temptable!Left = MyMeasure(0.3)
            temptable!Width = SettingArray(cCardWidth) - (Val(temptable!Left & "") * 2)
           ' temptable!Height = MyMeasure(0.4)
            temptable!Height = MyMeasure(0.6)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!isBarcode = True
            temptable!fontsize = 8
            temptable!TextAlign = taLeftTop
            temptable!Text = sourcetable!Item
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
                        
' DESCA
            temptable.AddNew
            temptable!Top = MyMeasure(0.7) + nHeight + IIf(nRow = 1, nFirst, 0)
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = SettingArray(cCardWidth) - (Val(temptable!Left & "") * 2)
            temptable!Height = MyMeasure(0)
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!fontsize = 8
            temptable!TextAlign = taCenterBottom
            'temptable!Text = sourcetable!Desca
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update


' PRICE
            temptable.AddNew
            temptable!Top = MyMeasure(1) + nHeight + IIf(nRow = 1, nFirst, 0)
            temptable!Left = MyMeasure(0.4) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(0)
            temptable!Height = MyMeasure(0)
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!fontsize = 8
            temptable!TextAlign = taLeftTop
            temptable!Text = turn(Myvalue(sourcetable!price3), "L.E. ") & Format(sourcetable!price3, "Fixed")
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
'            nHeight = SettingArray(cCardHeight)
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If Not validRows(grid1.Row) And grid1.Row <> grid1.Rows - 1 Then grid1.RemoveItem grid1.Row
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
    If .Col = 0 Then
        If Trim(.EditText) = "" Then
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Cancel = True
        Else
'            If GetDesca("Select item from file1_10 where ITEM = " & MyParn(.EditText)) = "" Then
'                MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
'                Cancel = True
'            End If
        End If
    End If
    
End With
End Sub

Private Sub PRINT_BAR_Click()
MyBar1
    Set myForm = Me
    CardPrint_1.Show 1

End Sub

Private Sub xTotal56_Click()
Dim nTotal144 As Double
Dim nTotal56 As Double
With grid1
    For i = 1 To .Rows - 1
        If TurnValue(.TextMatrix(i, 4), "", False) Then
            If .TextMatrix(i, 5) = "1" Then
                nTotal56 = nTotal56 + Val(.TextMatrix(i, 3))
            Else
                nTotal144 = nTotal144 + Val(.TextMatrix(i, 3))
            End If
        End If
    Next i
End With
xTotal144.Caption = nTotal144
xTotal56.Caption = nTotal56

xTot144.Caption = nTotal144
xTot56.Caption = nTotal56

End Sub
Private Sub xTotal144_Click()
Dim nTotal144 As Double
Dim nTotal56 As Double
With grid1
    For i = 1 To .Rows - 1
        If TurnValue(.TextMatrix(i, 4), "", False) Then
            If .TextMatrix(i, 5) = "1" Then
                nTotal56 = nTotal56 + Val(.TextMatrix(i, 3))
            Else
                nTotal144 = nTotal144 + Val(.TextMatrix(i, 3))
            End If
        End If
    Next i
End With
xTotal144.Caption = nTotal144
xTotal56.Caption = nTotal56

xTot144.Caption = nTotal144
xTot56.Caption = nTotal56

End Sub
Private Sub DoprintBar1()
Dim tCard As New ADODB.Recordset
Dim tPrint As New ADODB.Recordset
Dim nCost As Double
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(0.1)
SettingArray(cLeftMargin) = MyMeasure(0.25)
SettingArray(cCardWidth) = MyMeasure(5)
SettingArray(cCardHeight) = MyMeasure(2.5)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 1
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(21)

con.Execute "delete * From Card"
tCard.Open "Select * From card", con, adOpenKeyset, adLockOptimistic, adCmdText
tPrint.Open "Select addPrint.isPRICE ,addPrint.isPRICE2,addPrint.ishead , file1_10.ITEM,file1_10.DESCA,FILE1_10.PRICE ,FILE1_10.PRICE2 ,FILE1_10.COST , FILE1_10.Package ,addPrint.Quant , addPrint.doc_no From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item Where addPrint.isPrint  AND addPrint.TYPE_BARCODE = 1 ORDER BY FILE1_10.ITEM ", con, adOpenKeyset, adLockOptimistic, adCmdText

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
NCOLS = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * NCOLS) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = NCOLS, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡
If tPrint.RecordCount = 0 Then Exit Sub
Do
'************
    For i = 1 To tPrint!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = (nCol = NCOLS)
        
        nHeight = 0
        nLast = MyMeasure(0)
        nLastCol = MyMeasure(0)
        For nCount = 1 To 1
            
            If Not tPrint!ishead Then
            '≈”„ «·„Õ·
            tCard.AddNew
            tCard!Top = MyMeasure(0.25) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(1.7) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(3)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "Shurooq 07"
            tCard!FontBold = True
            tCard!FontUnderline = False
            tCard!fontsize = 16
            tCard!TextAlign = taCenterTop
            tCard!Text = "„ﬂ »… Œ«·œ «·„Â‰œ”"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            ' ·Ì›Ê‰
            tCard.AddNew
            tCard!Top = MyMeasure(0.3) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!fontsize = 7
            tCard!TextAlign = taLeftTop
            tCard!Text = "Tel." & cHPhone1
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Top = MyMeasure(0.55) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!fontsize = 7
            tCard!TextAlign = taLeftTop
            tCard!Text = cHPhone2
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            Else
                tCard.AddNew
                tCard!Top = MyMeasure(0.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
                tCard!Left = MyMeasure(2) - IIf(blastcol, nLastCol, 0)
                tCard!Width = MyMeasure(3)
                tCard!Height = MyMeasure(0)
                tCard!FontName = "Shurooq 07"
                tCard!FontName = "Arabic Transparent"
                tCard!FontBold = True
                tCard!FontUnderline = False
                tCard!fontsize = 7
                tCard!Text = "€Ì— „”„ÊÕ »≈— Ã«⁄ «·ﬂ »"
                tCard!ForeColor = vbBlack
                tCard!CardNo = nCard
                tCard.Update
    
            End If
' BARCODE
            tCard.AddNew
            tCard!Top = MyMeasure(0.85) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(3.3)
            tCard!Height = MyMeasure(0.45)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!isBarcode = True
            tCard!fontsize = 8
            tCard!TextAlign = taLeftTop
            tCard!Text = tPrint!Item
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Top = MyMeasure(1.3) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(3.3)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = False
            
            tCard!fontsize = 8
            tCard!TextAlign = taCenterTop
            tCard!Text = tPrint!Item
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            

' ITEM
            tCard.AddNew
            tCard!Top = MyMeasure(0.85) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(3.6) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(1)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!fontsize = 8
            tCard!TextAlign = taRightTop
            If Len(tPrint!Item) < 10 Then
'               tCard!Text = tPrint!Item
            Else
'                tCard!Text = "/" & DelZero(tPrint!code) & "/"
            End If
            tCard!FontUnderline = True

            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            
            
' DESCA
            tCard.AddNew
            tCard!Top = MyMeasure(1.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(4.8)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!fontsize = 9
            tCard!TextAlign = taCenterBottom
            tCard!Text = tPrint!desca
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update


' PRICE
            tCard.AddNew
            tCard!Top = MyMeasure(2) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.4) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!fontsize = 10
            tCard!TextAlign = taLeftTop
            If tPrint!isPRICE Then
                tCard!Text = "L.E. " & Format(tPrint!price2, "Fixed")
                tCard!FontUnderline = True
            Else
                If Not tPrint!isPRICE2 Then
                    tCard!Text = "L.E. " & Format(tPrint!PRICE, "Fixed")
                End If
            End If
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
' COST
            If Val(tPrint!cost & "") > 0 Then
                tCard.AddNew
                tCard!Top = MyMeasure(2) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
                tCard!Left = MyMeasure(3.6) - IIf(blastcol, nLastCol, 0)
                tCard!Width = MyMeasure(1)
                tCard!Height = MyMeasure(0)
                tCard!FontName = "arial"
                tCard!FontBold = False
                tCard!fontsize = 8
                tCard!TextAlign = taRightTop
                
                nCost = Val(LastPrice(tPrint!Item) & "")
                If nCost = 0 Then nCost = tPrint!cost
                If tPrint!isPRICE Then
                    nCost = (Val(tPrint!price2 & "") - Val(tPrint!cost & "")) * 100 / 2
                Else
                    If Val(tPrint!package & "") <> 0 Then nCost = (Val(tPrint!PRICE & "") - (nCost / Val(tPrint!package & ""))) / 2 * 100
                End If
                If IsNull(tPrint!doc_no) Then
                    tCard!Text = Format(nCost, "#0")
                Else
                    tCard!Text = Format(nCost, "#0") & "/" & DelZero(tPrint!doc_no)
                End If
                tCard!ForeColor = vbBlack
                tCard!CardNo = nCard
                tCard.Update
            End If
            nHeight = SettingArray(cCardHeight)
        Next
' ----------------
    Next
    tPrint.MoveNext
Loop Until tPrint.EOF
End With
End Sub
Private Sub DoprintBar2()
Dim tCard As New ADODB.Recordset
Dim tPrint As New ADODB.Recordset
Dim nCost As Double
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(0.1)
SettingArray(cLeftMargin) = MyMeasure(0)
SettingArray(cCardWidth) = MyMeasure(3.8)
SettingArray(cCardHeight) = MyMeasure(1.25)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 2
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(21)

con.Execute "delete * From Card"
tCard.Open "Select * From card", con, adOpenKeyset, adLockOptimistic, adCmdText
tPrint.Open "Select addPrint.isPRICE , addPrint.isPRICE2, addPrint.ishead , file1_10.ITEM,file1_10.DESCA,FILE1_10.PRICE,FILE1_10.PRICE2 ,FILE1_10.COST , FILE1_10.Package ,addPrint.Quant,addPrint.doc_no From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item Where addPrint.isPrint  AND addPrint.TYPE_BARCODE = 2 ORDER BY FILE1_10.ITEM ", con, adOpenKeyset, adLockOptimistic, adCmdText

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
NCOLS = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * NCOLS) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = NCOLS, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡
If tPrint.RecordCount = 0 Then Exit Sub
Do
'************
    For i = 1 To tPrint!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = (nCol = NCOLS)
        
        nHeight = 0
        nLast = MyMeasure(0)
        nLastCol = MyMeasure(0)
        For nCount = 1 To 1
            
            If Not tPrint!ishead Then
            
            '≈”„ «·„Õ·
            tCard.AddNew
            tCard!Top = MyMeasure(0.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(2) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(1.8)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "Arabic Transparent"
'           tCard!FontName = "Shurooq 07"
            tCard!FontBold = True
            tCard!FontUnderline = False
            tCard!fontsize = 7
            tCard!TextAlign = taCenterTop
            tCard!Text = "„ﬂ »… Œ«·œ «·„Â‰œ”"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            ' ·Ì›Ê‰
            tCard.AddNew
            tCard!Top = MyMeasure(0.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.5) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!fontsize = 7
            tCard!TextAlign = taLeftTop
            tCard!Text = "Tel." & cHPhone1
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            
        Else
            tCard.AddNew
            tCard!Top = MyMeasure(0.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.5) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(3)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "Arabic Transparent"
            tCard!FontBold = False
            tCard!FontUnderline = False
            tCard!fontsize = 7
            tCard!TextAlign = taCenterTop
            tCard!Text = "€Ì— „”„ÊÕ »≈— Ã«⁄ «·ﬂ »"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update

        End If

' BARCODE
            tCard.AddNew
            tCard!Top = MyMeasure(0.38) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.5) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(2.8)
            tCard!Height = MyMeasure(0.3)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!isBarcode = True
            tCard!fontsize = 8
            tCard!TextAlign = taLeftTop
            tCard!Text = tPrint!Item
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update

' ITEM
            tCard.AddNew
            tCard!Top = MyMeasure(0.4) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0.8)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!FontUnderline = True
            tCard!fontsize = 7
            tCard!TextAlign = taRightTop
                
            If Len(tPrint!Item) < 10 Then
                tCard!Text = tPrint!Item
            Else
'                tCard!Text = "/" & DelZero(tPrint!code) & "/"
            End If
                
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
' DESCA
            tCard.AddNew
            tCard!Top = MyMeasure(0.65) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(3.5)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!fontsize = 7
            tCard!TextAlign = taCenterBottom
            tCard!Text = tPrint!desca
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update

' PRICE
            tCard.AddNew
            tCard!Top = MyMeasure(0.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.6) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!fontsize = 8
            tCard!TextAlign = taLeftTop
            If tPrint!isPRICE Then
                tCard!Text = "L.E. " & Format(tPrint!price2, "Fixed")
                tCard!FontUnderline = True
            Else
                If Not tPrint!isPRICE2 Then
                    tCard!Text = "L.E. " & Format(tPrint!PRICE, "Fixed")
                End If
            End If
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
' COST
            If Val(tPrint!cost & "") > 0 Then
                tCard.AddNew
                tCard!Top = MyMeasure(0.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
                tCard!Left = MyMeasure(2.8) - IIf(blastcol, nLastCol, 0)
                tCard!Width = MyMeasure(1)
                tCard!Height = MyMeasure(0)
                tCard!FontName = "arial"
                tCard!FontBold = False
                tCard!fontsize = 7
                tCard!TextAlign = taRightTop
                
                nCost = Val(LastPrice(tPrint!Item) & "")
                If nCost = 0 Then nCost = tPrint!cost
                If tPrint!isPRICE Then
                    nCost = (Val(tPrint!price2 & "") - Val(tPrint!cost & "")) * 100 / 2
                Else
                    If Val(tPrint!package & "") <> 0 Then nCost = (Val(tPrint!PRICE & "") - (nCost / Val(tPrint!package & ""))) / 2 * 100
                End If
                
                If IsNull(tPrint!doc_no) Then
                    tCard!Text = Format(nCost, "#0")
                Else
                    tCard!Text = Format(nCost, "#0") & "/" & DelZero(tPrint!doc_no)
                End If
                tCard!ForeColor = vbBlack
                tCard!CardNo = nCard
                tCard.Update
            End If
            nHeight = SettingArray(cCardHeight)
        Next
' ----------------
    Next
    tPrint.MoveNext
Loop Until tPrint.EOF
End With
End Sub
Private Sub MyBar1()
Dim tCard As New ADODB.Recordset
Dim tPrint As New ADODB.Recordset
Dim nCost As Double


con.Execute "delete * From Card"
tCard.Open "Select * From card", con, adOpenKeyset, adLockOptimistic, adCmdText



nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(0.1)
SettingArray(cLeftMargin) = MyMeasure(0.25)
SettingArray(cCardWidth) = MyMeasure(5)
SettingArray(cCardHeight) = MyMeasure(2.5)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 1
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(21)

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
NCOLS = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * NCOLS) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
'«‰ Â«¡
'************
    For i = 1 To 1
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = (nCol = NCOLS)
        
        nHeight = 0
        nLast = MyMeasure(0)
        nLastCol = MyMeasure(0)
        For nCount = 1 To 1
            
' BARCODE
            tCard.AddNew
            tCard!Top = MyMeasure(0.3) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.5) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(4)
            tCard!Height = MyMeasure(1.8)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!isBarcode = True
            tCard!fontsize = 8
            tCard!TextAlign = taLeftTop
            tCard!Text = InputBox(" ", "—ﬁ„")
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            
            nHeight = SettingArray(cCardHeight)
        Next
' ----------------
    Next
End With
End Sub
Private Sub FixGrd()
With grid1
    .Cols = 6
    .WordWrap = True
    .TextMatrix(0, 0) = "—ﬁ„ «·’‰› "
    .TextMatrix(0, 1) = "«·’‰‹‹‹‹‹›"
    .TextMatrix(0, 2) = "«·ﬂ„Ì… "
    .TextMatrix(0, 3) = "«·”⁄—"
    .TextMatrix(0, 4) = "ÿ»«⁄…"
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 4000
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1000
    .ColHidden(.Cols - 1) = True
    .RowHeight(0) = 800
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignRightCenter
End With
End Sub
Private Sub myloadgrd()
With grid1
'    cString = "Select addPrint.Item,file1_10.DescA,addPrint.Quant,file1_10.Price3,addprint.isPrint,addPrint.id" & _
'              " From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item "
    cString = "Select addPrint.Item,ADDPRINT.DescA,addPrint.Quant,file1_10.Price3,addprint.isPrint,addPrint.id" & _
              " From addPrint LEFT join File1_10 on AddPrint.Item = File1_10.item "
    
    cString = cString & " order by ADDPRINT.ID"
    DATA10.RecordSource = cString
    DATA10.Refresh
    grid1.AddItem ""
End With
'Handlecontrols LoadMode
FixGrd
CalcTotals
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRows(Row) Then Exit Sub

If Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    grid1.TextMatrix(Row, 4) = -1
End If
If Col = 0 Then calcRow Row

CalcTotals
myreplaceGrdRow (Row)

End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 Then
    If Not validRows(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub
Private Function myreplaceGrdRow(i) As Boolean
Dim aInsert(3, 1)
With grid1

con.BeginTrans

aInsert(0, 0) = "item"
aInsert(0, 1) = addstring(grid1.TextMatrix(i, 0))

aInsert(1, 0) = "quant"
aInsert(1, 1) = Val(.TextMatrix(i, 2))

aInsert(2, 0) = "DESCA"
aInsert(2, 1) = addstring(.TextMatrix(i, 1))

aInsert(3, 0) = "isprint"
aInsert(3, 1) = IIf(Val(.TextMatrix(i, 4)) = 0, "0", "1")

If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
    con.Execute CreateInsert(aInsert, "ADDPRINT")
Else
    con.Execute CreateUpdate(aInsert, "ADDPRINT", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
End If
End With
con.CommitTrans

If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then myloadgrd

myreplaceGrdRow = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
myloadgrd
End Function
Private Function validRows(nRow) As Boolean
If Trim(grid1.TextMatrix(nRow, 0)) = "" Then
    Exit Function
End If
validRows = True
End Function
Private Sub calcRow(nRow)
With grid1
'.TextMatrix(nRow, 1) = ""
'.TextMatrix(nRow, 3) = ""

If Trim(grid1.TextMatrix(nRow, 0)) = "" Then Exit Sub

Dim aret As Variant
aret = aGetDesca("Select Desca,price3 from file1_10 where file1_10.item = " & MyParn(grid1.TextMatrix(nRow, 0)))
If UBound(aret) > 0 Then
    .TextMatrix(nRow, 1) = aret(1) & ""
    .TextMatrix(nRow, 3) = aret(2) & ""
End If
End With
End Sub
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    cItem = grid1.TextMatrix(grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 0) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = "1"
    If Trim(cItem) <> Trim(grid1.TextMatrix(grid1.Row, 0)) Then calcRow grid1.Row
    If grid1.Row = grid1.Rows - 1 Then
        grid1.Select grid1.Rows - 1, 0
        grid1_AfterEdit grid1.Row, grid1.Col
        grid1.Row = grid1.Rows - 1
        grid1.Col = 0
    Else
        grid1_AfterEdit grid1.Row, grid1.Col
    End If
End If
End Sub
Private Sub CalcTotals()
Dim nCount As Integer
For i = 1 To grid1.Rows - 1
    nCount = nCount + Val(grid1.TextMatrix(i, 2))
Next
xPageCount.Caption = Fix(nCount / 102) + IIf(nCount Mod 102 <> 0, 1, 0)

End Sub
