VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PurPayfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "”œ«œ «·›« Ê—…"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2625
      Left            =   90
      TabIndex        =   14
      Top             =   90
      Width           =   6000
      _cx             =   10583
      _cy             =   4630
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame3 
      Caption         =   "≈Ã„«·Ì"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   4860
      Width           =   4920
      Begin VB.Label xRestCur 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1530
         Width           =   2130
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " ’›Ì… „⁄ «· ⁄œÌ· :"
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
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1620
         Width           =   1425
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„ »ﬁÌ :"
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1215
         Width           =   600
      End
      Begin VB.Label xRest 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   2130
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈Ã„«·Ì «·›« Ê—… :"
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   2130
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„œ›Ê⁄ :"
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   765
         Width           =   705
      End
      Begin VB.Label xPaid 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   2130
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2745
      Width           =   3750
      Begin VB.CommandButton Command1 
         Caption         =   " ⁄œÌ·"
         Height          =   420
         Left            =   1215
         RightToLeft     =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Õ›Ÿ"
         Height          =   420
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
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
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
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
   Begin VB.Label lblRecord 
      Height          =   375
      Left            =   990
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   6795
      Width           =   2850
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1215
      Width           =   45
   End
End
Attribute VB_Name = "PurPayfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CARDTABLE As New ADODB.Recordset
Dim defbox As String
Private Sub CmdApply_Click()
'con.Execute "Delete * from file8_10h where doc_no = " & MyParn(cDoc_No)
'con.Execute "Delete * from file8_10 where doc_no = " & MyParn(cDoc_No)
If MyReplace Then
    MsgBox " „ «·Õ›Ÿ"
    'Doprint
End If
Unload Me

End Sub

Private Sub cmdAdd_Click()
If Val(xValue.Text) <> Val(xRest.Caption) Then
    If MsgBox("ﬁÌ„… «·›« Ê—… Ê«·”œ«œ €Ì— „ ”«ÊÌ…  ø", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
End If
'If xdoc_no.Caption <> "" And (Val(xvalue.Text) <> Val(xRest.Caption)) Then
'    If MsgBox("Â‰«ﬂ „” ‰œ ”œ«œ »«·›⁄· ··›« Ê—… .. «÷«›… „” ‰œ «Œ— ø", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
'End If
If MyReplace Then
    CARDTABLE.Requery
    CARDTABLE.MoveLast
    MyLoad
    xValue.Text = xRest.Caption
End If
End Sub

Private Sub cmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    If myDel Then
        CARDTABLE.Requery
        CARDTABLE.Find "doc_no < " & MyParn(xDoc_No.Caption), , adSearchBackward, adBookmarkLast
        If CARDTABLE.BOF And Not (CARDTABLE.EOF) Then CARDTABLE.MoveFirst
        MyLoad
        xValue.Text = xRest.Caption
    End If
End If
End Sub

Private Sub cmdEdit_Click()
If Val(xValue.Text) - Val(xDoc_Value.Caption) <> Val(xRest.Caption) Then
    If MsgBox("ﬁÌ„… «·›« Ê—… Ê«·”œ«œ €Ì— „ ”«ÊÌ…  ø", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
End If
If MyEdit Then
    MyLoad
    xValue.Text = xRest.Caption
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub CmdUndo_Click()
If MyReplace Then
    MsgBox " „ «·Õ›Ÿ"
    'doprint
End If
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
    
MYVALID = True
End Function
Private Sub doprint4()
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
Dim aHeader(3)

contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT Sum(Val(FILE7_20.[QUANT] & '')) AS SumOfQuant, Sum(Val(FILE7_20.[QUANT] & '')* VAL(FILE1_10.PRICE & '')) AS SumOfValue, FILE1_10.GROUP AS GroupCode, FILE1_50.DESCA AS GroupDesca, FILE1_50.group AS MainGroupCode, FILE1_51.DESCA AS MainGroupDesca, FILE1_13.DESCA AS UNITDESCA,FILE7_20.DOC_NO,FILE7_21.DATE,FILE7_21.CODE,FILE7_21.DESCA AS SUPDESCA " & _
          " FROM (((((FILE7_20 INNER JOIN FILE7_21 ON FILE7_20.DOC_NO = FILE7_21.DOC_NO) INNER JOIN FILE4_10 ON FILE7_21.CODE = FILE4_10.CODE)INNER JOIN FILE1_10 ON FILE7_20.ITEM = FILE1_10.ITEM) LEFT JOIN FILE1_13 ON FILE1_10.UNIT = FILE1_13.CODE) LEFT JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_51 ON FILE1_50.GROUP = FILE1_51.CODE"

If xCode.Text <> "" Then
    cString = cString & TurnFound(cString) & "File7_21.CODE = " & MyParn(xCode.Text)
    aHeader(0) = "[" & "«·„Ê—œ : " & xCodeDesca.Caption & "]"
End If

If xDoc_No.Caption <> "" Then
    cString = cString & TurnFound(cString) & "File7_21.doc_no = " & MyParn(xDoc_No.Caption)
    aHeader(1) = "[" & "«·„Ê—œ : " & xCodeDesca.Caption & "]"
End If

If xdate1.Text <> "" Then
    cString = cString & TurnFound(cString) & "File7_21.date >= " & DateSql(xdate1.Text)
    aHeader(2) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If xDate2.Text <> "" Then
    cString = cString & TurnFound(cString) & "File7_21.date <= " & DateSql(xDate2.Text)
    aHeader(2) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

cString = "Group by FILE1_10.GROUP, FILE1_50.DESCA, FILE1_50.group , FILE1_51.DESCA, FILE1_13.DESCA AS UNITDESCA,FILE7_20.DOC_NO,FILE7_21.DATE,FILE7_21.CODE,FILE7_21.DESCA AS SUPDESCA "

Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, CON, adOpenForwardOnly, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        cCondition = IIf(xShowMinus.Value = 0, sourcetable!BALANCE > 0, sourcetable!BALANCE < 0)
        If cCondition Then
            temptable.AddNew
            temptable!str6 = !mainGroupDesCa
            temptable!str5 = !MAINGROUPCODE
            temptable!str1 = !GroupCode
            temptable!str2 = !GroupDesca
            temptable!str8 = !unitDesca
            temptable!VAL1 = Val(!BALANCE & "")
            temptable!VAL2 = Val(!BalanceValue & "")
            temptable!str7 = "√—’œ… «·√’‰«›"
            temptable.Update
        End If
      .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    If xShowCost.Value = 0 Then
       Report1.ReportFileName = App.Path & "\Reports\Item1.rpt"
    Else
        Report1.ReportFileName = App.Path & "\Reports\item1_2.rpt"
    End If
    
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    Report1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub Form_Load()
CARDTABLE.Open "select file8_20H.doc_no,file8_20h.date,file8_20.value from file8_20 inner join file8_20h on file8_20.doc_no = file8_20h.doc_no where file8_20h.doc_no_pur = " & MyParn(Purchasefrm.xDoc_No.Text) & " Order by file8_20h.Doc_no Desc", CON, adOpenKeyset, adLockReadOnly, adCmdText
data1.ConnectionString = CON.ConnectionString
data1.RecordSource = "Select * From file0_50"

With grid1
    .Cols = 5
    .Rows = 2
    .TextMatrix(1, 0) = True
    .Editable = flexEDKbdMouse
    .FormatString = "—ﬁ„ «·„” ‰œ|" & " «—ÌŒ «·„” ‰œ|" & "Œ“‰…|" & "«·ﬁÌ„…"
    .ColWidth(0) = 1000
    .ColWidth(1) = 1100
    .ColWidth(2) = 1100
    .ColWidth(3) = 1500
    .ColWidth(4) = 1000
    For i = 1 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColComboList(0) = StrBox
End With
defbox = RetDefBox


If Not (data1.Recordset.EOF And data1.Recordset.BOF) Then
    data1.Recordset.MoveFirst
    XBOX.BoundText = data1.Recordset!CODE
End If
xValue.Text = Purchasefrm.xTotal.Caption
xDate.Text = Purchasefrm.xDate.Text
MyLoad
xValue.Text = xRest.Caption
End Sub
Private Function MyReplace() As Boolean
cDoc_No = RetZero(Val(GetDesca("Select Max(doc_no) from file8_20h")) + 1)
On Error Resume Next
For i = 1 To 10
    CON.BeginTrans
    CON.Execute "insert into file8_20h(doc_no,[date],Doc_No_Pur)" & _
              " Values(" & _
              addstring(cDoc_No) & "," & _
              DateSq(xDate.Text) & "," & _
              addstring(Purchasefrm.xDoc_No.Text) & _
              ")"
    If Err.Number = 0 Then
        CON.Execute "Insert Into file8_20(Doc_No,[Date],Code,Desca,[Value],Box,Row,username) " & _
                  " Values(" & _
                  addstring(cDoc_No) & "," & _
                  DateSq(xDate.Text) & "," & _
                  addstring(Purchasefrm.xCode.Text) & "," & _
                  addstring("”œ«œ ›« Ê—… „‘ —Ì«  —ﬁ„ " & Format(Purchasefrm.xDoc_No.Text) & " » «—ÌŒ : " & Purchasefrm.xDate.Text) & "," & _
                  Val(xValue.Text) & "," & _
                  addstring(XBOX.BoundText) & "," & _
                  i & "," & _
                  addstring(sUserName) & _
                  ")"
        If Err.Number <> 0 Then GoTo Myerror
    End If
    If Err.Number = 0 Then Exit For
    If Err.Number = -2147467259 Then
        cDoc_No = RetZero(Val(cDoc_No) + 1)
        Err.Clear
        CON.RollbackTrans
    End If
    If Err.Number <> 0 Then GoTo Myerror
Next
CON.CommitTrans
MyReplace = True
Exit Function
Myerror:
MsgBox Err.Description
CON.RollbackTrans
Err.Clear
End Function
Private Function MyEdit() As Boolean
'On Error GoTo MYERROR
CON.BeginTrans
CON.Execute "update file8_20 SET FILE8_20.VALUE = " & Val(xValue.Text) & "," & _
            " file8_20.box = " & addstring(XBOX.BoundText) & _
            " WHERE DOC_NO = " & MyParn(xDoc_No.Caption)
CON.Execute "UPDATE file8_20h SET FILE8_20H.DATE = " & DateSq(xDate.Text) & _
            " where file8_20h.doc_no = " & MyParn(xDoc_No.Caption)
CON.CommitTrans
MyEdit = True
Exit Function
Myerror:
MsgBox Err.Description
CON.RollbackTrans
Err.Clear
End Function
Private Sub MyLoad()
CARDTABLE.Requery
With grid1
.Rows = 1
Do Until listTable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = CARDTABLE!DOC_NO
    .TextMatrix(.Rows - 1, 1) = CARDTABLE!Date
    .TextMatrix(.Rows - 1, 2) = CARDTABLE!BOX
    .TextMatrix(.Rows - 1, 3) = CARDTBLE!Value
    listTable.MoveNext
Loop
.AddItem ""
End With
CalcTotals

xTotal.Caption = Format(Val(Purchasefrm.xTotal.Caption))
'xPaid.Caption = Format(Val(GetDesca("Select Sum(value) From file8_20 inner join file8_20h on file8_20.doc_no = file8_20h.doc_no where file8_20h.doc_no_pur = " & MyParn(Purchasefrm.xDoc_No.Text))), "Fixed")
'xRest.Caption = Format(Val(Purchasefrm.xTotal.Caption) - Val(GetDesca("Select Sum(value) From file8_20 inner join file8_20h on file8_20.doc_no = file8_20h.doc_no where file8_20h.doc_no_pur = " & MyParn(Purchasefrm.xDoc_No.Text))) & "", "Fixed")
'xRestCur.Caption = Format(Val(xRest.Caption) + Val(xDoc_Value.Caption), "Fixed")
'xRest.ForeColor = IIf(Val(xRest.Caption) = 0, vbBlack, vbRed)
'lblPaid.Visible = Val(xRest.Caption) <= 0
End Sub
Private Function myDel() As Boolean
CON.BeginTrans
CON.Execute "Delete * from file8_20 where doc_no = " & MyParn(xDoc_No.Caption)
CON.Execute "Delete * from file8_20h where doc_no = " & MyParn(xDoc_No.Caption)
CON.CommitTrans
myDel = True
Exit Function
Myerror:
MsgBox Err.Description
CON.RollbackTrans
Err.Clear
End Function
Private Sub Form_Unload(Cancel As Integer)
CARDTABLE.Close
Set CARDTABLE = Nothing
Unload Me
End Sub

Private Sub xDate_Change()
Handlecontrols
End Sub

Private Sub xRest_Click()
xValue.Text = xRest.Caption
End Sub

Private Sub xRestCur_Click()
xValue.Text = xRestCur.Caption
End Sub
Private Sub xvalue_Change()
Handlecontrols
'xRest.Caption = Format(Val(xRest.Caption) - Val(xvalue.Text), "Fixed")
End Sub
Private Sub CmdFirst_Click()
CARDTABLE.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
End Sub
Private Sub CmdLast_Click()
CARDTABLE.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
CARDTABLE.MoveNext
If CARDTABLE.EOF Then
    CARDTABLE.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
CARDTABLE.MovePrevious
If CARDTABLE.BOF Then
    CARDTABLE.MoveNext
Else
    MyLoad
End If
End Sub
Private Sub Handlecontrols()
cmdAdd.Enabled = IsDate(xDate.Text) And (Val(xRest.Caption) > 0 Or Val(xPaid.Caption) = 0) And Val(xValue.Text) > 0 And XBOX.BoundText <> ""
cmdEdit.Enabled = IsDate(xDate.Text) And Val(xValue.Text) > 0 And XBOX.BoundText <> "" And CmdDel.Enabled
End Sub
Private Function StrBox()
Set BoxTable = New ADODB.Recordset
BoxTable.Open "SELECT * FROM file0_50 ORDER BY CODE ", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
StrBox = "#  " & ";       "
Do Until BoxTable.EOF
    StrBox = StrBox & "|#" & BoxTable!CODE & ";" & BoxTable!desca
    BoxTable.MoveNext
Loop
End Function
Private Function RetDefBox() As String
Dim loctable As New ADODB.Recordset
loctable.Open "file0_50", CON, adOpenStatic, adLockReadOnly, adCmdTable
If loctable.EOF And loctable.BOF Then Exit Function
loctable.MoveLast
If loctable.RecordCount = 1 Then
    loctable.MoveFirst
    RetDefBox = Trim(loctable!CODE & "")
End If
End Function

