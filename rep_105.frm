VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Rep_105 
   Caption         =   " ﬁ«—Ì— ⁄«„…"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Report1 
      Left            =   1800
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   2385
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Rep_105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim DataTable As Recordset
Dim BalItemTable As Recordset
Dim ClinTable As Recordset
Dim cString As String, cSupp As String
Dim PucrchTable As Recordset
Dim RetTable As Recordset
Dim SalTable As Recordset
Dim nBal As Double
Dim nOption As Integer
Function MYVALID()
'If xCode.BoundText = "" Then Exit Function
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xCode.Text = ""
xFact.BoundText = ""
xdesca.Caption = ""
End Sub
Private Sub Form_Load()
Set TempTable = tempdb.OpenRecordset("TEMP")
Set ClinTable = mydb.OpenRecordset("FILE4_10", dbOpenDynaset)
Set FlagTable = mydb.OpenRecordset("SELECT * FROM FILE1_70", dbOpenDynaset)
Data2.DatabaseName = MdbPath
Data2.RecordSource = "SELECT MOSM, MOSM FROM MOSM  "
xMosm.ListField = "MOSM"
xMosm.BoundColumn = "MOSM"

Data1.DatabaseName = MdbPath
Data1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 3 ORDER BY DESCA "
xFact.ListField = "DESCA"
xFact.BoundColumn = "CODE"
Data1.Refresh

Data3.DatabaseName = MdbPath
Data3.RecordSource = "SELECT CODE, DESCA FROM FILE1_50 ORDER BY DESCA "
xStore.ListField = "desca"
xStore.BoundColumn = "code"
xMosm.BoundText = "000"
End Sub
Private Sub CmdApply_Click()
Dim nRate As Double
Dim nTPurch, nTRet, nTSal, nTCost, nTSuppBal, nTCostBal
Dim nTIn As Double
If Not MYVALID Then Exit Sub

cField1 = myiif(" Type = '7' " _
           , "val(Format([out]))") & _
           " As Sale "

cField2 = myiif(" Type = '2' " _
           , "val(Format([IN]))") & _
           " As Purch "

cField3 = myiif(" Type = '9' " _
           , "val(Format([out]))") & _
           " As RetPurch "

cField4 = myiif(" Type = 'z' " _
           , "val(Format([IN ]))") & _
           " As Comp "

cField5 = myiif(" Type = 'F' " _
           , "val(Format([OUT ]))") & _
           " As Trans1 "

cField6 = myiif(" Type = 'T' " _
           , "val(Format([IN ]))") & _
           " As Trans2 "

cString = "SELECT FILE1_10.MOSM , FILE1_10.FACT , FILE1_10.MODELFACT , FILE1_10.SUPP , File1_10.modelno , FILE1_10.CODE,File1_10.model, FILE1_10.ITEM , First(File1_10.DescA) as ItemDesc, First(File1_10.SCAL) as F_SCAL, First(File1_10.C_SCAL) as C_SCAL, First(File1_10.COLOR) as COLOR,First(File1_10.C_COLOR) as C_COLOR,First(File1_10.COST) as COST ,First(File1_10.PRICE) as PRICE , " & _
           cField1 & "," & cField2 & "," & cField3 & "," & _
           cField4 & "," & cField5 & "," & cField6 & "," & _
          " Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT " & _
          " FROM FILE1_10 left JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM " & _
          " WHERE FILE1_10.ITEM IS NOT NULL "
If xFact.BoundText <> "" Then cString = cString & " AND  FILE1_10.fact = " & MyParn(xFact.BoundText)
If xCode.Text <> "" Then cString = cString & " AND  FILE1_10.CODE = " & MyParn(xCode.Text)
If xMosm.Text <> "" Then cString = cString & " and file1_10.mosm = " & MyParn(xMosm.Text)
If xStore.BoundText <> "" Then cString = cString & " and file1_11.STORE = " & MyParn(xStore.BoundText)

cString = cString & " GROUP BY FILE1_10.FACT , FILE1_10.MOSM ,FILE1_10.MODELFACT ,FILE1_10.CODE,FILE1_10.SUPP ,FILE1_10.MODELNO , FILE1_10.MODEL , FILE1_10.ITEM ORDER BY FILE1_10.MODEL, First(FILE1_10.C_COLOR), First(FILE1_10.C_SCAL) "
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
tempdb.Execute "DELETE * FROM TEMP"
Set TempTable = tempdb.OpenRecordset("TEMP")

With SourceTable
    Do While Not .EOF
        TempTable.AddNew
        TempTable.str4 = .SUPP
        TempTable.str1 = .MODELNO
        TempTable.str12 = DelZero(.MODELFACT)
        TempTable.Str11 = .mosm
        TempTable.str13 = SayCode(FlagTable, "3", .FACT)
        TempTable.str3 = .ItemDesc
        TempTable.STR5 = .Item
        TempTable.str2 = .Color & "  " & "„ﬁ«” " & .F_SCAL
        If XBAL.Value = 0 Then
            TempTable.str7 = " „Êﬁ› Ã—œ «·„ÊœÌ·«  " & xFact.Text & "   " & xStore.Text
        Else
            TempTable.str7 = " „Êﬁ› Ã—œ «·„ÊœÌ·«   ( „‘ —Ì«  - „»Ì⁄«  - —’Ìœ ) " & xFact.Text & "   " & xStore.Text
        End If
        TempTable.str8 = TurnValue(xdesca.Caption, "", Null)
        TempTable.str19 = firsttitle
        TempTable.str20 = Secondtitle
        TempTable.VAL1 = .purch
        TempTable.VAL2 = .RETPURCH
        TempTable.VAL3 = TurnValue(.SALE, Null, 0)
        TempTable.VAL5 = TurnValue(.SUMIN, Null, 0) - TurnValue(.SUMOUT, Null, 0)
'       TempTable.VAL4 = TurnValue(.COMP, Null, 0)
        TempTable.VAL4 = TurnValue(.TRANS2, Null, 0) - TurnValue(.TRANS1, Null, 0)
        nRate = 0
        nTIn = TurnValue(.purch, Null, 0) - TurnValue(.RETPURCH, Null, 0) + TurnValue(.TRANS2, Null, 0) - TurnValue(.TRANS1, Null, 0)
        If nTIn > 0 Then nRate = TempTable.VAL5 / nTIn * 100
        TempTable.VAL7 = .COST
        TempTable.VAL8 = .price
        TempTable.val6 = nRate
        
        TempTable.VAL11 = TurnValue(nTPurch, Null, 0)
        TempTable.VAL12 = TurnValue(nTRet, Null, 0)
        TempTable.VAL13 = TurnValue(nTPurch, Null, 0) - TurnValue(nTRet, Null, 0)
        
        TempTable.VAL14 = TurnValue(nTCost, Null, 0)
        TempTable.VAL15 = TurnValue(nTSal, Null, 0)
        TempTable.VAL16 = TurnValue(nTSal, Null, 0) - TurnValue(nTCost, Null, 0)
        
        TempTable.VAL17 = TurnValue(nTSuppBal, Null, 0)
        TempTable.VAL18 = TurnValue(nTCostBal, Null, 0)
        If xBal0.Value Then
            If TurnValue(TempTable.VAL5, Null, 0) <> 0 Then TempTable.Update
        Else
            TempTable.Update
        End If
        .MoveNext
    Loop
End With
If XBAL.Value = 0 Then
    Report1.ReportFileName = PublicPath & "\Reports\Rep_105.rpt"
Else
    Report1.ReportFileName = PublicPath & "\Reports\Rep_105S.rpt"
End If

Report1.DataFiles(0) = TempPath
Report1.Action = 1
End Sub
Function TurnValue(pSource, pOld, pNew)
   TurnValue = IIf(pSource = pOld Or (IsNull(pSource) And IsNull(pOld)), pNew, pSource)
End Function

Private Sub xCode_Change()
ClinTable.FindFirst " code = " & MyParn(xCode.Text)
xdesca.Caption = IIf(ClinTable.NoMatch, "", ClinTable.DESCA)
End Sub
Private Sub xCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(3)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„  From FILE4_10  WHERE CLOSE IS NULL "
    Generalarray(3) = "AND  DescA Like '%cFilter%'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 4000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Sub myProc()
If TypeOf ActiveControl Is TextBox Then
    ActiveControl.Text = GrdText(Search.Grid1, 0)
End If
Unload Search
End Sub
Private Sub xCODE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    ClinTable.FindFirst "Code = " & MyParn(xCode.Text)
    xdesca.Caption = IIf(ClinTable.NoMatch, "", TurnValue(ClinTable.DESCA, Null, ""))
End If
End Sub

