VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form SupRpt1 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
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
   ScaleHeight     =   2070
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   1275
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1575
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1575
      Width           =   1140
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1050
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   975
      Width           =   1290
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   975
      Width           =   1215
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3960
      Top             =   1320
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
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -975
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   405
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -945
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   45
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBCtls.DBCombo xGroup1 
      Bindings        =   "SUPRPT1.frx":0000
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   75
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DBCombo1"
      RightToLeft     =   -1  'True
   End
   Begin MSDBCtls.DBCombo xGroup2 
      Bindings        =   "SUPRPT1.frx":0014
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   525
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DBCombo1"
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   270
      Left            =   3825
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1050
      Width           =   555
   End
   Begin VB.Label iLabel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3825
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   975
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label LblGroup 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ï „Ã„Ê⁄… :"
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
      Left            =   3750
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰ „Ã„Ê⁄… :"
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
      Left            =   3825
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   150
      Width           =   1005
   End
End
Attribute VB_Name = "SupRpt1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClientTable As Recordset, movetable As Recordset, compTable As Recordset
Dim ChkTable As Recordset
Dim FlagTable As Recordset, BalanceTable As Recordset, CompBalanceTable As Recordset
Dim nOption As Integer
Private Sub SupModule2()
Dim sourcetable As Recordset
Dim TargetTable As Recordset
'Dim nMin As Integer
'nMin = 10
If Not IsDate(xdate1.Text) Or Not IsDate(xDate2.Text) Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
cField1 = myiif( _
          "[Date] < " & DateSql(xdate1.Text), _
          " iif(IsNull(sal),0,sal) - iif(IsNull(pay),0,pay)") & _
          " As FirstBalance"

cField2 = myiif( _
        " [Date] >= " & DateSql(xdate1.Text) & " and " & _
        " Type = '4' ", "sal") & _
        " As Purchases"
        
cField3 = myiif( _
        " Date >= " & DateSql(xdate1.Text) & " and " & _
        " Type = '5' ", "pay") & _
        " As RetPurchases"
            
cField4 = myiif( _
        " Date >= " & DateSql(xdate1.Text) & " and " & _
        " Type ='7' ", "pay") & _
        " As Payments"
        
cField5 = myiif( _
        " Date >= " & DateSql(xdate1.Text) & " and " & _
        " Type ='6' ", "sal") & _
        " As Cash"
        
cField6 = myiif("", _
        "iif(IsNull(sal),0,sal) - iif(IsNull(pay),0,pay)") & _
        " As LastBalance "
                
cString = "Select File4_11.Code,File4_10.DescA,File4_10.Group," & _
          cField1 & "," & cField2 & "," & cField3 & "," & _
          cField4 & "," & cField5 & "," & cField6 & _
          " , FILE1_70.DESCA AS GrDesc " & _
          " FROM (File4_11 INNER JOIN File4_10 ON File4_11.CODE = File4_10.CODE) LEFT JOIN FILE1_70 ON File4_10.group = FILE1_70.CODE " & _
          " Where FILE1_70.FLAG = 4 AND File4_10.Group Between " & MyParn(xGroup1.BoundText) & " and " & MyParn(xGroup2.BoundText) & _
          " and Date <= " & DateSql(xDate2.Text) & _
          " Group by File4_11.Code,File4_10.DescA,File4_10.Group , FILE1_70.DESCA"
      
Set sourcetable = mydb.OpenRecordset(cString, dbOpenSnapshot)
Set TargetTable = tempdb.CreateDynaset("TEMP")
If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

Do
    If Abs(sourcetable!FirstBalance) > nMin Or Abs(sourcetable!Purchases) > nMin Or Abs(sourcetable!retpurchases) > nMin Or _
       Abs(sourcetable!payments) > nMin Or Abs(sourcetable!CASH) > nMin Or _
       Abs(sourcetable!LastBalance) > nMin Then
        
        TargetTable.AddNew
        TargetTable.str1 = sourcetable!CODE
        TargetTable.str2 = sourcetable!desca
        TargetTable.str3 = sourcetable!Group
        TargetTable.str4 = sourcetable.GRDESC
        
        TargetTable.val1 = sourcetable.FirstBalance
        TargetTable.val2 = sourcetable!Purchases
        TargetTable.val3 = sourcetable!retpurchases
        TargetTable.val4 = sourcetable!payments
        TargetTable.val5 = sourcetable!CASH
        TargetTable.val7 = sourcetable.LastBalance
        TargetTable.str9 = firsttitle
        TargetTable.str10 = Secondtitle
        
        TargetTable.Date1 = xdate1.Text
        TargetTable.date2 = xDate2.Text

        TargetTable.Update
    End If
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\Supler02.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
Private Sub SupModule3()
Dim sourcetable As Recordset
Dim TargetTable As Recordset
'Dim nMin As Integer
nMin = 10

tempdb.Execute "DELETE * FROM TEMP"
cField1 = myiif( _
          "[Date] < " & DateSql(xdate1.Text), _
          " iif(IsNull(Plus),0,Plus) - iif(IsNull(Minus),0,Minus)") & _
          " As FirstBalance"

cField2 = myiif( _
        " [Date] >= " & DateSql(xdate1.Text) & " and " & _
        " Type = '4' ", "Plus") & _
        " As Purchases"
        
cField3 = myiif( _
        " Date >= " & DateSql(xdate1.Text) & " and " & _
        " Type = '5' ", "Minus") & _
        " As RetPurchases"
            
cField4 = myiif( _
        " Date >= " & DateSql(xdate1.Text) & " and " & _
        " Type ='7' ", "Minus") & _
        " As Payments"
        
cField5 = myiif( _
        " Date >= " & DateSql(xdate1.Text) & " and " & _
        " Type ='6' ", "Plus") & _
        " As Cash"
        
cField6 = myiif("", _
        "iif(IsNull(Plus),0,PLus) - iif(IsNull(Minus),0,Minus)") & _
        " As LastBalance "
                
cString = "Select File4_10.Group," & _
          cField1 & "," & cField2 & "," & cField3 & "," & _
          cField4 & "," & cField5 & "," & cField6 & _
          " From File4_11 Inner Join File4_10 on File4_11.Code = File4_10.Code " & _
          " Where File4_10.Group Between " & MyParn(xGroup1.BoundText) & " and " & MyParn(xGroup2.BoundText) & _
          " and Date <= " & DateSql(xDate2.Text) & _
          " Group by File4_10.Group "
      
Set sourcetable = mydb.OpenRecordset(cString, dbOpenSnapshot)
Set TargetTable = tempdb.CreateDynaset("TEMP")
If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

Do
    If Abs(sourcetable!FirstBalance) > nMin Or Abs(sourcetable!Purchases) > nMin Or Abs(sourcetable!retpurchases) > nMin Or _
       Abs(sourcetable!payments) > nMin Or Abs(sourcetable!CASH) > nMin Or _
       Abs(sourcetable!LastBalance) > nMin Then
        
        TargetTable.AddNew
        TargetTable.str1 = sourcetable!Group
        TargetTable.str2 = SayCode(FlagTable, 4, sourcetable!Group)
        
        TargetTable.val1 = sourcetable.FirstBalance
        TargetTable.val2 = sourcetable!Purchases
        TargetTable.val3 = sourcetable!retpurchases
        TargetTable.val4 = sourcetable!payments
        TargetTable.val5 = sourcetable!CASH
        TargetTable.val7 = sourcetable.LastBalance
        TargetTable.Date1 = xdate1.Text
        TargetTable.date2 = xDate2.Text
        TargetTable.STR19 = firsttitle
        ' TargetTable.str20 = Secondtitle

        TargetTable.Update
    End If
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\Supler03.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
Private Sub GroupReport()
Dim SourceTable1 As Recordset
Dim SourceTable2 As Recordset
Dim TargetTable As Recordset

tempdb.Execute "DELETE * FROM TEMP"
cField1 = myiif( _
          " vDate < DateValue(" & MyParn(xdate1.Text) & ")", _
          "iif(IsNull(Sal),0,Sal) - iif(IsNull(Pay),0,Pay)") & _
          " As FirstBalance"

cField2 = myiif( _
        " vDate >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " vType = '4' ", "Sal") & _
        " As Sales"
        
cField3 = myiif( _
        " vDate >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " vType = '5' ", "Pay") & _
        " As RetSales"
            
cField4 = myiif( _
        " vDate >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        "( vType ='8' OR VTYPE = '6' ) ", _
        " SAL") & _
        " As Cheq"

cField5 = myiif( _
        " vDate >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " vType ='7' " & _
        " And ( vNo = 'zz' or ( vNo Between '01' and '89' )) ", _
        " Pay") & _
        " As Payments"
        
cField6 = myiif( _
        " vDate >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " vType ='7' " & _
        " And  vNo Between '90' and '95'  ", _
        " Pay") & _
        " As Discount"

cField7 = myiif( _
        " vDate >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " vType ='7' " & _
        " And  vNo Between '96' and '99'  ", _
        " PAY") & _
        " As PriceDiffer"

cField8 = myiif( _
        " vDate >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " vType ='7' " & _
        " And  vNo = '00'  ", _
        " Pay") & _
        " As Settle "
        
cField9 = myiif("", _
        "iif(IsNull(Sal),0,Sal) - iif(IsNull(Pay),0,Pay)") & _
        " As LastBalance "
        
        
cString = "Select File3_10.vGroup," & _
          cField1 & "," & cField2 & "," & cField3 & "," & _
          cField4 & "," & cField5 & "," & cField6 & "," & _
          cField7 & "," & cField8 & "," & cField9 & _
          " From File3_11 Inner Join File3_10 on File3_11.Code = File3_10.Code " & _
          " Where File3_10.vGroup Between " & MyParn(xGroup1.BoundText) & " and " & MyParn(xGroup2.BoundText) & _
          " and vDate <= DateValue (" & MyParn(xDate2.Text) & ")" & _
          " Group by File3_10.vGroup "

Set sourcetable = mydb.OpenRecordset(cString, dbOpenSnapshot)
Set TargetTable = tempdb.CreateDynaset("TEMP")

If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = sourcetable.vGroup
    TargetTable.str2 = MyRetFlag(FlagTable, 3, sourcetable.vGroup)
    
    TargetTable.val1 = sourcetable.FirstBalance
    TargetTable.val2 = sourcetable.Sales - sourcetable.PriceDiffer
    TargetTable.val3 = sourcetable.Retsales
    TargetTable.val4 = sourcetable.payments
    TargetTable.val5 = sourcetable.Cheq
    TargetTable.Val6 = sourcetable.discount
    TargetTable.val7 = sourcetable.LastBalance
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle

    TargetTable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\Cl3_2.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
Private Sub CMDEXIT_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If TypeOf ActiveControl Is DBCombo And KeyCode = 46 Then ActiveControl.BoundText = ""
End Sub
Private Sub Form_Load()
Set FlagTable = mydb.OpenRecordset("File1_70", dbOpenDynaset)
Set SupTable = mydb.OpenRecordset("File4_10")
Select Case publicFlag
Case 0
    xDate2.Visible = False
Case 4, 6, 10
    LblGroup.Visible = False
    xGroup2.Visible = False
End Select
If publicFlag = 10 Then
    Label1.Visible = False
    LblGroup.Visible = False
    xGroup1.Visible = False
    xGroup2.Visible = False
End If
data1.DatabaseName = MdbPath
data1.RecordSource = "Select * From File1_70 where Flag = 4"
xGroup1.ListField = "Desca"
xGroup1.BoundColumn = "Code"

'Data2.DatabaseName = MdbPath
DATA2.DatabaseName = MdbPath
DATA2.RecordSource = "Select * From File1_70 where Flag = 4"
xGroup2.ListField = "Desca"
xGroup2.BoundColumn = "Code"
End Sub
Private Sub iLabel_Change()
If iLabel.Caption = "" Then
    xGroup2.BoundText = ""
    xGroup2.Enabled = False
Else
    DATA2.RecordSource = "Select * From ClGroup where Code >= " & MyParn(iLabel.Caption)
    DATA2.Refresh
    If Not xGroup2.Enabled Then xGroup2.Enabled = True
End If
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
Select Case publicFlag
Case 0
    SupModule1
Case 1
    SupModule2
Case 2
    SupModule3
Case 4, 6 'The Group invoices
    SupModule4    ' Group invoices
Case 10
    Supmoudule5
End Select
End Sub
Private Sub xGroup1_Click(Area As Integer)
iLabel.Caption = xGroup1.BoundText
End Sub
Private Function MYVALID()

If Not publicFlag = 10 Then
    If xGroup1.BoundText = "" Then Exit Function
    If xGroup2.Visible And xGroup2.BoundText = "" Then Exit Function
    MYVALID = True
Else
    MYVALID = True
End If
End Function
Private Sub SupModule1()
Dim SourceTable1 As Recordset
Dim SourceTable2 As Recordset
Dim TargetTable As Recordset

If Not MYVALID Then Exit Sub

tempdb.Execute "DELETE * FROM TEMP"
cField1 = myiif("", _
        " Val(Format(Sal)) - Val(Format(Pay))") & _
        " As Balance "

cField2 = myiif("Type = '9'", _
        " Val(Format(Pay))") & _
        " As SumofChq"

cString = "SELECT File4_11.CODE, File4_10.DESCA, File4_10.group, First(FILE1_70.DESCA) AS GrDesc , " & _
          cField1 & "," & _
          cField2 & _
          " FROM (File4_11 INNER JOIN File4_10 ON File4_11.CODE = File4_10.CODE) INNER JOIN FILE1_70 ON File4_10.group = FILE1_70.CODE" & _
          " WHERE (FILE1_70.FLAG = 4)  AND File4_10.Group Between " & MyParn(xGroup1.BoundText) & " and " & MyParn(xGroup2.BoundText) & _
          IIf(xdate1.Text <> "", " and Date <= " & DateSql(xdate1.Text), "") & _
          " GROUP BY File4_11.CODE, File4_10.DESCA, File4_10.group"

                          
Set sourcetable = mydb.OpenRecordset(cString, dbOpenSnapshot)
Set SourceTable2 = mydb.OpenRecordset("Select Code,Sum(value) as sumofValue From File5_20 Where Closed = '0' Group by code")

Set TargetTable = tempdb.OpenRecordset("TEMP")
If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

Do
    If Abs(sourcetable.Balance) >= 10 Then
        TargetTable.AddNew
        TargetTable.str1 = sourcetable.CODE
        TargetTable.str2 = sourcetable.[desca]
        TargetTable.str3 = sourcetable.[Group]
        TargetTable.str4 = sourcetable.GRDESC
        TargetTable.val1 = sourcetable.Balance
        If xdate1.Text = "" Then
            TargetTable.Str5 = " √—’œ… «·„Ê—œÌ‰ "
        Else
            TargetTable.Str5 = " √—’œ… «·„Ê—œÌ‰ Õ Ï  «—ÌŒ " & xdate1.Text
        End If
        TargetTable.val2 = sourcetable.sumOfChq
        TargetTable.STR19 = firsttitle
        ' TargetTable.str20 = Secondtitle

        TargetTable.Update
    End If
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\Comp1.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
Private Sub SupModule4() ' Group of suppliers invoice total
Dim sourcetable As Recordset
Dim TargetTable As Recordset
If publicFlag = 4 Then
    cString = " Select File4_11.Doc_No,File4_11.Plus,File4_11.Minus,File4_11.Date" & _
              " From file4_11 inner join file4_10 on file4_11.Code = File4_10.Code " & _
              " Where File4_11.[Type] = '4' " & _
              " And [Date] Between " & DateSql(xdate1.Text) & _
              " And " & DateSql(xDate2.Text) & _
              " And File4_10.Group = " & MyParn(xGroup1.BoundText)
ElseIf publicFlag = 6 Then
    cString = " Select File4_11.Doc_No,File4_11.Minus,File4_11.plus,File4_11.Date" & _
              " From file4_11 inner join file4_10 on file4_11.Code = File4_10.Code " & _
              " Where File4_11.[Type] = '5' " & _
              " And [Date] Between " & DateSql(xdate1.Text) & _
              " And " & DateSql(xDate2.Text) & _
              " And File4_10.Group = " & MyParn(xGroup1.BoundText)
End If
tempdb.Execute "DELETE * FROM TEMP"
Set sourcetable = mydb.CreateSnapshot(cString)
Set TargetTable = tempdb.CreateDynaset("TEMP")
If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = sourcetable.doc_no
    TargetTable.str2 = xGroup1.Text
    TargetTable.str3 = IIf(publicFlag <> 3, "„ «»⁄… √Ã„«·Ï ›Ê« Ì— „‘ —Ì« ", "„ «»⁄… ≈Ã„«·Ï ›Ê« Ì— „—œÊœ „‘ —Ì« ")
    TargetTable.str4 = "·„Ã„Ê⁄… «·„Ê—œÌ‰"
    
    TargetTable.val1 = IIf(publicFlag = 4, sourcetable.PLUS, sourcetable.minus)
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = sourcetable.Date
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle

    TargetTable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\Supler04.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
Sub Supmoudule5()
Dim sourcetable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")

cField1 = myiif( _
               "cType = '2'", "[IN]") & _
               " As PurQuant"

cField3 = myiif( _
               "cType = '7'", "[OUT]") & _
               " As RetPurQuant"


cField5 = myiif( _
               "cType = '2'", "[Total]") & _
               " As PurValue"

cField6 = myiif( _
               "cType = '7'", "[Total]") & _
               " As RetPurValue"

cString = "Select File1_11.Item," & _
          "File1_10.DescA,File1_10.Group," & _
          "File1_50.DescA," & _
          cField1 & "," & cField3 & "," & _
          cField5 & "," & cField6 & _
          " From ((File1_11 Inner Join File1_10 on File1_11.Item = File1_10.Item)" & _
          " Inner Join File4_10 on File1_11.Code = File4_10.Code)" & _
          " Inner Join file1_50 on File1_10.Group = File1_50.Code" & _
          " Where cDate >= " & DateSql(xdate1.Text) & _
          " and cDate <= " & DateSql(xDate2.Text) & _
          " And (ctype = '2' or ctype = '7') " & _
          "Group By " & _
          " File1_11.Item," & _
          "File1_10.DescA,File1_10.[Group]," & _
          "File1_50.DescA"
          
Set sourcetable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With sourcetable
If .RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str2 = xGroup1.Text
    TargetTable.str3 = sourcetable!Item
    TargetTable.str4 = sourcetable![FILE1_10.DESCA]
    TargetTable.val1 = !PurQuant
    TargetTable.val2 = !RetPurQuant
    TargetTable.val3 = !PurQuant - !RetPurQuant
'   TargetTable.str8 = TurnValue(!PurNumber, 0, "")
'   TargetTable.str9 = TurnValue(!RetPurNumber, 0, "")
'   TargetTable.str10 = TurnValue(!PurNumber - !RetPurNumber, 0, "")
    TargetTable.Str11 = ![file1_50.desca]
    
    TargetTable.val7 = !Purvalue
    TargetTable.val8 = !retPurvalue
    TargetTable.Val9 = !Purvalue - !retPurvalue
    
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle

    TargetTable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\supler06.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End With
End Sub
Function TCHK(pTable, cPass)
TCHK = 0
'pTable.MoveFirst
'Do While True
'    If pTable.code = cPass And pTable.CLOSED = "0" Then
'        TCHK = TCHK + pTable.Value
'    End If
'    pTable.MoveNext
'    If pTable.EOF Then Exit Do
'Loop
End Function
