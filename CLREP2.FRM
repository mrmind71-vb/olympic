VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form CLIENTREPORT2 
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
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
   ScaleHeight     =   2535
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   -900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   810
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   975
      Width           =   1410
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   975
      Width           =   1410
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3960
      Top             =   1890
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
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1425
      Width           =   2280
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   180
         Width           =   1020
      End
   End
   Begin MSDBCtls.DBCombo xGroup1 
      Bindings        =   "Clrep2.frx":0000
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
      Bindings        =   "Clrep2.frx":0014
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1050
      Width           =   390
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
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
      TabIndex        =   10
      Top             =   1050
      Width           =   765
   End
   Begin VB.Label iLabel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3825
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   975
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label2 
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
      Left            =   3675
      RightToLeft     =   -1  'True
      TabIndex        =   7
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
      Left            =   3675
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   150
      Width           =   1005
   End
End
Attribute VB_Name = "CLIENTREPORT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClientTable As Recordset, movetable As Recordset, compTable As Recordset
Dim flagTable As Recordset, BalanceTable As Recordset, CompBalanceTable As Recordset
Dim nOption As Integer
Private Sub clientReport()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
Dim nMin As Integer
nMin = 10
tempdb.Execute "DELETE * FROM TEMP"

cField1 = myiif("Date < " & DateSql(xdate1.Text), "iif(IsNull(Sal),0,Sal) - iif(IsNull(Pay),0,Pay)") & " As FirstBalance"
cField2 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " Type = '4' ", "Sal") & _
        " As Sales"
        
cField3 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " Type = '5' ", "Pay") & _
        " As RetSales"
            
cField4 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " (  Type = '6' )", _
        " SAL") & _
        " As Cheq"

cField5 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " Type = '7' ", "Pay") & _
        " As Payments"

cField6 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " Type = '8' ", "sal") & _
        " As retchq"

cField7 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " Type = '9' ", "sal") & _
        " As chq"
        
cField9 = myiif("", _
        "iif(IsNull(Sal),0,Sal) - iif(IsNull(Pay),0,Pay)") & _
        " As LastBalance "
                
cString = "Select File3_11.Code  ,File3_10.DescA AS ClinDesca ,File3_10.Group, FILE1_70.DESCA AS GrDesca ," & _
          cField1 & "," & cField2 & "," & cField3 & "," & _
          cField4 & "," & cField5 & "," & cField6 & "," & _
          cField7 & "," & cField9 & _
          " FROM (File3_11 INNER JOIN File3_10 ON File3_11.CODE = File3_10.CODE) LEFT JOIN FILE1_70 ON File3_10.group = FILE1_70.CODE " & _
          " Where FILE1_70.FLAG = 3 AND File3_10.Group Between " & MyParn(xGroup1.BoundText) & " and " & MyParn(xGroup2.BoundText) & _
          " and Date <= DateValue (" & MyParn(xDate2.Text) & ")" & _
          " Group by File3_11.Code,File3_10.DescA,File3_10.Group , FILE1_70.DESCA"
      
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
Set TargetTable = tempdb.CreateDynaset("TEMP")
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

Do
    If Abs(SourceTable.FirstBalance) > nMin Or Abs(SourceTable.RetSales) > nMin Or Abs(SourceTable.RetSales) > nMin Or _
       Abs(SourceTable.payments) > nMin Or Abs(SourceTable.chq) > nMin Or Abs(SourceTable.retchq) > nMin Or _
       Abs(SourceTable.LastBalance) > nMin Then
       TargetTable.AddNew
       TargetTable.str1 = SourceTable.CODE
       TargetTable.str2 = SourceTable.ClinDescA
       TargetTable.str3 = SourceTable.[Group]
       TargetTable.STR4 = "„Ã„Ê⁄… ⁄„·«¡ " & SourceTable.GrDescA
       TargetTable.str6 = "≈Ã„«·Ï Õ—ﬂ… ⁄„·«¡ „‰  «—ÌŒ " & xdate1.Text & "≈·Ï " & xDate2.Text
        
       TargetTable.VAL1 = SourceTable.FirstBalance
       TargetTable.VAL2 = TurnValue(SourceTable.Sales, Null, 0)
       TargetTable.VAL3 = SourceTable.RetSales
       TargetTable.VAL4 = SourceTable.payments
       TargetTable.VAL5 = SourceTable.chq
       TargetTable.VAL6 = SourceTable.retchq
       TargetTable.VAL7 = SourceTable.LastBalance
       TargetTable.Date1 = xdate1.Text
       TargetTable.date2 = xDate2.Text
       TargetTable.str9 = firsttitle
       TargetTable.str10 = Secondtitle

        TargetTable.Update
    End If
    SourceTable.MoveNext
Loop Until SourceTable.EOF
Report1.ReportFileName = PublicPath & "\Reports\report2.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Private Sub GroupReport()
Dim SourceTable1 As Recordset
Dim SourceTable2 As Recordset
Dim TargetTable As Recordset

tempdb.Execute "DELETE * FROM TEMP"
cField1 = myiif( _
          " Date < DateValue(" & MyParn(xdate1.Text) & ")", _
          "iif(IsNull(Sal),0,Sal) - iif(IsNull(Pay),0,Pay)") & _
          " As FirstBalance"

cField2 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " Type = '4' ", "Sal") & _
        " As Sales"
        
cField3 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " Type = '5' ", "Pay") & _
        " As RetSales"
            
cField4 = myiif( _
        "  Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " (Type ='8' OR Type = '6' ) ", _
        " SAL") & _
        " As Cheq"

cField5 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " Type ='7' " & _
        " And ( No = 'zz' or ( No Between '01' and '89' )) ", _
        " Pay") & _
        " As Payments"
        
cField6 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " Type ='7' " & _
        " And  No Between '90' and '95'  ", _
        " Pay") & _
        " As Discount"

cField7 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " Type ='7' " & _
        " And  No Between '96' and '99'  ", _
        " PAY") & _
        " As PriceDiffer"

cField8 = myiif( _
        " Date >= DateValue(" & MyParn(xdate1.Text) & ") and " & _
        " Type ='7' " & _
        " And  No = '00'  ", _
        " Pay") & _
        " As Settle "
        
cField9 = myiif("", _
        "iif(IsNull(Sal),0,Sal) - iif(IsNull(Pay),0,Pay)") & _
        " As LastBalance "
        
cString = "Select File3_10.Group," & _
          cField1 & "," & cField2 & "," & cField3 & "," & _
          cField4 & "," & cField5 & "," & cField6 & "," & _
          cField7 & "," & cField8 & "," & cField9 & _
          " From File3_11 Inner Join File3_10 on File3_11.Code = File3_10.Code " & _
          " Where File3_10.Group Between " & MyParn(xGroup1.BoundText) & " and " & MyParn(xGroup2.BoundText) & _
          " and Date <= DateValue (" & MyParn(xDate2.Text) & ")" & _
          " Group by File3_10.Group "

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
Set TargetTable = tempdb.CreateDynaset("TEMP")

If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
    Do
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.Group
    TargetTable.str2 = RetFlag(flagTable, 3, SourceTable.Group)
    
    TargetTable.VAL1 = SourceTable.FirstBalance
    TargetTable.VAL2 = SourceTable.Sales
    TargetTable.VAL3 = SourceTable.RetSales
    TargetTable.VAL4 = SourceTable.payments
    TargetTable.VAL5 = SourceTable.Cheq
    TargetTable.VAL6 = SourceTable.discount
    TargetTable.VAL7 = SourceTable.LastBalance
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
  Loop Until SourceTable.EOF

Report1.ReportFileName = PublicPath & "\Reports\Cl3_2.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
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
Set flagTable = mydb.OpenRecordset("File1_70")
Set CompBalanceTable = mydb.OpenRecordset("File3_20")

Data1.DatabaseName = MdbPath
Data1.RecordSource = "ClGroup"
xGroup1.ListField = "Desca"
xGroup1.BoundColumn = "Code"
Data1.Refresh

Data2.DatabaseName = MdbPath
xGroup2.ListField = "Desca"
xGroup2.BoundColumn = "Code"

End Sub
Private Sub iLabel_Change()
If iLabel.Caption = "" Then
    xGroup2.BoundText = ""
    xGroup2.Enabled = False
Else
    Data2.RecordSource = "Select * From ClGroup where Code >= " & MyParn(iLabel.Caption)
    Data2.Refresh
    If Not xGroup2.Enabled Then xGroup2.Enabled = True
End If
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
Select Case publicFlag
Case 1
clientReport
Case 2
GroupReport
End Select
End Sub
Private Sub xGroup1_Click(Area As Integer)
iLabel.Caption = xGroup1.BoundText
End Sub
Private Function MYVALID()
If Not IsDate(xdate1.Text) Or Not IsDate(xDate2.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If DateValue(xdate1.Text) > DateValue(xDate2.Text) Then
    MsgBox " !! «· «—ÌŒ «·√Ê· √ﬂ»— „‰ «·À«‰Ï "
    Exit Function
End If
If xGroup1.BoundText = "" Then Exit Function
If xGroup2.BoundText = "" Then Exit Function
MYVALID = True
End Function

