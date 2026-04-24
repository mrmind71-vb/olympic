VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Chqrep4 
   Caption         =   " ﬁ«—Ì— «·‘Ìﬂ« "
   ClientHeight    =   885
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
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   885
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2475
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   225
      Width           =   1230
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   3975
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
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   360
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·⁄„Ì·"
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
      Left            =   4095
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   225
      Width           =   510
   End
End
Attribute VB_Name = "Chqrep4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChqTable As Recordset
Dim TempTable As Recordset
Dim ClientTable As Recordset
Private Sub clientReport()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "Delete * From Temp"
cField1 = "Sum(iif(vDate <  DateValue('" & xDate1.Text & "'),Balance,0)) As SumFirstBalance"
cField2 = "Sum(iif(vDate >= DateValue('" & xDate1.Text & "'),Sales-Discount2,0)) as  SumSales"
cField3 = "Sum(iif(vDate >= DateValue('" & xDate1.Text & "'),RetSales,0)) as SumRetSales"
cField4 = "Sum(iif(vDate >= DateValue('" & xDate1.Text & "'),Cheq,0)) as SumCheq"
cField5 = "Sum(iif(vDate >= DateValue('" & xDate1.Text & "'),Payments,0)) as  SumPayments"
cField6 = "Sum(iif(vDate >= DateValue('" & xDate1.Text & "'),Discount,0)) as  SumDiscount"
cField7 = "Sum(Balance) as SumLastBalance"

cString = "Select DescA,vGroup,Client,ClientName," & _
           cField1 & "," & cField2 & "," & cField3 & "," & _
           cField4 & "," & cField5 & "," & cField6 & "," & cField7 & _
           " From qReport2 " & _
           " Where vDate <= DateValue('" & xDate2.Text & "')" & _
           " And vGroup >= " & MyParn(xGroup1.BoundText) & _
           " And vGroup <= " & MyParn(xGroup2.BoundText) & _
           " Group By DescA,vGroup,client,clientName " & _
           "Order by ClientName "
Set SourceTable = mydb.CreateSnapshot(cString)
Set TargetTable = mydb.CreateDynaset("TEMP")
Do
TargetTable.AddNew
TargetTable.str1 = SourceTable.Client
TargetTable.str2 = SourceTable.ClientName
TargetTable.str3 = SourceTable.vGroup
TargetTable.str4 = SourceTable.DESCA
TargetTable.val1 = SourceTable.SumFirstbalance
TargetTable.val2 = SourceTable.SumSales
TargetTable.VAL3 = SourceTable.SumRetSales
TargetTable.VAL4 = SourceTable.sumpayments
TargetTable.VAL5 = SourceTable.sumcheq
TargetTable.VAL6 = SourceTable.sumdiscount
TargetTable.VAL7 = SourceTable.sumLastBalance + SumFirstbalance
TargetTable.Date1 = xDate1.Text
TargetTable.date2 = xDate2.Text
TargetTable.str19 = firsttitle
TargetTable.str20 = Secondtitle

TargetTable.Update
SourceTable.MoveNext
Loop Until SourceTable.EOF
Report1.ReportFileName = PublicPath & "\Reports\report2.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub

Function CountBalance(Client)
Dim ClientMove As Recordset
NRETURN = 0
Set ClientMove = mydb.CreateDynaset("select * from file3_11 where code = " & MyParn(Client))
If ClientMove.RecordCount = 0 Then
    CountBalance = 0
    Exit Function
End If
ClientMove.MoveFirst
Do
    NRETURN = NRETURN + TurnValue(ClientMove.sal, Null, 0) - TurnValue(ClientMove.PAY, Null, 0)
    ClientMove.MoveNext
Loop Until ClientMove.EOF
CountBalance = NRETURN
End Function


Private Sub GroupReport()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "Delete * From Temp"
cField1 = "Sum(iif(vDate <  DateValue('" & xDate1.Text & "'),Balance,0)) As SumFirstBalance"
cField2 = "Sum(iif(vDate >= DateValue('" & xDate1.Text & "'),Sales,0)) as  SumSales"
cField3 = "Sum(iif(vDate >= DateValue('" & xDate1.Text & "'),RetSales,0)) as SumRetSales"
cField4 = "Sum(iif(vDate >= DateValue('" & xDate1.Text & "'),Cheq,0)) as SumCheq"
cField5 = "Sum(iif(vDate >= DateValue('" & xDate1.Text & "'),Payments,0)) as  SumPayments"
cField6 = "Sum(iif(vDate >= DateValue('" & xDate1.Text & "'),Discount,0)) as  SumDiscount"
cField7 = "Sum(Balance) as SumLastBalance"

cString = "Select DescA,vGroup," & _
           cField1 & "," & cField2 & "," & cField3 & "," & _
           cField4 & "," & cField5 & "," & cField6 & "," & cField7 & _
           " From qReport2 " & _
           " Where vDate <= DateValue('" & xDate2.Text & "')" & _
           " And vGroup >= " & MyParn(xGroup1.BoundText) & _
           " And vGroup <= " & MyParn(xGroup2.BoundText) & _
           " Group By vGroup,DescA"
Set SourceTable = mydb.CreateSnapshot(cString)
Set TargetTable = mydb.CreateDynaset("TEMP")

Do
TargetTable.AddNew
TargetTable.str1 = SourceTable.vGroup
TargetTable.str2 = SourceTable.DESCA
TargetTable.val1 = SourceTable.SumFirstbalance
TargetTable.val2 = SourceTable.SumSales
TargetTable.VAL3 = SourceTable.SumRetSales
TargetTable.VAL4 = SourceTable.sumpayments
TargetTable.VAL5 = SourceTable.sumcheq
TargetTable.VAL6 = SourceTable.sumdiscount
TargetTable.VAL7 = SourceTable.sumLastBalance
TargetTable.Date1 = xDate1.Text
TargetTable.date2 = xDate1.Text
TargetTable.str19 = firsttitle
TargetTable.str20 = Secondtitle

TargetTable.Update
SourceTable.MoveNext
Loop Until SourceTable.EOF
Report1.ReportFileName = PublicPath & "\Reports\Client2.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub

Function MYVALID()
'If Not IsDate(xdate1.Text) Then Exit Function
'If Not IsDate(xDate2.Text) Then Exit Function
If xCode.Text = "" Then Exit Function
MYVALID = True
End Function

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdUndo_Click()
xGroup1.BoundText = ""
xGroup2.BoundText = ""
xDate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub Form_Load()
Set TempTable = tempdb.OpenRecordset("TEMP")
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
tempdb.Execute "Delete * From Temp"
nOption = 1
cClosed = "0"

cString = "select file5_20.*  " & _
    "from file5_20 WHERE  File5_20.CODE = " & MyParn(xCode.Text)
Set ChqTable = mydb.CreateSnapshot(cString)
             
With ChqTable
If .RecordCount > 0 Then
.MoveFirst
cHead = "  ‘Ìﬂ«   Õ  «· Õ’Ì· ··⁄„Ì·  " & .Name1
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .CHK_ID
    TempTable.val1 = .Value
    TempTable.Date1 = .date_1
    TempTable.date2 = .date_R
    TempTable.str7 = cHead
    TempTable.str19 = firsttitle
    TempTable.str20 = Secondtitle

    TempTable.Update
    .MoveNext
   '  If .EOF Then Exit Do
Loop

End If
End With
Report1.ReportFileName = PublicPath & "\Reports\CHQREP4.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Function myDateValue(pDate1)
myDateValue = "DateValue('" & pDate1 & "')"
End Function

Private Sub xCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xCode.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(3)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê,GroupName As «·„Ã„Ê⁄… From qFile3_10_1"
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 2600
    GrdArray(3) = 1500
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub


