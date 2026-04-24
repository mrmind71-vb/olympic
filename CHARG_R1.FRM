VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Charg_R1 
   Caption         =   " ﬁ«—Ì— «·„’«—Ì›"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
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
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1965
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2715
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1010
      Width           =   1365
   End
   Begin VB.TextBox xVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2715
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   1365
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   195
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2715
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   580
      Width           =   1365
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2715
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1365
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1440
      Top             =   0
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ﬁÌ„… —’Ìœ √Ê·"
      Height          =   195
      Left            =   4305
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1145
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ﬁÌ„… —’Ìœ √Œ—"
      Height          =   195
      Left            =   4305
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1605
      Width           =   990
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ï  «—ÌŒ :"
      Height          =   195
      Left            =   4305
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   685
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
      Height          =   195
      Left            =   4305
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   225
      Width           =   675
   End
End
Attribute VB_Name = "Charg_R1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nBal1 As Double
Dim nBal2 As Double
Dim FBalCTable As Recordset
Dim EBalCTable As Recordset
Dim FBalVTable As Recordset
Dim EBalVTable As Recordset
Dim SalTable As Recordset
Dim BalItem1 As Recordset
Dim BalItem2 As Recordset
Dim PurchTable As Recordset
Dim RetSalTable As Recordset
Dim RetPurTable As Recordset
Dim Cash1Table As Recordset
Dim cash2table As Recordset
Dim Cash3Table As Recordset
Dim Cash4Table As Recordset
Dim OutChq1Table As Recordset
Dim OutChq2Table As Recordset
Dim OutChq3Table As Recordset
Dim InChq1Table As Recordset
Dim InChq2Table As Recordset
Dim InChq3Table As Recordset
Dim ChargTable As Recordset
Dim IncomTable As Recordset
Dim TargetTable As Recordset
Private Sub CmdApply_Click()
RepCharge1
End Sub
Private Sub RepCharge1()
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")

cString = "SELECT Sum(FILE6_20.TOTAL) AS T_Sum " & _
           "From FILE6_20 " & _
           " Where FILE6_20.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
           Set SalTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
SalTable.MoveFirst

cString = "SELECT Sum(FILE7_20.TOTAL) AS T_SUM " & _
           "From FILE7_20 " & _
           " Where FILE7_20.[DATE] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
           Set PurchTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
PurchTable.MoveFirst


cString = "SELECT Sum(FILE6_10.TOTAL) AS T_Sum " & _
           "From FILE6_10 " & _
           " Where FILE6_10.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
            
Set RetSalTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
RetSalTable.MoveFirst

cString = "SELECT Sum(FILE6_11.TOTAL) AS T_Sum " & _
           "From FILE6_11 " & _
           " Where FILE6_11.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
            
Set RetPurTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
RetPurTable.MoveFirst

cString = "SELECT Sum(FILE8_10!value) AS t_Sum " & _
           "From FILE8_10 " & _
           " Where FILE8_10.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set Cash1Table = mydb.OpenRecordset(cString, dbOpenSnapshot)
Cash1Table.MoveFirst

cString = "SELECT Sum(FILE8_40!value) AS T_Sum " & _
           "From FILE8_40 " & _
           " Where FILE8_40.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set Cash4Table = mydb.OpenRecordset(cString, dbOpenSnapshot)
Cash4Table.MoveFirst

cString = "SELECT Sum(FILE8_20!value) AS T_Sum " & _
           "From FILE8_20 " & _
           " Where FILE8_20.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set cash2table = mydb.OpenRecordset(cString, dbOpenSnapshot)
cash2table.MoveFirst

cString = "SELECT Sum(FILE8_30.value) AS T_Sum " & _
           "From FILE8_30 " & _
           " Where FILE8_30.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set Cash3Table = mydb.OpenRecordset(cString, dbOpenSnapshot)
Cash3Table.MoveFirst

cString = "SELECT Sum(FILE5_20.value) AS T_Sum " & _
           "From FILE5_20 " & _
           " Where FILE5_20.[Date_R] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set InChq1Table = mydb.OpenRecordset(cString, dbOpenSnapshot)
InChq1Table.MoveFirst

cString = "SELECT Sum(FILE5_20.value) AS T_Sum " & _
           "From FILE5_20 " & _
           " Where FILE5_20.CLOSED = '1' AND FILE5_20.[Date_3] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set InChq2Table = mydb.OpenRecordset(cString, dbOpenSnapshot)
InChq2Table.MoveFirst

cString = "SELECT Sum(FILE5_20.value) AS T_Sum " & _
           "From FILE5_20 " & _
           " Where FILE5_20.CLOSED = '2' AND FILE5_20.[Date_3] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set InChq3Table = mydb.OpenRecordset(cString, dbOpenSnapshot)
InChq3Table.MoveFirst

cStrOUTg = "SELECT Sum(FILE5_21.value) AS T_Sum " & _
           "From FILE5_21 " & _
           " Where FILE5_21.[Date_R] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set OutChq1Table = mydb.OpenRecordset(cStrOUTg, dbOpenSnapshot)
OutChq1Table.MoveFirst

cStrOUTg = "SELECT Sum(FILE5_21.value) AS T_Sum " & _
           "From FILE5_21 " & _
           " Where FILE5_21.CLOSED = '1' AND FILE5_21.[Date_3] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set OutChq2Table = mydb.OpenRecordset(cStrOUTg, dbOpenSnapshot)
OutChq2Table.MoveFirst

cStrOUTg = "SELECT Sum(FILE5_21.value) AS T_Sum " & _
           "From FILE5_21 " & _
           " Where FILE5_21.CLOSED = '2' AND FILE5_21.[Date_3] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set OutChq3Table = mydb.OpenRecordset(cStrOUTg, dbOpenSnapshot)
OutChq3Table.MoveFirst

cString = "SELECT Sum(FILE8_50.value) AS T_Sum " & _
           "From FILE8_50 " & _
           " Where FILE8_50.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set ChargTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
ChargTable.MoveFirst

cString = "SELECT Sum(FILE8_60!value) AS T_Sum " & _
           "From FILE8_60 " & _
           " Where FILE8_60.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")"
Set IncomTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
IncomTable.MoveFirst
    
If Val(xVal1.Text) = 0 Then
    cString = " SELECT File1_11.ITEM, file1_10.COST, Sum(IIf(IsNull([in]),0,[in]))-Sum(IIf(IsNull([out]),0,[out])) AS Quant " & _
              " FROM File1_11 RIGHT JOIN file1_10 ON File1_11.ITEM = file1_10.ITEM " & _
              " WHERE FILE1_11!Date < DateValue(" & MyParn(Date1.Text) & ")" & _
              " GROUP BY File1_11.ITEM, file1_10.COST "
              
    Set BalItem1 = mydb.OpenRecordset(cString)
    If BalItem1.RecordCount > 0 Then
    BalItem1.MoveFirst
    nBal1 = 0
    Do While True
        nBal1 = nBal1 + (BalItem1.COST * BalItem1.Quant)
        BalItem1.MoveNext
        If BalItem1.EOF Then Exit Do
    Loop
    End If
Else
    nBal1 = Val(xVal1.Text)
End If

If Val(xVal2.Text) = 0 Then
    cString = " SELECT File1_11.ITEM, file1_10.COST, Sum(IIf(IsNull([in]),0,[in]))-Sum(IIf(IsNull([out]),0,[out])) AS Quant " & _
              " FROM File1_11 RIGHT JOIN file1_10 ON File1_11.ITEM = file1_10.ITEM " & _
              " WHERE FILE1_11!Date <= DateValue(" & MyParn(date2.Text) & ")" & _
              " GROUP BY File1_11.ITEM, file1_10.COST "
              
    Set BalItem2 = mydb.OpenRecordset(cString, dbOpenSnapshot)
    
    If BalItem2.RecordCount > 0 Then
    BalItem2.MoveFirst
        nBal2 = 0
        Do While True
            nBal2 = nBal2 + (BalItem2.COST * BalItem2.Quant)
            BalItem2.MoveNext
            If BalItem2.EOF Then Exit Do
        Loop
    End If
Else
    nBal2 = Val(xVal2.Text)
End If

cString = "SELECT Sum(FILE3_11.PAY) AS SumPAY, Sum(FILE3_11.SAL) AS SumSAL FROM FILE3_11 " & _
           " Where FILE3_11.[Date] <  DateValue(" & MyParn(Date1.Text) & ")"
Set FBalCTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
FBalCTable.MoveFirst
    
cString = "SELECT Sum(FILE3_11.PAY) AS SumPAY, Sum(FILE3_11.SAL) AS SumSAL FROM FILE3_11 " & _
           " Where FILE3_11.[Date] <=  DateValue(" & MyParn(date2.Text) & ")"
Set EBalCTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
EBalCTable.MoveFirst

cString = "SELECT Sum(FILE4_11.PAY) AS SumPAY, Sum(FILE4_11.SAL) AS SumSAL FROM FILE4_11 " & _
           " Where FILE4_11.[Date] <  DateValue(" & MyParn(Date1.Text) & ")"
Set FBalVTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
FBalVTable.MoveFirst

cString = "SELECT Sum(FILE4_11.PAY) AS SumPAY, Sum(FILE4_11.SAL) AS SumSAL FROM FILE4_11 " & _
           " Where FILE4_11.[Date] <=  DateValue(" & MyParn(date2.Text) & ")"
Set EBalVTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
EBalVTable.MoveFirst
     
    
TargetTable.AddNew
If SalTable.RecordCount > 0 Then TargetTable.VAL6 = TurnValue(SalTable.T_SUM, Null, 0)
If RetSalTable.RecordCount > 0 Then TargetTable.VAL7 = TurnValue(RetSalTable.T_SUM, Null, 0)

If PurchTable.RecordCount > 0 Then TargetTable.val2 = TurnValue(PurchTable.T_SUM, Null, 0)
If RetPurTable.RecordCount > 0 Then TargetTable.VAL3 = TurnValue(RetPurTable.T_SUM, Null, 0)

If ChargTable.RecordCount > 0 Then TargetTable.VAL4 = TurnValue(ChargTable.T_SUM, Null, 0)
If IncomTable.RecordCount > 0 Then TargetTable.VAL8 = TurnValue(IncomTable.T_SUM, Null, 0)

If FBalCTable.RecordCount > 0 Then TargetTable.VAL9 = TurnValue(FBalCTable.SUMSAL, Null, 0) - TurnValue(FBalCTable.SUMPAY, Null, 0)
If EBalCTable.RecordCount > 0 Then TargetTable.VAL14 = TurnValue(EBalCTable.SUMSAL, Null, 0) - TurnValue(EBalCTable.SUMPAY, Null, 0)
If FBalVTable.RecordCount > 0 Then TargetTable.VAL16 = TurnValue(FBalVTable.SUMSAL, Null, 0) - TurnValue(FBalVTable.SUMPAY, Null, 0)
If EBalVTable.RecordCount > 0 Then TargetTable.VAL21 = TurnValue(EBalVTable.SUMSAL, Null, 0) - TurnValue(EBalVTable.SUMPAY, Null, 0)

If Cash1Table.RecordCount > 0 Then TargetTable.VAL10 = TurnValue(Cash1Table.T_SUM, Null, 0)
If Cash4Table.RecordCount > 0 Then TargetTable.VAL11 = TurnValue(Cash4Table.T_SUM, Null, 0)
If cash2table.RecordCount > 0 Then TargetTable.VAL17 = TurnValue(cash2table.T_SUM, Null, 0)
If Cash3Table.RecordCount > 0 Then TargetTable.VAL18 = TurnValue(Cash3Table.T_SUM, Null, 0)

If InChq1Table.RecordCount > 0 Then TargetTable.VAL12 = TurnValue(InChq1Table.T_SUM, Null, 0)
If InChq2Table.RecordCount > 0 Then TargetTable.VAL13 = TurnValue(InChq2Table.T_SUM, Null, 0)
If InChq3Table.RecordCount > 0 Then TargetTable.VAL15 = TurnValue(InChq3Table.T_SUM, Null, 0)

If OutChq1Table.RecordCount > 0 Then TargetTable.VAL19 = TurnValue(OutChq1Table.T_SUM, Null, 0)
If OutChq2Table.RecordCount > 0 Then TargetTable.VAL20 = TurnValue(OutChq2Table.T_SUM, Null, 0)
If OutChq3Table.RecordCount > 0 Then TargetTable.VAL22 = TurnValue(OutChq3Table.T_SUM, Null, 0)

TargetTable.val1 = nBal1
TargetTable.VAL5 = nBal2
TargetTable.str1 = " „‰  «—ÌŒ " & Date1.Text & " ≈·Ï  «—ÌŒ " & date2.Text
TargetTable.str19 = firsttitle
TargetTable.str20 = Secondtitle
TargetTable.Update
    
   
REPORT1.ReportFileName = App.Path & "\Reports\T_CASH2.rpt"
REPORT1.DataFiles(0) = App.Path & "\Temp.MDB"
REPORT1.Action = 1
End Sub
Private Sub RepCharge2()
If Not MYVALID Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "Select Charge,Sum(vValue) as SumofValue," & _
          "File8_70.MainGroup,File8_70.DescA " & _
          " From file8_50 Inner Join File8_70 on File8_50.Charge = File8_70.Code " & _
          " Where vDate >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and vDate <= DateValue(" & MyParn(date2.Text) & ")" & _
          IIf(xCOMP.BoundText = "", "", " And Comp = " & MyParn(xCOMP.BoundText)) & _
          " Group By Charge," & _
          "File8_70.MainGroup,File8_70.DescA"
          
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .CHARGE
    TargetTable.str2 = .DESCA
    If xCOMP.BoundText <> "" Then TargetTable.str3 = xCOMP.Text
    TargetTable.str4 = .MainGroup
    TargetTable.STR5 = RetCharge(.MainGroup)
    TargetTable.val1 = .Sumofvalue
    
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
REPORT1.ReportFileName = App.Path & "\Reports\RepChrg1.rpt"
REPORT1.DataFiles(0) = App.Path & "\Temp.MDB"
REPORT1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub RepCharge3()
If Not MYVALID Then Exit Sub
If xCharge.BoundText = "" Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset

tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "Select vDate,DescA,vValue" & _
          " From file8_50" & _
          " Where vDate >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and vDate <= DateValue(" & MyParn(date2.Text) & ")" & _
          " and Charge = " & MyParn(xCharge.BoundText) & _
          IIf(xCOMP.BoundText = "", "", " And Comp = " & MyParn(xCOMP.BoundText)) & _
          " Order  By vDate"
Set SourceTable = mydb.CreateSnapshot(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .DESCA
    TargetTable.str2 = xCharge.Text
    If xCOMP.BoundText <> "" Then TargetTable.str3 = xCOMP.Text
    TargetTable.val1 = .vValue
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.date3 = .vDate
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
REPORT1.ReportFileName = App.Path & "\Reports\RepChrg3.rpt"
REPORT1.DataFiles(0) = App.Path & "\Temp.MDB"
REPORT1.Action = 1
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub

Private Function RetCharge(pCharge)
ChargeTable.FindFirst "Code = " & MyParn(pCharge)
If Not ChargeTable.NoMatch Then RetCharge = ChargeTable.DESCA
End Function

Private Sub Lbl_Click(Index As Integer)

End Sub

Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function MYVALID()
If Not IsDate(Date1.Text) Then Exit Function
If Not IsDate(date2.Text) Then Exit Function
If DateValue(Date1.Text) > DateValue(date2.Text) Then Exit Function
MYVALID = True
End Function
Function myDateValue(pDate)
myDateValue = "#" & DateValue(pDate) & "#"
End Function

