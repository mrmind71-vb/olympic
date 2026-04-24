VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form T_Shop 
   Caption         =   "ÊÞÇÑíÑ ÇáãÕÇÑíÝ"
   ClientHeight    =   2580
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
   ScaleHeight     =   2580
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2715
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   600
      Width           =   1365
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   75
      Width           =   1140
   End
   Begin VB.TextBox xVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2715
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1012
      Width           =   1365
   End
   Begin VB.TextBox xVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2715
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1443
      Width           =   1365
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "ÚÑÖ"
      Height          =   465
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   525
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "ÎÑæÌ"
      Height          =   465
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1515
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
   Begin MSDBCtls.DBCombo xStore 
      Bindings        =   "T_Shop.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1215
      TabIndex        =   9
      Top             =   1875
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ãÎÒä"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1950
      Width           =   450
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÞíãÉ ÑÕíÏ Ãæá"
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
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÞíãÉ ÑÕíÏ ÃÎÑ"
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
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1605
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Çáì ÊÇÑíÎ :"
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
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   690
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ãä ÊÇÑíÎ :"
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
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   225
      Width           =   765
   End
End
Attribute VB_Name = "T_Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nBal1 As Double
Dim nBal2 As Double
Dim SalTable As Recordset
Dim BalItem1 As Recordset
Dim BalItem2 As Recordset
Dim PurchTable As Recordset
Dim RetSalTable As Recordset
Dim RetPurTable As Recordset
Dim ChargTable As Recordset
Dim IncomTable As Recordset
Dim DataTable As Recordset

Dim TargetTable As Recordset
Private Sub CmdApply_Click()
RepCharge1
End Sub
Private Sub RepCharge1()
Dim nComp As Double
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")

cString = "SELECT Sum(FILE6_20.TOTAL) AS T_Sum , Sum(FILE6_20.quant * FILE6_20.cost ) AS T_Cost" & _
           " From FILE6_20 " & _
           " Where FILE6_20.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " AND STORE = " & MyParn(xStore.BoundText)
           Set SalTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
SalTable.MoveFirst

cString = "SELECT Sum(FILE7_20.TOTAL) AS T_SUM " & _
           " From FILE7_20 " & _
           " Where FILE7_20.[DATE] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " AND STORE = " & MyParn(xStore.BoundText)
           Set PurchTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
PurchTable.MoveFirst


cString = "SELECT Sum(FILE6_10.TOTAL) AS T_Sum , Sum(FILE6_10.quant * cost ) AS T_Cost" & _
           " From FILE6_10 " & _
           " Where FILE6_10.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " AND STORE = " & MyParn(xStore.BoundText)
            
Set RetSalTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
RetSalTable.MoveFirst

cString = "SELECT Sum(FILE6_11.TOTAL) AS T_Sum " & _
           "From FILE6_11 " & _
           " Where FILE6_11.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " AND STORE = " & MyParn(xStore.BoundText)
            
Set RetPurTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
RetPurTable.MoveFirst

cString = "SELECT Sum(FILE1_81.TOTAL) AS T_Sum " & _
           "From FILE1_81 " & _
           " Where FILE1_81.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " AND STORE = " & MyParn(xStore.BoundText)
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
DataTable.MoveFirst
nout = TurnValue(DataTable.T_SUM, Null, 0)

cString = "SELECT Sum(FILE1_82.TOTAL) AS T_Sum " & _
           "From FILE1_82 " & _
           " Where FILE1_82.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " AND STORE = " & MyParn(xStore.BoundText)
            
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
DataTable.MoveFirst
ndam = TurnValue(DataTable.T_SUM, Null, 0)

cString = "SELECT Sum(FILE1_80.TOTAL) AS T_Sum " & _
           "From FILE1_80 " & _
           " Where FILE1_80.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " AND STORE = " & MyParn(xStore.BoundText)
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
DataTable.MoveFirst
nIn = TurnValue(DataTable.T_SUM, Null, 0)

cString = "SELECT Sum(FILE1_60.TOTAL) AS T_Sum " & _
           "From FILE1_60 " & _
           " Where FILE1_60.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " AND STORE1 = " & MyParn(xStore.BoundText)
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
DataTable.MoveFirst
nTrans_Out = TurnValue(DataTable.T_SUM, Null, 0)

cString = "SELECT Sum(FILE1_60.TOTAL) AS T_Sum " & _
           "From FILE1_60 " & _
           " Where FILE1_60.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " AND STORE2 = " & MyParn(xStore.BoundText)
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
DataTable.MoveFirst
nTrans_In = TurnValue(DataTable.T_SUM, Null, 0)


'***********
cString = "SELECT Sum(FILE8_50.value) AS T_Sum " & _
           "FROM FILE8_50 LEFT JOIN FILE0_50 ON FILE8_50.BOX = FILE0_50.CODE " & _
           " Where FILE8_50.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " and store = " & MyParn(xStore.BoundText)
Set ChargTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
ChargTable.MoveFirst

cString = "SELECT Sum(FILE8_60!value) AS T_Sum " & _
           "FROM FILE8_60 LEFT JOIN FILE0_50 ON FILE8_60.BOX = FILE0_50.CODE " & _
           " Where FILE8_60.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " and store = " & MyParn(xStore.BoundText)
Set IncomTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
IncomTable.MoveFirst

cString = "SELECT Sum(FILE0_10.Differ * FILE0_10.COST ) AS T_Sum " & _
           "From FILE0_10 " & _
           " Where FILE0_10.[Date] Between DateValue(" & MyParn(Date1.Text) & ")" & _
           " and  DateValue(" & MyParn(date2.Text) & ")" & _
           " AND STORE = " & MyParn(xStore.BoundText) & _
           " AND CLOSED = TRUE "
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
DataTable.MoveFirst
nComp = TurnValue(DataTable.T_SUM, Null, 0)
 
'*****************

nBal1 = Val(xVal1.Text)

nBal2 = Val(xVal2.Text)
    
TargetTable.AddNew
If SalTable.RecordCount > 0 Then TargetTable.VAL6 = TurnValue(SalTable.T_SUM, Null, 0)
If RetSalTable.RecordCount > 0 Then TargetTable.VAL7 = TurnValue(RetSalTable.T_SUM, Null, 0)

If PurchTable.RecordCount > 0 Then TargetTable.val2 = TurnValue(PurchTable.T_SUM, Null, 0)
If RetPurTable.RecordCount > 0 Then TargetTable.VAL3 = TurnValue(RetPurTable.T_SUM, Null, 0)

If ChargTable.RecordCount > 0 Then TargetTable.VAL4 = TurnValue(ChargTable.T_SUM, Null, 0)
If IncomTable.RecordCount > 0 Then TargetTable.VAL8 = TurnValue(IncomTable.T_SUM, Null, 0)


TargetTable.val1 = nBal1
TargetTable.val2 = nBal2
TargetTable.VAL3 = TurnValue(PurchTable.T_SUM, Null, 0)
TargetTable.VAL4 = TurnValue(SalTable.T_SUM, Null, 0)
TargetTable.VAL5 = TurnValue(RetPurTable.T_SUM, Null, 0)
TargetTable.VAL6 = TurnValue(RetSalTable.T_SUM, Null, 0)

TargetTable.VAL7 = nTrans_In
TargetTable.VAL8 = nTrans_Out

TargetTable.VAL9 = nIn
TargetTable.VAL10 = nout + ndam

TargetTable.VAL11 = TurnValue(ChargTable.T_SUM, Null, 0)
TargetTable.VAL12 = TurnValue(IncomTable.T_SUM, Null, 0)

If nComp > 0 Then
    TargetTable.VAL13 = Abs(nComp)
    TargetTable.VAL14 = 0
Else
    TargetTable.VAL14 = Abs(nComp)
    TargetTable.VAL13 = 0
End If
TargetTable.VAL16 = TurnValue(SalTable.T_SUM, Null, 0) - TurnValue(RetSalTable.T_SUM, Null, 0) - TurnValue(SalTable.t_COST, Null, 0) + TurnValue(RetSalTable.t_COST, Null, 0)

TargetTable.str1 = " ãä ÊÇÑíÎ " & Date1.Text & " Åáì ÊÇÑíÎ " & date2.Text
TargetTable.str2 = xStore.Text

TargetTable.str19 = firsttitle
TargetTable.str20 = Secondtitle
TargetTable.Update
       
Report1.ReportFileName = App.Path & "\Reports\T_SHOP.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
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
Function MYVALID()
If Not IsDate(Date1.Text) Then Exit Function
If Not IsDate(date2.Text) Then Exit Function
If xStore.BoundText = "" Then Exit Function

If DateValue(Date1.Text) > DateValue(date2.Text) Then Exit Function
MYVALID = True
End Function
Function myDateValue(pDate)
myDateValue = "#" & DateValue(pDate) & "#"
End Function
Private Sub Form_Load()
Data1.DatabaseName = MdbPath
Data1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
xStore.ListField = "Desca"
xStore.BoundColumn = "code"
End Sub
