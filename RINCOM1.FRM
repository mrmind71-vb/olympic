VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form RIncom1 
   Caption         =   "ÊÞÇÑíÑ ÇáäÞÏíÉ"
   ClientHeight    =   1725
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
   ScaleHeight     =   1725
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "ÚÑÖ"
      Height          =   390
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1230
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "ÎÑæÌ"
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1230
      Width           =   1215
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3075
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1365
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3075
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1365
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   0
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
   Begin MSDBCtls.DBCombo xCharge 
      Bindings        =   "RIncom1.frx":0000
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label LblCharge 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÅíÑÇÏ"
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
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   360
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
      Left            =   4470
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   585
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
      Left            =   4485
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   225
      Width           =   765
   End
End
Attribute VB_Name = "RIncom1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChargeTable As Recordset
Private Sub CmdApply_Click()
Select Case publicFlag
Case 5
    RepCharge2
Case 6
    RepCharge3
End Select
End Sub
Private Sub RepCharge1()
If Not MYVALID Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "Select Charge,vValue,File8_50.DescA," & _
          "File8_70.MainGroup,File8_70.DescA " & _
          " From file8_50 Inner Join File8_70 on File8_50.Charge = File8_70.Code " & _
          " Where vDate >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and vDate <= DateValue(" & MyParn(date2.Text) & ")" & _
          " and File3_10.vGroup = " & MyParn(xGroup.BoundText) & _
          IIf(xCOMP.BoundText = "", "", " And Comp = " & MyParn(xCOMP.BoundText)) & _
          " Order By MainGroup,Charge"

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "áÇ ÊæÌÏ ÈíÇäÇÊ Ýì ÇáÊÞÑíÑ ¿"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .CHARGE
    TargetTable.str2 = .[file8_70.desca]
    If xCOMP.BoundText <> "" Then TargetTable.str3 = xCOMP.Text
    TargetTable.str4 = .MainGroup
    TargetTable.STR5 = RetCharge(.MainGroup)
    TargetTable.STR6 = .[file8_50.desca]
    TargetTable.VAL1 = .vValue
    
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
Report1.ReportFileName = App.Path & "\Reports\RepChrg1.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Private Sub RepCharge2()
If Not MYVALID Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "Select Charge,Sum(Value) as SumofValue," & _
          "File8_70.MainGroup,File8_70.DescA " & _
          " From file8_50 Inner Join File8_70 on File8_50.Charge = File8_70.Code " & _
          " Where [Date] >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and [Date] <= DateValue(" & MyParn(date2.Text) & ")" & _
          " Group By Charge," & _
          "File8_70.MainGroup,File8_70.DescA"
          
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "áÇ ÊæÌÏ ÈíÇäÇÊ Ýì ÇáÊÞÑíÑ ¿"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .CHARGE
    TargetTable.str2 = .DESCA
    TargetTable.str4 = .MainGroup
    TargetTable.STR5 = RetCharge(.MainGroup)
    TargetTable.VAL1 = .Sumofvalue

    TargetTable.str8 = "ÅÌãÇáì ãÕÇÑíÝ ãä ÊÇÑíÎ " & Date1.Text & " Åáì ÊÇÑíÎ " & date2.Text
'   TargetTable.Date1 = Date1.Text
'   TargetTable.date2 = date2.Text
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
Report1.ReportFileName = App.Path & "\Reports\RepChrg1.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
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
cString = "Select Date,DescA,Value" & _
          " From file8_50" & _
          " Where Date >= DateValue(" & MyParn(Date1.Text) & ")" & _
          " and Date <= DateValue(" & MyParn(date2.Text) & ")" & _
          " and Charge = " & MyParn(xCharge.BoundText) & _
          " Order  By Date"
Set SourceTable = mydb.CreateSnapshot(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "áÇ ÊæÌÏ ÈíÇäÇÊ Ýì ÇáÊÞÑíÑ ¿"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .DESCA
    TargetTable.str2 = xCharge.Text
    TargetTable.VAL1 = !Value
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.date3 = !Date
    TargetTable.str7 = " ÊÝÕíáì ãÕÑæÝ " & xCharge.Text
    TargetTable.str8 = " ãä ÊÇÑíÎ " & Date1.Text & " Åáì ÊÇÑíÎ " & date2.Text
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
Report1.ReportFileName = App.Path & "\Reports\RepChrg3.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Set ChargeTable = mydb.CreateSnapshot("Select * From File1_70 where Flag = 7")

If publicFlag = 3 Then
    Data2.DatabaseName = MdbPath
    Data2.RecordSource = "file8_70 "
    xCharge.BoundColumn = "Code"
    xCharge.ListField = "DescA"
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

