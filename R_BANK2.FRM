VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Bank2 
   Caption         =   "ÊÞÇÑíÑ ÇáÈäæß"
   ClientHeight    =   1860
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
   ScaleHeight     =   1860
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1155
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   1230
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4035
      Top             =   1215
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
      Left            =   165
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1035
      Width           =   2385
      Begin VB.CommandButton CmdExit 
         Caption         =   "ÎÑæÌ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "ÚÑÖ"
         Height          =   375
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   1005
      End
   End
   Begin MSDBCtls.DBCombo xBank 
      Bindings        =   "R_bank2.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   375
      TabIndex        =   6
      Top             =   675
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáÈäß :"
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
      Left            =   3975
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   675
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Çáì :"
      Height          =   225
      Left            =   1635
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   285
      Width           =   390
   End
   Begin VB.Label Label3 
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
      Left            =   3750
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   195
      Width           =   765
   End
End
Attribute VB_Name = "R_Bank2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim MoveBankTable As Recordset
Dim CodeTable As Recordset
Dim nBal
Dim nOption As Integer
Function MYVALID()
If Not IsDate(xdate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
If IsEmpty(xBank.BoundText) Then Exit Function
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xBank.BoundText = ""
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub Form_Load()
Set TempTable = tempdb.OpenRecordset("TEMP")
Set CodeTable = mydb.OpenRecordset("FILE1_70")
Set MoveBankTable = mydb.OpenRecordset("SELECT * FROM FILE5_11 ORDER BY BANK ,DATE , IN ")
Data1.DatabaseName = MdbPath
Data1.RecordSource = "FILE5_10"
xBank.BoundColumn = "ID"
xBank.ListField = "DESCA"
Data1.Refresh
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "Delete * From Temp"
With MoveBankTable
    .FindFirst " BANK = " & MyParn(xBank.BoundText)
    If Not .NoMatch Then
        nBal = 0
        Do While xBank.BoundText = MoveBankTable.bank
            nBal = nBal + TurnValue(.In, Null, 0) - TurnValue(.OUT, Null, 0)
            TempTable.AddNew
            TempTable.str3 = .DOC_ID
'           TempTable.STR2 = SayCode(CodeTable, cType, CODE_C)
            TempTable.str4 = .DESCA
            TempTable.STR2 = " ãä ÊÇÑíÎ " & xdate1.Text & "Åáì ÊÇÑíÎ" & xDate2.Text
            TempTable.str1 = "ÍÑßÉ Èäß " & xBank.Text
            TempTable.val1 = .In
            TempTable.VAL2 = .OUT
            TempTable.VAL3 = nBal
            TempTable.Date1 = .Date
            TempTable.str19 = firsttitle
            TempTable.str20 = Secondtitle
            TempTable.Update
            .MoveNext
            If .EOF Then Exit Do
        Loop
    End If
End With
Report1.ReportFileName = PublicPath & "\Reports\BANKREP2.RPT"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub

Function TurnValue(pSource, pOld, pNew)
   TurnValue = IIf(pSource = pOld Or (IsNull(pSource) And IsNull(pOld)), pNew, pSource)
End Function
Function myDateValue(pDate)
myDateValue = "#" & DateValue(pDate) & "#"
End Function

Function CalcPiad(cPass, pTable)
    CalcPiad = pTable.Value - pTable.PAID2 - IIf(IsNull(pTable.E_EXPIRE), 0, pTable.RATIO2)
    SubCreditTable.Seek ">=", cPass
    If Not SubCreditTable.NoMatch Then
        Do While SubCreditTable.CREDIT_NO = cPass
            CalcPiad = CalcPiad - IIf(IsNull(SubCreditTable.PAY_DATE), 0, SubCreditTable.Value_S)
            SubCreditTable.MoveNext
            If SubCreditTable.EOF Then Exit Do
        Loop
    End If
End Function
