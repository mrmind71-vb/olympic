VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Rep_201 
   Caption         =   " ﬁ«—Ì— ⁄«„…"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
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
   ScaleHeight     =   1590
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4050
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   1125
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2775
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "file_2"
      Top             =   1050
      Visible         =   0   'False
      Width           =   1065
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   300
      Top             =   150
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
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   675
      Width           =   2385
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1005
      End
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "Rep_201.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   825
      TabIndex        =   0
      Top             =   150
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄…"
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
      Left            =   4110
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   225
      Width           =   630
   End
End
Attribute VB_Name = "Rep_201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim ClinTable As Recordset
Dim Clin2Table As Recordset
Dim FClinTable As Recordset
Dim BalTable As Recordset
Dim SalTable As Recordset
Dim CashTable As Recordset
Dim CHQ2Table As Recordset
Dim nOption As Integer
Dim cString As String
Dim InvTable As Recordset
Dim RInvTable As Recordset

Dim nFBal As Double
Function MYVALID()
If xMosm.Text = "" Then
    MsgBox " ÕœÌœ «·„Ê”„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
End Sub
Private Sub Form_Load()
Set TempTable = tempdb.OpenRecordset("TEMP", dbOpenDynaset)
Data1.DatabaseName = MdbPath
Data1.RecordSource = "SELECT * FROM FILE1_70 WHERE FLAG = 12 "
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Data2.DatabaseName = MdbPath
Data2.RecordSource = "SELECT MOSM, MOSM FROM MOSM  "
xMosm.ListField = "MOSM"
xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm

End Sub
Private Sub CmdApply_Click()
    If Not MYVALID Then Exit Sub
        
    cStr1 = " SELECT FILE4_11.CODE, Sum(FILE4_11.PAY) AS SumPAY , Sum(FILE4_11.sal) AS Sumsal FROM FILE4_11 " & _
            " Where MOSM = " & MyParn(xMosm.Text)
    cStr1 = cStr1 & " GROUP BY FILE4_11.CODE "
    Set BalTable = mydb.OpenRecordset(cStr1, dbOpenSnapshot)
        
    cStr1 = " SELECT FILE4_11.CODE, Sum(FILE4_11.PAY) AS SumPAY , Sum(FILE4_11.sal) AS Sumsal FROM FILE4_11 " & _
            " Where MOSM = " & MyParn(xMosm.Text)
    cStr1 = cStr1 & " GROUP BY FILE4_11.CODE "
    Set BalTable = mydb.OpenRecordset(cStr1, dbOpenSnapshot)
        
    cField1 = myiif(" Type = '4' " _
               , "val(Format([SAL]))") & _
               " As Purch "
    
    cField2 = myiif(" Type = '5' " _
               , "val(Format([PAY]))") & _
               " As RET "
    
    cField3 = myiif(" Type = '7' " _
               , "val(Format([PAY]))") & _
               " As Cash "
    
    cField4 = myiif(" Type = '6' " _
               , "val(Format([SAL]))") & _
               " As CASH2 "
    
    cString = " SELECT FILE4_10.CODE, FILE4_10.F_balance, Sum(FILE4_11.PAY) AS SumPAY, Sum(FILE4_11.SAL) AS SumSAL , FILE4_10.DESCA , " & _
               cField1 & "," & cField2 & "," & cField3 & "," & cField4 & _
              " FROM FILE4_10 LEFT JOIN FILE4_11 ON FILE4_10.CODE = FILE4_11.CODE WHERE MOSM = " & MyParn(xMosm.Text)
    If xGroup.BoundText <> "" Then cString = cString & " and file4_10.GROUP = " & MyParn(xGroup.BoundText)
    cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.F_balance , FILE4_10.DESCA "

Set ClinTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
Set Clin2Table = mydb.OpenRecordset("SELECT FILE4_10.DESCA, FILE4_10.CODE FROM FILE4_10 ", dbOpenSnapshot)

cStr1 = " SELECT CODE , SUM(VALUE) AS TVALUE , CASH FROM FILE8_20 WHERE MOSM = " & MyParn(xMosm.Text) & " GROUP BY CODE , CASH "
Set CashTable = mydb.OpenRecordset(cStr1, dbOpenSnapshot)

tempdb.Execute "DELETE * FROM TEMP"
Set TempTable = tempdb.OpenRecordset("TEMP", dbOpenDynaset)
With Clin2Table
    .MoveFirst
    Do While Not .EOF
        nFBal = 0
        FClinTable.FindFirst " CODE = " & MyParn(.CODE)
        If Not FClinTable.NoMatch Then
            nFBal = TurnValue(FClinTable.BAL, Null, 0)
        End If
        TempTable.AddNew
        TempTable.STR1 = .CODE
        TempTable.str3 = !DESCA
        TempTable.VAL1 = nFBal
        
        ClinTable.FindFirst " code = " & MyParn(.CODE)
        If Not ClinTable.NoMatch Then
            InvTable.FindFirst " CODE = " & MyParn(.CODE) & " AND STORE = '0' "
            If Not InvTable.NoMatch Then TempTable.VAL2 = InvTable.TTOTAL
            
            RInvTable.FindFirst " CODE = " & MyParn(.CODE) & " AND STORE = '0' "
            If Not RInvTable.NoMatch Then TempTable.VAL3 = RInvTable.TTOTAL
            
            InvTable.FindFirst " CODE = " & MyParn(.CODE) & " AND STORE = 'zz' "
            If Not InvTable.NoMatch Then TempTable.VAL4 = InvTable.TTOTAL * -1
            
            RInvTable.FindFirst " CODE = " & MyParn(.CODE) & " AND STORE = 'zz' "
            If Not RInvTable.NoMatch Then TempTable.VAL4 = TurnValue(TempTable.VAL4, Null, 0) - (TurnValue(RInvTable.TTOTAL, Null, 0) * -1)
            
'           TempTable.val3 = ClinTable.RET
'            TempTable.val4 = TurnValue(ClinTable.CASH, Null, 0) - TurnValue(ClinTable.CASH2, Null, 0)
        End If
        
        CashTable.FindFirst " CASH  AND CODE = " & MyParn(.CODE)
        If Not CashTable.NoMatch Then TempTable.VAL5 = TurnValue(CashTable.TVALUE, Null, 0)
        
        CashTable.FindFirst " NOT CASH  AND CODE = " & MyParn(.CODE)
        If Not CashTable.NoMatch Then TempTable.val6 = TurnValue(CashTable.TVALUE, Null, 0)
        
        BalTable.FindFirst " CODE = " & MyParn(.CODE)
        If Not BalTable.NoMatch Then
            nBal = nFBal + TurnValue(BalTable.SUMSAL, Null, 0) - TurnValue(BalTable.SumPAY, Null, 0)
            TempTable.val7 = nBal
        End If
        TempTable.str7 = " ≈Ã„«·Ï √—’œ… &  ⁄«„·«  «·„Ê—œÌ‰ "
        TempTable.str8 = " „Ê”„ " & xMosm.Text
        If xGroup.BoundText <> "" Then
            TempTable.str7 = TempTable.str7 & " „Ã„Ê⁄… " & xGroup.Text
        End If
        TempTable.str19 = firsttitle
        TempTable.str20 = Secondtitle
        If TurnValue(TempTable.VAL1, Null, 0) <> 0 Or TurnValue(TempTable.VAL2, Null, 0) <> 0 Or TurnValue(TempTable.VAL3, Null, 0) <> 0 Or TurnValue(TempTable.VAL4, Null, 0) <> 0 Or TurnValue(TempTable.val6, Null, 0) <> 0 Then
            TempTable.Update
        End If
        .MoveNext
    Loop
End With

Report1.ReportFileName = PublicPath & "\Reports\Rep_201.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Function TurnValue(pSource, pOld, pNew)
   TurnValue = IIf(pSource = pOld Or (IsNull(pSource) And IsNull(pOld)), pNew, pSource)
End Function
