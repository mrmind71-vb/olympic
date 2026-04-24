VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form ClinInv2 
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
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
   ScaleHeight     =   1800
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   525
      Width           =   1365
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDBCtls.DBCombo xComp 
      Bindings        =   "Clininv2.frx":0000
      Height          =   315
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin Crystal.CrystalReport Report1 
      Bindings        =   "Clininv2.frx":0014
      Left            =   3825
      Top             =   1125
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   975
      Width           =   3420
      Begin VB.CommandButton Command3 
         Caption         =   "Œ—ÊÃ"
         Height          =   360
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   195
         Width           =   1035
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   360
         Left            =   1192
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   195
         Width           =   990
      End
      Begin VB.CommandButton Cmdapply 
         Caption         =   "⁄—÷"
         Height          =   360
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   195
         Width           =   1035
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«· «—ÌŒ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   525
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‘—ﬂ… :"
      Height          =   195
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   90
      Width           =   615
   End
End
Attribute VB_Name = "ClinInv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cardtable As Recordset
Dim SubClientTable As Recordset
Dim CompClient As Recordset
Dim CashTable As Recordset
Dim InvChqTable As Recordset
Dim ClientTable As Recordset, movetable As Recordset
Dim invtable As Recordset
Dim nOption As Integer
Dim vDate As Date
Dim TempTable As Recordset
Private Sub CmdApply_Click()
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.CreateDynaset("TEMP")
vDate = Format(DateValue(xDate.Text), "DD-MM-YYYY")
With TempTable
ClientTable.MoveFirst
Do While Not ClientTable.EOF
    If ClientTable.code = "500037" Then
'        .str1 = ClientTable.vGroup
    End If
    .AddNew
    .str1 = ClientTable.vGroup
    .str2 = ClientTable.code
    .str3 = ClientTable.DESCA
    .str4 = xCOMP.Text
    .STR5 = Format(vDate, "mm-yyyy")
    .val1 = BalInv(ClientTable.code, vDate, True) - Def_Price(ClientTable.code, vDate)
    If Month(vDate) > 1 Then
        .STR6 = Format(DateAdd("m", -1, vDate), "mm-yyyy")
        .val2 = BalInv(ClientTable.code, DateAdd("m", -1, vDate), True)
    Else
        .STR6 = "„« ﬁ»· "
        .val2 = BalInv(ClientTable.code, vDate, False) - Return_Inv(ClientTable.code, vDate)
        GoTo 1000
    End If
    If Month(vDate) > 2 Then
        .STR7 = Format(DateAdd("m", -2, vDate), "mm-yyyy")
        .VAL3 = BalInv(ClientTable.code, DateAdd("m", -2, vDate), True)
    Else
        .STR7 = "„« ﬁ»· "
        .VAL3 = BalInv(ClientTable.code, DateAdd("m", -1, vDate), False) - Return_Inv(ClientTable.code, vDate)
        GoTo 1000
    End If
    If Month(vDate) >= 3 Then
        .str8 = "„« ﬁ»· "
        .VAL4 = BalInv(ClientTable.code, DateAdd("m", -2, Date), False) - Return_Inv(ClientTable.code, vDate)
    End If
    If .VAL4 < 0 Then
        .VAL3 = .VAL3 + .VAL4
        .VAL4 = 0
    End If
    If .VAL3 < 0 Then
        .val2 = .VAL3 + .val2
        .VAL3 = 0
    End If
    If .val2 < 0 Then
        .val1 = .val1 + .val2
        .val2 = 0
    End If
        .str19 = firsttitle
        .str20 = Secondtitle
1000    .Update
        ClientTable.MoveNext
        If ClientTable.EOF Then Exit Do
Loop
End With
Report1.ReportFileName = PublicPath & "\Reports\CLINV2.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdUndo_Click()
xDate.Text = ""
End Sub
Private Sub Form_Load()

Set ClientTable = mydb.OpenRecordset("FILE3_10")
Set SubClientTable = mydb.OpenRecordset("FILE3_11")
Set InvChqTable = mydb.OpenRecordset("FILE5_21")
Set CashTable = mydb.OpenRecordset("FILE8_10")
Set CompClient = mydb.OpenRecordset("FILE3_20")
Set invtable = mydb.OpenRecordset("FILE6_20")
CashTable.Index = "nFILE8_10_3"
SubClientTable.Index = "nFILE3_11_4"
InvChqTable.Index = "nFILE5_21_2"
CompClient.Index = "nFILE3_20_1"
invtable.Index = "nFILE6_20_1"

Data3.DatabaseName = MdbPath
Data3.RecordSource = "Company"
xCOMP.ListField = "Desca"
xCOMP.BoundColumn = "Code"
Data3.Refresh

End Sub
Function MYVALID()
If xGroup1.BoundText = "" Then Exit Function
If xGroup2.BoundText = "" Then Exit Function
If xDate.Text = "" Then Exit Function
MYVALID = True
End Function
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Function BalInv(cCode, dDate, lPass)
    BalInv = 0
    If lPass Then
        SubClientTable.Seek ">=", cCode, "4", dDate
        If Not SubClientTable.EOF Then
            If Not SubClientTable.NoMatch Then
            Do While SubClientTable.code = cCode And SubClientTable.vtype = "4"
                If Month(SubClientTable.vDate) = Month(dDate) And Year(SubClientTable.vDate) = Year(dDate) And SubClientTable.COMP = xCOMP.BoundText Then
                    BalInv = BalInv + SubClientTable.sal - T_PayInv(SubClientTable.DOC_ID)
                End If
                SubClientTable.MoveNext
                If SubClientTable.EOF Then Exit Do
            Loop
            End If
        End If
    Else
        SubClientTable.Seek ">=", cCode, "4"
        If Not SubClientTable.NoMatch Then
            Do While SubClientTable.code = cCode And SubClientTable.vtype = "4" And SubClientTable.vDate < dDate
                If SubClientTable.COMP = xCOMP.BoundText Then
                   BalInv = BalInv + SubClientTable.sal - T_PayInv(SubClientTable.DOC_ID)
                End If
                SubClientTable.MoveNext
                If SubClientTable.EOF Then Exit Do
            Loop
         End If
         CompClient.Seek "=", xCOMP.BoundText, cCode
         BalInv = BalInv + CompClient.balance
         CashTable.Index = "nFILE8_10_2"
         CashTable.Seek ">=", cCode
         If myFound(CashTable, Array(cCode)) Then
         Do While CashTable.code = cCode
            If CashTable.VNO >= "95" And CashTable.VNO <= "99" Then
            Else
                If CashTable.COMP = xCOMP.BoundText Then
                    If IsNull(CashTable.inv_No) Then
                        BalInv = BalInv - CashTable.vValue
                    Else
                       invtable.Seek "=", CashTable.inv_No
                       If invtable.NoMatch Then
                           BalInv = BalInv + CashTable.vValue
                        End If
                    End If
                End If
            
            End If
            CashTable.MoveNext
            If CashTable.EOF Then Exit Do
            Loop
        End If
        End If
End Function
Function T_PayInv(cPass)
    InvChqTable.Seek "=", cPass
    If Not InvChqTable.NoMatch Then
        Do While InvChqTable.chq_no = cPass
            T_PayInv = T_PayInv + InvChqTable.PAID
            InvChqTable.MoveNext
            If InvChqTable.EOF Then Exit Do
        Loop
    End If
    CashTable.Index = "nFILE8_10_3"
    CashTable.Seek "=", cPass
    If myFound(CashTable, Array(cPass)) Then
        Do While CashTable.inv_No = cPass
            T_PayInv = T_PayInv + CashTable.vValue
            CashTable.MoveNext
            If CashTable.EOF Then Exit Do
        Loop
    End If
End Function
Function Def_Price(cCode, dDate)
    CashTable.Index = "nFILE8_10_2"
    CashTable.Seek ">=", cCode, dDate
    If Not CashTable.NoMatch Then
        Do While CashTable.code = cCode
            If CashTable.VNO >= "95" And CashTable.VNO <= "99" And CashTable.COMP = xCOMP.BoundText Then
                Def_Price = Def_Price + CashTable.vValue
            End If
            CashTable.MoveNext
            If CashTable.EOF Then Exit Do
        Loop
    End If
End Function
Function Return_Inv(cCode, dDate)
    SubClientTable.Index = "nFILE3_11_4"
    SubClientTable.Seek ">=", cCode, "5", dDate
    If Not SubClientTable.NoMatch Then
        Do While SubClientTable.code = cCode And SubClientTable.vtype = "5" And SubClientTable.vDate <= dDate
            If SubClientTable.COMP = xCOMP.BoundText Then
                Return_Inv = Return_Inv + SubClientTable.PAY
            End If
            SubClientTable.MoveNext
            If SubClientTable.EOF Then Exit Do
        Loop
    End If
   
    CashTable.Index = "nFILE8_10_2"
    CashTable.Seek ">=", cCode, dDate
    If Not CashTable.NoMatch Then
        Do While CashTable.code = cCode
            If CashTable.VNO >= "95" And CashTable.VNO <= "99" Then
            Else
                If CashTable.COMP = xCOMP.BoundText Then
                    If IsNull(CashTable.inv_No) Then
                        Return_Inv = Return_Inv + CashTable.vValue
                    Else
'                       invtable.Seek "=", CashTable.inv_No
'                       If invtable.NoMatch Then
'                           Return_Inv = Return_Inv + CashTable.VValue
'                       End If
                    End If
                End If
                
            End If
            CashTable.MoveNext
            If CashTable.EOF Then Exit Do
        Loop
    End If
    
 End Function
