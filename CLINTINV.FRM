VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form ClintInv 
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
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
   ScaleHeight     =   1950
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   -990
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   810
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBCtls.DBCombo xComp 
      Bindings        =   "Clintinv.frx":0000
      Height          =   315
      Left            =   180
      TabIndex        =   11
      Top             =   900
      Visible         =   0   'False
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.TextBox xclient 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1230
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   510
      Width           =   1230
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1260
      Width           =   3435
      Begin VB.CommandButton Command3 
         Caption         =   "Œ—ÊÃ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   375
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4140
      Top             =   1440
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·‘—ﬂ… :"
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
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   990
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label ClientName 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   2295
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
      Left            =   3990
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   555
      Width           =   765
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
      Height          =   300
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   540
      Width           =   390
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·⁄„Ì· :"
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
      Left            =   3990
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   135
      Width           =   600
   End
End
Attribute VB_Name = "ClintInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClientTable As Recordset
Dim TempTable As Recordset
Dim SubVenderTable As Recordset
Dim Inv2Table As Recordset
Dim CashTable As Recordset
Dim ChqTable As Recordset
Dim InvChqTable As Recordset
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Sub Client_Inv_1()
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.CreateDynaset("TEMP")
'If xComp.Text <> "" Then
'    cString = cString & " and Comp = " & MyParn(xComp.Text)
'End If
With TempTable
Inv2Table.MoveFirst
Inv2Table.Seek ">=", xClient.Text
If Not Inv2Table.NoMatch Then
    Do While Not Inv2Table.EOF And Inv2Table.code = xClient.Text
        .AddNew
        .str1 = Inv2Table.DOC_ID
        .Date1 = Inv2Table.[Date]
        .val1 = Inv2Table.TOTAL
        .val2 = Inv2Table.PAID
        .str2 = ClientName.Caption
        .str3 = Inv2Table.code
        .str4 = "OLD"
        .str19 = firsttitle
        .str20 = Secondtitle

        .Update
        Inv2Table.MoveNext
        If Inv2Table.EOF Then Exit Do
    Loop
End If
SubVenderTable.Seek ">=", "4", xClient.Text
If Not SubVenderTable.NoMatch Then
    Do While Not SubVenderTable.EOF And SubVenderTable.code = xClient.Text
        .AddNew
        .str1 = SubVenderTable.DOC_ID
        .Date1 = SubVenderTable.vDate
        .val1 = SubVenderTable.sal
        .val2 = T_PayInv(SubVenderTable.DOC_ID)
        .str2 = ClientName.Caption
        .str3 = SubVenderTable.code
        .str19 = firsttitle
        .str20 = Secondtitle

        .Update
        SubVenderTable.MoveNext
        If SubVenderTable.EOF Then Exit Do
    Loop
End If
If .RecordCount = 0 Then
    MsgBox "«·   ÊÃœ »Ì‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
End With
'report1.reportfilename = publicPath & "\Reports\Client11.rpt"
'Report1.DataFiles(0) = APP.PATH & "\temp.mdb"
'Report1.Action = 1
End Sub
Private Sub CmdApply_Click()

tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.CreateDynaset("TEMP")
With TempTable
Inv2Table.MoveFirst
Inv2Table.Seek ">=", xClient.Text
If Not Inv2Table.NoMatch Then
    Do While Not Inv2Table.EOF And Inv2Table.code = xClient.Text
        If xCOMP.BoundText = "" Or xCOMP.BoundText = Inv2Table.COMP Then
            .AddNew
            .str1 = Inv2Table.DOC_ID
            .Date1 = Inv2Table.[Date]
            .val1 = Inv2Table.TOTAL
            .val2 = Inv2Table.PAID
            .str2 = ClientName.Caption
            .str3 = Inv2Table.code
            .str4 = "OLD"
            .str19 = firsttitle
            .str20 = Secondtitle
            .Update
        End If
        Inv2Table.MoveNext
        If Inv2Table.EOF Then Exit Do
    Loop
End If
SubVenderTable.Seek ">=", "4", xClient.Text
If Not SubVenderTable.NoMatch Then
    Do While Not SubVenderTable.EOF And SubVenderTable.code = xClient.Text
        If xCOMP.BoundText = "" Or xCOMP.BoundText = SubVenderTable.COMP Then
            .AddNew
            .str1 = SubVenderTable.DOC_ID
            .Date1 = SubVenderTable.vDate
            .val1 = SubVenderTable.sal
            .str2 = ClientName.Caption
            .str3 = SubVenderTable.code
            .str19 = firsttitle
            .str20 = Secondtitle
            .Update
        End If
        SubVenderTable.MoveNext
        If SubVenderTable.EOF Then Exit Do
    Loop
End If
.MoveFirst
If .RecordCount Then
Do While Not .EOF
    .Edit
    .val2 = T_PayInv(.str1) + .val2
    .str19 = firsttitle
    .str20 = Secondtitle

    .Update
    .MoveNext
Loop

End If
If .RecordCount = 0 Then
    MsgBox "«·   ÊÃœ »Ì‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\CLInv.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdUndo_Click()
xClient.Text = ""
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Form_Load()

Set ClientTable = mydb.CreateDynaset("file3_10")
Set SubVenderTable = mydb.OpenRecordset("FILE3_11")
Set Inv2Table = mydb.OpenRecordset("FILE0_30")
Set InvChqTable = mydb.OpenRecordset("FILE5_21")
Set CashTable = mydb.OpenRecordset("FILE8_10")
CashTable.Index = "nFILE8_10_3"
SubVenderTable.Index = "nFILE3_11_2"
Inv2Table.Index = "nFILE0_30_2"
InvChqTable.Index = "nFILE5_21_2"
If publicFlag = 1 Then
    xdate1.Visible = False
    xDate2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
Else
    xdate1.Visible = True
    xDate2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
End If
Data1.DatabaseName = MdbPath
Data1.RecordSource = "Company"
xCOMP.ListField = "Desca"
xCOMP.BoundColumn = "Code"
Data1.Refresh
End Sub
Function MYVALID()
ClientTable.FindFirst "code = " & MyParn(xClient.Text)
If ClientTable.NoMatch Then Exit Function
MYVALID = True
End Function
Private Sub xclient_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xClient.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(3)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê,GroupName As «·„Ã„Ê⁄… From QFile3_10_1"
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
Private Sub xClient_LostFocus()
If xClient.Text = "" Then Exit Sub
ClientTable.FindFirst "Code = " & MyParn(xClient.Text)
If ClientTable.NoMatch Then Exit Sub
ClientName.Caption = ClientTable.DESCA
End Sub
Function T_PayInv(cPass)
    InvChqTable.Index = "NFILE5_21_2"
    InvChqTable.Seek "=", cPass
    If myFound(InvChqTable, Array(cPass)) Then
        Do While InvChqTable.chq_no = cPass
            T_PayInv = T_PayInv + InvChqTable.PAID
            InvChqTable.MoveNext
            If InvChqTable.EOF Then Exit Do
        Loop
    End If
    CashTable.Index = "NFILE8_10_3"
    CashTable.Seek "=", cPass
    If myFound(CashTable, Array(cPass)) Then
        Do While CashTable.inv_No = cPass
            T_PayInv = T_PayInv + CashTable.vValue
            CashTable.MoveNext
            If CashTable.EOF Then Exit Do
        Loop
    End If
End Function
