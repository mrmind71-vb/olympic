VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Supp_Inv 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   1590
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
   ScaleHeight     =   1590
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
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
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   3435
      Begin VB.CommandButton Command3 
         Caption         =   "Œ—ÊÃ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   375
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4080
      Top             =   960
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
   Begin VB.Label ClientName 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·„Ê—œ"
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
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "Supp_Inv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim ClientTable As Recordset
Dim SubVenderTable As Recordset
Dim CashTable As Recordset
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Private Sub CmdApply_Click()
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.CreateDynaset("TEMP")
cString = "SELECT FILE4_11.DOC_ID, First(FILE4_11.DATE) AS FirstDATE, " & _
          "First(FILE4_11.plus) AS Firstplus, Sum(FILE8_20.value) AS Sumvalue " & _
          "FROM FILE4_11 LEFT JOIN FILE8_20 ON FILE4_11.DOC_NO = FILE8_20.INV_NO " & _
          "GROUP BY FILE4_11.DOC_no, FILE4_11.TYPE " & _
          "HAVING FILE4_11.TYPE = '4' AND FILE4_11.CODE = " & MyParn(xClient.Text)

With TempTable
SubVenderTable.Seek ">=", "4", xClient.Text
If Not SubVenderTable.NoMatch Then
    Do While Not SubVenderTable.EOF And SubVenderTable.code = xClient.Text
        .AddNew
        .str1 = SubVenderTable!DOC_NO
        .Date1 = SubVenderTable!Date
        .val1 = SubVenderTable!PLUS
        .str2 = ClientName.Caption
'       .Str3 = SubVenderTable.CODE
        .str3 = "„ «»⁄… √—’œ… ›Ê« Ì— „‘ —Ì«  „Ê—œ "
        .Update
        SubVenderTable.MoveNext
        If SubVenderTable.EOF Then Exit Do
    Loop
End If

.MoveFirst
If .RecordCount Then
    Do While Not .EOF
        .Edit
        .VAL2 = T_PayInv(.str1) + .VAL2
        .str19 = firsttitle
        .str20 = Secondtitle
        .Update
        .MoveNext
    Loop
End If

If .RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\CLInv.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdUndo_Click()
xClient.Text = ""
End Sub
Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Form_Load()
Set ClientTable = mydb.CreateDynaset("file4_10")
Set SubVenderTable = mydb.OpenRecordset("FILE4_11")
'Set Inv2Table = mydb.OpenRecordset("FILE0_30")
'Set InvChqTable = mydb.OpenRecordset("FILE5_21")
Set CashTable = mydb.OpenRecordset("FILE8_20")
'CashTable.Index = "nFILE8_20_3"
'SubVenderTable.Index = "nFILE4_11_2"
'Inv2Table.Index = "nFILE0_30_2"
'InvChqTable.Index = "nFILE5_21_2"
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
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê From File4_10"
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 3000
        
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
    CashTable.Index = "NFILE8_20_3"
    CashTable.Seek "=", cPass
    If myFound(CashTable, Array(cPass)) Then
        Do While CashTable.inv_No = cPass
            T_PayInv = T_PayInv + CashTable.Value
            CashTable.MoveNext
            If CashTable.EOF Then Exit Do
        Loop
    End If
End Function
