VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form SupRpt16 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
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
   ScaleHeight     =   1620
   ScaleWidth      =   4980
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
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   750
      Width           =   3690
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   390
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   165
         Width           =   1140
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4200
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label xCodeName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   225
      Width           =   2340
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
      Left            =   4170
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   285
      Width           =   480
   End
End
Attribute VB_Name = "SupRpt16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SuplierTable As Recordset
Dim temptable As Recordset
Sub myProc()
ActiveControl.Text = GrdText(Search.grid1, 0)
Unload Search
End Sub
Private Sub CmdApply_Click()
Dim datatable As Recordset
Dim nTot1 As Double
Dim nTot2 As Double
nTot1 = 0
nTot2 = 0
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")

TargetTable.AddNew
cStr1 = " SELECT Sum(FILE1_11.[IN]) AS TIN, Sum(FILE1_11.OUT) AS TOUT,Sum(FILE1_11.[IN] * ITEMCODE.cost) AS VIN, Sum(FILE1_11.OUT * ITEMCODE.cost) AS VOUT ,Sum(FILE1_11.[IN] * ITEMCODE.price) AS VIN2, Sum(FILE1_11.OUT * ITEMCODE.price) AS VOUT2 , FILE1_11.ITEM, FILE1_10.DESCA , ITEMCODE.cost, FILE1_10.PRICE , FILE1_10.FACTCODE  " & _
        " FROM (FILE1_11 LEFT JOIN ITEMCODE ON FILE1_11.ITEM = ITEMCODE.ITEM) INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM WHERE LastOfCODE = " & MyParn(xCode.Text)
cStr1 = cStr1 & " GROUP BY FILE1_11.ITEM, FILE1_10.DESCA , FILE1_10.PRICE , ITEMCODE.cost , FILE1_10.FACTCODE "
Set datatable = mydb.OpenRecordset(cStr1, dbOpenSnapshot)

With datatable

If .RecordCount > 0 Then
    Do While Not .EOF
        nTot1 = nTot1 + TurnValue(.TIN, Null, 0) - TurnValue(.TOUT, Null, 0)
        nTot2 = nTot2 + TurnValue(.VIN, Null, 0) - TurnValue(.VOUT, Null, 0)
       .MoveNext
    Loop
End If
If .RecordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
        If (TurnValue(.TIN, Null, 0) - TurnValue(.TOUT, Null, 0)) <> 0 Then
            temptable.AddNew
            temptable.Str5 = .Item
            temptable.str9 = .desca
            temptable.STR13 = Format(.COST, "#0.00")
            temptable.Str11 = Format(TurnValue(.TIN, Null, 0) - TurnValue(.TOUT, Null, 0), "#0.00")
            temptable.STR14 = Format(TurnValue(.VIN, Null, 0) - TurnValue(.VOUT, Null, 0), "#0.00")
            
            temptable.str18 = Format(nTot1, "#0.00")
            temptable.str17 = Format(nTot2, "#0.00")
            
            temptable.STR15 = Format(nTot2 * TurnValue(SuplierTable.DISC, Null, 0) / 100, "#0.00")
            temptable.STR19 = Format(nTot2 - (nTot2 * TurnValue(SuplierTable.DISC, Null, 0) / 100), "#0.00")
            
            temptable.str4 = .FACTCODE
            temptable.str3 = xCodeName.Caption
            temptable.Update
        End If
       .MoveNext
    Loop
End If
End With
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\RSupp_16.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1

End Sub
Private Sub CmdUndo_Click()
xclient.Text = ""
End Sub
Private Sub CMDEXIT_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set SuplierTable = mydb.OpenRecordset("File4_10", dbOpenDynaset)
Set temptable = tempdb.OpenRecordset("TEMP")
Set FlagTable = mydb.OpenRecordset("File1_70")
End Sub
Function MYVALID() As Boolean
If xCode.Visible Then
    SuplierTable.FindFirst "Code = " & MyParn(xCode.Text)
    If SuplierTable.NoMatch Then Exit Function
End If
If Not (IsDate(xdate1.Text) And IsDate(xDate2.Text)) Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function

Private Sub xCode_Change()
SuplierTable.FindFirst "Code =" & MyParn(xCode.Text)
xCodeName.Caption = IIf(SuplierTable.NoMatch, "", SuplierTable.desca)
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xCode.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(3)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê From File4_10"
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
