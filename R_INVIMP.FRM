VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_InvImp 
   Caption         =   " ﬁ«—Ì— «·√’‰«›"
   ClientHeight    =   1590
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
   ScaleHeight     =   1590
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xInv_No 
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
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   750
      Width           =   3390
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   1065
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   390
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
         Width           =   1065
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2175
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1140
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4140
      Top             =   900
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
   Begin VB.Label ImpName 
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
      Caption         =   "«·›« Ê—…"
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
      Left            =   4065
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   285
      Width           =   585
   End
End
Attribute VB_Name = "R_InvImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ImpTable As Recordset
Dim TempTable As Recordset
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Private Sub CmdApply_Click()
    Salesinv
End Sub
Sub Salesinv()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
cString = "Select file6_20.*,FILE3_10.DESCA AS NameClient ,File1_10.DescA AS NameItem " & _
          "From (file6_20 inner join file1_10 on file6_20.item = file1_10.item ) INNER JOIN FILE3_10 ON FILE3_10.CODE = FILE6_20.CODE " & _
          " Where FILE6_20.CREDIT = " & MyParn(xInv_No.Text) & _
          " order by FILE6_20.Date,FILE6_20.DOC_NO "

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì—"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = xInv_No.Text
    TargetTable.str2 = ImpName.Caption
    TargetTable.str3 = SourceTable.DOC_NO
    TargetTable.str4 = SourceTable.Item
    TargetTable.STR5 = SourceTable.NameItem
    TargetTable.STR6 = SourceTable.NameClient
    
    
    TargetTable.val1 = SourceTable.price
    TargetTable.val2 = SourceTable.discount
    TargetTable.VAL3 = SourceTable.TOTAL
    TargetTable.VAL4 = SourceTable.Quant
    TargetTable.VAL5 = SourceTable.Number
   
    TargetTable.date3 = SourceTable.Date
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
Report1.ReportFileName = PublicPath & "\Reports\R_InvImp.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdUndo_Click()
xInv_No.Text = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set ImpTable = mydb.CreateDynaset("file7_30")
Set TempTable = tempdb.CreateDynaset("TEMP")
xInv_No.Text = ""
End Sub
Function MYVALID()
MYVALID = True
End Function
Private Sub xInv_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xInv_No.Text = ""
    
    Dim Generalarray(4)
    Dim GrdArray(3)
    
    Set Generalarray(1) = Me
    Generalarray(2) = "Select FILE7_30.INV_NO AS [—ﬁ„ «·›« Ê—…] , FILE7_30.DATE_INV AS [ «—ÌŒ «·›« Ê—…], FILE4_10.DESCA aS [«”„ «·⁄„Ì·] " & _
                      " FROM FILE7_30 INNER JOIN FILE4_10 ON FILE7_30.CODE = FILE4_10.CODE "
    Generalarray(3) = "Where File4_10.descA like '*cFilter*'"
    Generalarray(4) = "Group By Inv_No,Date_inv,File4_10.DescA"
    
    GrdArray(1) = 1000
    GrdArray(2) = 1000
    GrdArray(3) = 4000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub xInv_No_LostFocus()
If xInv_No.Text = "" Then Exit Sub
ImpTable.FindFirst "INV_NO = " & MyParn(xInv_No.Text)
If ImpTable.NoMatch Then Exit Sub
ImpName.Caption = ImpTable.BOAT
End Sub
