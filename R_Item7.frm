VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Item7 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
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
   ScaleHeight     =   1950
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   375
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   225
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   75
      Width           =   1515
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   450
      Width           =   1515
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   " ›—Ì€"
      Height          =   390
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1350
      Width           =   1515
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "«” Ã«»…"
      Height          =   390
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1350
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1515
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   450
      Top             =   825
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
      Bindings        =   "R_Item7.frx":0000
      DataSource      =   "Data3"
      Height          =   315
      Left            =   2100
      TabIndex        =   7
      Top             =   825
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„Œ“‰"
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
      Index           =   0
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   975
      Width           =   1005
   End
   Begin VB.Label Label4 
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
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ï  «—ÌŒ :"
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
      Index           =   1
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   825
   End
End
Attribute VB_Name = "R_Item7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
Select Case publicFlag
Case 7
    repitem7
End Select
End Sub
Private Sub CmdClear_Click()
xStore.BoundText = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
Data3.DatabaseName = MdbPath
Data3.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
xStore.ListField = "Desca"
xStore.BoundColumn = "code"
End Sub
Private Sub repitem7()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cField1 = myiif("Type = '2' ", "[IN]") & " as PURCH,"
cField2 = myiif("Type = '7'", "[OUT]") & " as RetPurch, "
cField3 = myiif("Type = '2'", "[TOTAL]") & " as PurchValue, "
cField4 = myiif("Type = '7'", "[TOTAL]") & " as RetPurchValue "

cString = "Select File1_11.Item as MidOfItem," & _
          "First(File1_10.DescA) as FirstOfDescA," & _
          cField1 & cField2 & cField3 & cField4 & _
          " From File1_11 Inner Join file1_10 on file1_11.Item = file1_10.Item " & _
          " Where Date Between " & DateSql(Date1.Text) & _
          " and " & DateSql(date2.Text)
If xStore.BoundText <> "" Then cString = cString & " and file1_11.store = " & MyParn(xStore.BoundText)
cString = cString & " Group By File1_11.Item"

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "«·   ÊÃœ »Ì‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    If .PURCH + .RetPURCH <> 0 Then
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.MidofItem
    TargetTable.str2 = SourceTable.FirstofDescA
    TargetTable.str3 = " ≈Ã„«·Ï „‘ —Ì«  «·√’‰«› ﬂ„Ì… - ﬁÌ„…" & xStore.Text
    TargetTable.val1 = .PURCH
    TargetTable.val2 = .RetPURCH
    TargetTable.VAL3 = .PURCH - .RetPURCH
    TargetTable.VAL7 = SourceTable.PURCHvalue
    TargetTable.VAL8 = SourceTable.retPURCHvalue
    TargetTable.VAL9 = SourceTable.PURCHvalue - SourceTable.retPURCHvalue
    
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.str9 = Mid(firsttitle, 1, 50)
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    End If
    SourceTable.MoveNext

Loop Until SourceTable.EOF
End With
REPORT1.ReportFileName = PublicPath & "\Reports\RepItm7.rpt"
REPORT1.DataFiles(0) = App.Path & "\Temp.MDB"
REPORT1.Action = 1
End Sub
Private Function MYVALID() As Boolean
If Not (IsDate(Date1.Text) And IsDate(date2.Text)) Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function

