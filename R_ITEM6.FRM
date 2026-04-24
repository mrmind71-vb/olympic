VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Item6 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   2835
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
   ScaleHeight     =   2835
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   75
      Width           =   1515
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   450
      Width           =   1515
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   " ›—Ì€"
      Height          =   390
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2250
      Width           =   1515
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   2100
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "R_Item6.frx":0000
      Height          =   315
      Left            =   1050
      TabIndex        =   7
      Top             =   825
      Visible         =   0   'False
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4650
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   2475
      Visible         =   0   'False
      Width           =   1440
   End
   Begin MSDBCtls.DBCombo xGroup1 
      Bindings        =   "R_Item6.frx":0014
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1050
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "«” Ã«»…"
      Height          =   390
      Left            =   3825
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2250
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   675
      RightToLeft     =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2250
      Width           =   1515
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   1050
      Top             =   2475
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
   Begin MSDBCtls.DBCombo xGroup2 
      Bindings        =   "R_Item6.frx":0028
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   1575
      Visible         =   0   'False
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
      Height          =   195
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   150
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ï  «—ÌŒ :"
      Height          =   195
      Index           =   1
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "„‰ „Ã„Ê⁄… :"
      Height          =   240
      Index           =   0
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1275
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "≈·Ï „Ã„Ê⁄… :"
      Height          =   255
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1650
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "„Ã„Ê⁄… —∆Ì”Ì… :"
      Height          =   195
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   900
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "R_Item6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
Select Case publicFlag
Case 6
    REPITEM6
Case 7
    repitem7
End Select
End Sub
Private Sub CmdClear_Click()
xGroup.BoundText = ""
xGroup1.BoundText = ""
xGroup2.BoundText = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
Data1.DatabaseName = MdbPath
Data1.RecordSource = "Select Code,DescA From File1_70 Where Flag = 2"
Data2.DatabaseName = MdbPath
Data2.RecordSource = "Select Code,DescA From File1_50 "
xGroup.BoundColumn = "Code"
xGroup.ListField = "DescA"
xGroup1.BoundColumn = "Code"
xGroup1.ListField = "Desca"
xGroup2.BoundColumn = "Code"
xGroup2.ListField = "Desca"
End Sub
Private Sub xGroup_Change()
Data2.RecordSource = "Select * From File1_50" & IIf(xGroup.BoundText <> "", " where M_GROUP = " & MyParn(xGroup.BoundText), "")
Data2.Refresh
End Sub
Private Sub REPITEM6()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cField1 = myiif("Type = '6' or type = '0' ", "[OUT]") & " as Sales,"
cField2 = myiif("Type = '3'", "[IN]") & " as RetSales, "
cField3 = myiif("Type = '6'or type = '0' ", "[TOTAL]") & " as SalesValue, "
cField4 = myiif("Type = '3'", "[TOTAL]") & " as RetSalesValue "

cString = "Select File1_11.Item as MidOfItem," & _
          "First(File1_10.DescA) as FirstOfDescA,File1_10.Pack," & _
          cField1 & cField2 & cField3 & cField4 & _
          " From File1_11 Inner Join file1_10 on file1_11.Item = file1_10.Item " & _
          " Where Date Between DateValue(" & MyParn(Date1.Text) & ")" & _
          " and DateValue(" & MyParn(date2.Text) & ")" & _
          " Group By File1_11.Item,File1_10.Pack"

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    If .Sales + .RetSales <> 0 Then
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.MidofItem
    TargetTable.str2 = SourceTable.FirstofDescA
    TargetTable.str3 = " ≈Ã„«·Ï „»Ì⁄«  «·√’‰«› ﬂ„Ì… - ﬁÌ„…"
    TargetTable.val1 = .Sales
    TargetTable.val2 = .RetSales
    TargetTable.VAL3 = .Sales - .RetSales
    TargetTable.VAL7 = SourceTable.salesvalue
    TargetTable.VAL8 = SourceTable.retsalesvalue
    TargetTable.VAL9 = SourceTable.salesvalue - SourceTable.retsalesvalue
    
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.str9 = Mid(firsttitle, 1, 50)
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    End If
    SourceTable.MoveNext

Loop Until SourceTable.EOF
End With
REPORT1.ReportFileName = PublicPath & "\Reports\RepItem6.rpt"
REPORT1.DataFiles(0) = App.Path & "\Temp.MDB"
REPORT1.Action = 1
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
          " and " & DateSql(date2.Text) & _
          " Group By File1_11.Item"

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
    TargetTable.str3 = " ≈Ã„«·Ï „‘ —Ì«  «·√’‰«› ﬂ„Ì… - ﬁÌ„…"
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

