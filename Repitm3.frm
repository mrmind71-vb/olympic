VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form R_item3 
   Caption         =   " ﬁ«—Ì— «·«’‰«›"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
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
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   -900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBCtls.DBCombo Store 
      Bindings        =   "Repitm3.frx":0000
      Height          =   315
      Left            =   1875
      TabIndex        =   3
      Top             =   1425
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.TextBox XITEM 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   1515
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   1290
   End
   Begin VB.CommandButton Cmdexit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   195
      RightToLeft     =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1290
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   975
      Width           =   1515
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   525
      Width           =   1515
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   600
      Top             =   75
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
   Begin MSDBCtls.DBCombo Store2 
      Bindings        =   "Repitm3.frx":0014
      Height          =   315
      Left            =   1875
      TabIndex        =   4
      Top             =   1875
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label LblStore2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï „Œ“‰"
      Height          =   195
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1950
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label LblStore 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰ „Œ“‰"
      Height          =   195
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1500
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·’‰› :"
      Height          =   195
      Left            =   4515
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   150
      Width           =   540
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ï  «—ÌŒ :"
      Height          =   195
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
      Height          =   195
      Left            =   4515
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   675
   End
End
Attribute VB_Name = "R_item3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim movetable As Recordset
Dim storeTable As Recordset
Dim itemTable As Recordset
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
Select Case publicFlag
Case 3
    RepItem3
Case 4
    repitem4
Case 5
    RepItem5
Case 9
    RepItem9
End Select
End Sub
Private Sub RepItem3()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
cString = "Select File1_11.item as itemCode,File1_11.Doc_Id,FILE1_11.PRICE AS SumPrice," & _
          "File1_11.out As sumofOut, [total] as SumofTotal,File1_11.Date," & _
          "File1_10.Pack,File1_10.DescA As ItemDesc," & _
          "File3_10.DescA As ClientName " & _
          "From (File1_10 inner Join file1_11 on file1_10.item = File1_11.item) inner join file3_10 on file1_11.code = file3_10.code" & _
          " Where Date Between DateValue(" & MyParn(Date1.Text) & ")" & _
          " and DateValue(" & MyParn(date2.Text) & ")" & _
          " And File1_11.item = " & MyParn(xItem.Text) & _
          " and (Type = '6' OR TYPE = '0' )"
Set SourceTable = mydb.CreateSnapshot(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— "
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = !ItemCode
    TargetTable.str2 = !ItemDesc
    TargetTable.str3 = !DOC_ID
    TargetTable.str4 = !sumofOut

    TargetTable.STR5 = !ClientName
    
    TargetTable.VAL1 = !SumofTotal
    TargetTable.VAL2 = !sumofOut
    TargetTable.VAL3 = !SumPrice
    
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.date3 = !Date
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    
    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm4.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Sub repitem4()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
cString = "Select File1_11.item as MidOfItem,File1_11.Doc_Id," & _
          "File1_11.[IN] As sumofIn, [total] as SumofTotal,File1_11.Date,FILE1_11.PRICE AS SUMOFPRICE, " & _
          "File1_10.Pack,File1_10.DescA," & _
          "File4_10.DescA" & _
          " From (File1_10 inner Join file1_11 on file1_10.item = File1_11.item) inner join file4_10 on file1_11.code = file4_10.code" & _
          " Where Date Between DateValue(" & MyParn(Date1.Text) & ")" & _
          " and  DateValue(" & MyParn(date2.Text) & ")" & _
          " And File1_11.item = " & MyParn(xItem.Text) & _
          " and Type = '2' "
Set SourceTable = mydb.OpenRecordset(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— "
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .MidofItem
    TargetTable.str2 = .[FILE1_10.DESCA]
    TargetTable.str3 = .DOC_ID
    TargetTable.STR5 = .[file4_10.DescA]
    TargetTable.str4 = "„‰ " & Date1.Text & " ≈·Ï  «—ÌŒ " & date2.Text
    TargetTable.VAL1 = .SumofTotal
    TargetTable.VAL2 = .SUMOFPRICE
    TargetTable.VAL3 = .sumOfIn
    
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.date3 = !Date
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm5.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Sub RepItem5()
Dim nFirstBalance As Double
Dim SourceTable As Recordset
Dim TargetTable As Recordset
Dim CustTable As Recordset
Dim SuppTable As Recordset

Set SuppTable = mydb.OpenRecordset("FILE4_10", dbOpenSnapshot)
Set CustTable = mydb.OpenRecordset("FILE3_10", dbOpenSnapshot)

tempdb.Execute "DELETE * FROM TEMP"

' ⁄„· Ã„·… «·«” ⁄·«„ „‰ «Ã· —’Ìœ «Ê· «·„œ…
sWhere = " Where File1_10.item = " & MyParn(xItem.Text)
If Store.Text <> "" Then sWhere = sWhere & " and store = " & MyParn(Store.BoundText)
If Date1.Text <> "" Then sWhere = sWhere & " and Date < " & DateSql(Date1.Text)
' ⁄„·  «»· «Ê· «·„œ…
Set TargetTable = tempdb.OpenRecordset("TEMP")
cString = " Select Sum(iif(isNull(File1_11.[In]),0,[IN]) - iif(isNull(File1_11.[OUT]),0,[OUT])) as FirstBalance" & _
          " From File1_10 inner Join file1_11 on file1_10.item = File1_11.item" & _
          sWhere & _
          " Group by File1_10.Item"
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount > 0 Then
    nFirstBalance = SourceTable!FirstBalance
End If

' ⁄„· Ã„·… «·«” ⁄·«„ «·—∆Ì”Ì…

sWhere = ""
sWhere = " Where File1_10.item = " & MyParn(xItem.Text)
If Store.Text <> "" Then sWhere = sWhere & " and store >= " & MyParn(Store.BoundText)
If Store2.Text <> "" Then sWhere = sWhere & " and store <= " & MyParn(Store2.BoundText)
sWhere = sWhere & " and Date Between " & DateSql(Date1.Text) & _
                  " and " & DateSql(date2.Text)

cString = "Select File1_11.Item,File1_11.Doc_Id,File1_11.DescA, File1_11.CODE," & _
          "File1_11.[In],File1_11.out,File1_11.Date,Type," & _
          "File1_10.Pack,File1_10.DescA" & _
          " From File1_10 inner Join file1_11 on file1_10.item = File1_11.item" & _
           sWhere & _
          " Order by Date,File1_11.[in]"

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— "
    Exit Sub
End If

With SourceTable
    TargetTable.AddNew
    TargetTable.str1 = .[FILE1_10.DESCA] & " ====> " & xItem.Text
    TargetTable.str3 = "—’Ìœ ›Ï " & Date1.Text
    TargetTable.STR5 = nFirstBalance
    If nFirstBalance >= 0 Then
        TargetTable.VAL2 = nFirstBalance
    Else
        TargetTable.VAL1 = Abs(nFirstBalance)
    End If
    TargetTable.VAL3 = nFirstBalance
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.date3 = Date1.Text
    TargetTable.str9 = firsttitle
    TargetTable.str4 = TurnValue(Store.Text, "", Null)
    TargetTable.str6 = "  „‰  «—ÌŒ " & Date1.Text & " ≈·Ï  «—ÌŒ " & date2.Text
    
    TargetTable.Update
    Do While Not .EOF
            TargetTable.AddNew
            TargetTable.str1 = .[FILE1_10.DESCA] & " ====> " & xItem.Text
            TargetTable.str2 = .DOC_ID
            TargetTable.str3 = .[FILE1_11.DESCA]
            TargetTable.str7 = xItem.Text
            TargetTable.str4 = TurnValue(Store.Text, "", Null)
            If !Type = "2" Or !Type = "7" Then
                SuppTable.FindFirst " CODE = " & MyParn(.CODE)
                If Not SuppTable.NoMatch Then TargetTable.str8 = SuppTable.DESCA
            End If
            If !Type = "3" Or !Type = "6" Then
                CustTable.FindFirst " CODE = " & MyParn(.CODE)
                If Not CustTable.NoMatch Then TargetTable.str8 = CustTable.DESCA
            End If
            TargetTable.VAL1 = TurnValue(.OUT, Null, 0)
            TargetTable.VAL2 = TurnValue(.[In], Null, 0)
            TargetTable.VAL3 = nFirstBalance + TurnValue(.[In], Null, 0) - TurnValue(.[OUT], Null, 0)
'            TargetTable.str8 = cTitleString
            
            TargetTable.str6 = "  „‰  «—ÌŒ " & Date1.Text & " ≈·Ï  «—ÌŒ " & date2.Text
            TargetTable.date3 = !Date
            
            nFirstBalance = nFirstBalance + TurnValue(.[In], Null, 0) - TurnValue(.[OUT], Null, 0)
            TargetTable.str9 = firsttitle
            TargetTable.str10 = Secondtitle
            TargetTable.Update
            SourceTable.MoveNext
    Loop
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm6.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdClear_Click()
xItem.Text = ""
Date1.Text = ""
date2.Text = ""
Store.BoundText = ""
Store2.BoundText = ""
Expire.Value = False
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set storeTable = mydb.OpenRecordset("File1_30", dbOpenSnapshot)
Set movetable = mydb.OpenRecordset("file1_11", dbOpenSnapshot)
Set itemTable = mydb.OpenRecordset("File1_10", dbOpenSnapshot)

Data1.DatabaseName = MdbPath
Data1.RecordSource = "Select * From File1_70 where flag = 1 Order by Code"
Store.BoundColumn = "Code"
Store.ListField = "Desca"
Data1.Refresh
Data2.DatabaseName = MdbPath
Data2.RecordSource = "Select * From File1_70 where flag = 1 Order by Code"
Store2.BoundColumn = "Code"
Store2.ListField = "Desca"
Data2.Refresh
If publicFlag = 5 Then
    Store2.Visible = False
    LblStore2.Visible = False
End If
End Sub
Private Sub xItem_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ItemsLookup
End Sub
Sub ItemsLookup()
Dim GrdArray(2)
Dim Generalarray()
ReDim Generalarray(4)
Set Generalarray(1) = Me
Generalarray(2) = "Select Item as «·’‰›,DescA as [«”„ «·’‰›] From file1_10 "
Generalarray(3) = " Where DescA Like('*cFilter*')"
Generalarray(4) = "Order by Item"

    
GrdArray(1) = 1000
GrdArray(2) = 4500
    
Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search.Caption = "«” ⁄·«„ "
Search.Show 1
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Sub RepItem9()
Dim SourceTable As Recordset
Dim TargetTable As Recordset

sWhere = ""
sWhere = " Where File1_10.item = " & MyParn(xItem.Text)
If Store.Text <> "" Then sWhere = sWhere & " and store >= " & MyParn(Store.BoundText)
If Store2.Text <> "" Then sWhere = sWhere & " and store <= " & MyParn(Store2.BoundText)
sWhere = sWhere & " and Date Between " & DateSql(Date1.Text) & _
                  " and " & DateSql(date2.Text)
sWhere = sWhere & " and (Type = '6' or Type = '0') "

tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "Select FILE1_10.GROUP , File1_11.item as itemCode," & _
          "Sum(File1_11.Out) As sumofOut,Sum([total]) as SumofTotal," & _
          "First(File1_10.DescA) As ItemDesc," & _
          "File3_10.DescA As ClientName,File3_10.Code as ClientCode " & _
          "From (File1_10 inner Join file1_11 on file1_10.item = File1_11.item) inner join file3_10 on file1_11.code = file3_10.code" & _
          sWhere & _
          " Group by FILE1_10.GROUP , File1_11.item," & _
          "pack,File3_10.Code,File3_10.Desca"
          
Set SourceTable = mydb.CreateSnapshot(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— "
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .ItemCode
    TargetTable.str2 = .ItemDesc
    TargetTable.str3 = .sumofOut

    TargetTable.str4 = .ClientCode
    TargetTable.STR5 = .ClientName
    
    TargetTable.VAL1 = .SumofTotal
    TargetTable.VAL3 = .sumofOut
    
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
Report1.ReportFileName = PublicPath & "\Reports\RItm9_1.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Private Sub Store_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then Store.BoundText = ""
End Sub
Private Sub Store2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then Store2.BoundText = ""
End Sub
Private Function MYVALID() As Boolean
If xItem.Text = "" Then
    MsgBox "·« ÌÊÃœ ﬂÊœ «·’‰›"
    Exit Function
End If

itemTable.FindFirst "item = " & MyParn(xItem.Text)
If itemTable.NoMatch Then
    MsgBox "·« ÌÊÃœ ’‰› »Â–« «·ﬂÊœ "
    Exit Function
End If

If Not (IsDate(Date1.Text) And IsDate(date2.Text)) Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function
