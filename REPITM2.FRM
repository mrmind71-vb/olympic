VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form RepItem2 
   Caption         =   " ﬁ«—Ì— «·«’‰«›"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
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
   ScaleHeight     =   2670
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xGroup1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3225
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox xGroup2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3225
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1575
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   150
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2130
      Width           =   990
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2130
      Width           =   990
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2550
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   450
      Width           =   1365
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2550
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   1365
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
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
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "Repitm2.frx":0000
      DataSource      =   "Data3"
      Height          =   315
      Left            =   1125
      TabIndex        =   9
      Top             =   825
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label xDescGr2 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1650
      Width           =   1890
   End
   Begin VB.Label xDescGr1 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   1275
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1275
      Width           =   1740
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„Ã„Ê⁄… —∆Ì”Ì…"
      Height          =   195
      Left            =   4020
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   975
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï „Ã„Ê⁄…"
      Height          =   195
      Left            =   4020
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1725
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label LblGroup 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰ „Ã„Ê⁄…"
      Height          =   195
      Left            =   4020
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "«·ﬁ«∆„… :"
      Height          =   195
      Left            =   2475
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   75
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ï  «—ÌŒ :"
      Height          =   195
      Left            =   4020
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   555
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
      Height          =   195
      Left            =   4020
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   150
      Width           =   675
   End
End
Attribute VB_Name = "RepItem2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub RepItem21()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MyValid Then Exit Sub
i = 1
tempdb.Execute "DELETE * FROM TEMP"

Set TargetTable = tempdb.OpenRecordset("temp")
cCondition1 = " CDate < DateValue(" & MyParn(Date1.Text) & ")"
cField1 = myiif( _
           cCondition1 _
          , "IIF(ISNULL([IN]),0,[IN])-IIF(ISNULL([OUT]),0,[OUT])") & " As FirstBalance"

cField2 = myiif( _
               "CDate >= DateValue(" & MyParn(Date1.Text) & ")" & " and " & _
               " ctype = '6' OR cType = '0' " _
                , "[OUT]") & " as Sales"

cField3 = myiif( _
               "CDate >= DateValue(" & MyParn(Date1.Text) & ")" & " and " & _
               " ctype = '3' " _
                , "[IN]") & " as RetSales"

cField4 = myiif( _
               "CDate >= DateValue(" & MyParn(Date1.Text) & ")" & " and " & _
               " ctype = '2' " _
                , "[IN]") & " as Purchase"

cField5 = myiif( _
               "CDate >= DateValue(" & MyParn(Date1.Text) & ")" & " and " & _
               " ctype = '4' " _
                , "[IN]") & " as INPUT"

cField6 = myiif( _
               "CDate >= DateValue(" & MyParn(Date1.Text) & ")" & " and " & _
               " ctype = '8' " _
                , "[OUT]") & " as OUTPUT"

'CFeld8 = myiif(" IIF(ISNULL([IN]),0,[IN])-IIF(ISNULL([OUT]),0,[OUT])") & " As LastBalance"

cField9 = myiif( _
               "CDate >= DateValue(" & MyParn(Date1.Text) & ")" & " and " & _
               " ctype = 'z' " _
                , "[IN]") & " as Inventory"

cString = "Select File1_10.item as DescItem ,File1_10.Desca" & _
          cField1 & "," & cField2 & "," & cField3 & "," & cField4 & "," & cField5 & "," & _
          cField6 & "," & cField8 & "," & cField9 & _
          " from file1_11 inner join file1_10 on file1_11.item = file1_10.item " & _
          " Where cDate <= DateValue(" & MyParn(date2.Text) & ") " & _
          "Group by file1_10.item,FILE1_10.DescA"

Set SourceTable = mydb.OpenRecordset(cString, dbOpenForwardOnly)

If SourceTable.RecordCount = 0 Then Exit Sub
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .MidofItem
    TargetTable.str2 = .DESCA
    TargetTable.str3 = .FirstBalance
    TargetTable.str4 = .Sales
    TargetTable.STR5 = .RetSales
    TargetTable.STR6 = .Purchase
    TargetTable.str8 = .Output
    TargetTable.str9 = .inventory
    TargetTable.str10 = .LastBalance
    TargetTable.Str11 = .input
    TargetTable.Str12 = "ﬂ„Ì…"
        
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    
    TargetTable.str19 = firsttitle
    TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until .EOF
End With
Report1.ReportFileName = PublicPath & "\Reports\Repitm2.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Private Sub CmdApply_Click()
Select Case publicFlag
Case 2
    RepItem21
Case 7
    RepItem27
Case 8
    RepItem28
Case 14
    RepItem14
End Select
End Sub
Sub RepItem27()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MyValid Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cField1 = myiif("cType = '6' or ctype = '0' ", "[OUT]") & " as Sales,"
cField2 = myiif("cType = '3'", "[IN]") & " as RetSales, "
cField3 = myiif("cType = '6'or ctype = '0' ", "[TOTAL]") & " as SalesValue, "
cField4 = myiif("cType = '3'", "[TOTAL]") & " as RetSalesValue "

cString = "Select File1_11.Item as MidOfItem," & _
          "First(File1_10.DescA) as FirstOfDescA,File1_10.Pack," & _
          cField1 & cField2 & cField3 & cField4 & _
          " From File1_11 Inner Join file1_10 on file1_11.Item = file1_10.Item " & _
          " Where cDate Between DateValue(" & MyParn(Date1.Text) & ")" & _
          " and DateValue(" & MyParn(date2.Text) & ")" & _
          " Group By File1_11.Item,File1_10.Pack"

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "«·   ÊÃœ »Ì‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    If .Sales + .RetSales <> 0 Then
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.MidofItem
    TargetTable.str2 = SourceTable.FirstofDescA
    TargetTable.str3 = " ≈Ã„«·Ï „»Ì⁄«  «·√’‰«› ﬂ„Ì… - ﬁÌ„…"
    TargetTable.VAL1 = .Sales
    TargetTable.VAL2 = .RetSales
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
Report1.ReportFileName = PublicPath & "\Reports\RepItm7.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Sub RepItem28()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MyValid Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cField1 = myiif("cType = '2' ", "[IN]") & " as PURCH,"
cField2 = myiif("cType = '7'", "[OUT]") & " as RetPurch, "
cField3 = myiif("cType = '2'", "[TOTAL]") & " as PurchValue, "
cField4 = myiif("cType = '7'", "[TOTAL]") & " as RetPurchValue "


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
    TargetTable.VAL1 = .PURCH
    TargetTable.VAL2 = .RetPURCH
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
Report1.ReportFileName = PublicPath & "\Reports\RepItm7.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
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
Set FlagTable = mydb.OpenRecordset("File1_70")
FlagTable.Index = "nFile1_70_1"
If publicFlag = 12 Or 16 Then
'    Data1.DatabaseName = MdbPath
'    Data1.RecordSource = "Select Doc_No,DescA From File0_40 Group by Doc_No,DescA"
'    xList.BoundColumn = "Doc_no"
'    xList.ListField = "Desca"
End If
If publicFlag = 16 Then
    Data2.DatabaseName = MdbPath
    Data2.RecordSource = "Select Code,DescA From File1_70 Where Flag =  3 "
    xGroup.BoundColumn = "Code"
    xGroup.ListField = "Desca"
End If
If publicFlag = 2 Then
    xGroup.Visible = True
    xGroup1.Visible = True
    xGroup2.Visible = True
    xDescGr1.Visible = True
    xDescGr2.Visible = True
    
    Data1.DatabaseName = MdbPath
    Data1.RecordSource = "Select Code,DescA From File1_70 Where Flag =  3 "
    xGroup.BoundColumn = "Code"
    xGroup.ListField = "Desca"
End If
End Sub
Sub RepItem14()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MyValid Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cField1 = myiif("cType = '2'", "[in]") & " as Sales,"
cField2 = myiif("cType = '7'", "[out]") & " as RetSales, "
cField3 = myiif("cType = '2'", "[TOTAL]") & " as SalesValue, "
cField4 = myiif("cType = '7'", "[TOTAL]") & " as RetSalesValue "

cString = "Select File1_11.Item as MidOfItem," & _
          "First(File1_10.DescA) as FirstOfDescA," & _
          cField1 & cField2 & cField3 & cField4 & _
          " From File1_11 Inner Join file1_10 on file1_11.Item = file1_10.Item " & _
          " Where cDate Between DateValue(" & MyParn(Date1.Text) & ")" & _
          " and DateValue(" & MyParn(date2.Text) & ")" & _
          " Group By File1_11.Item"

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "«·   ÊÃœ »Ì‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    If .Sales > 0 Then
        TargetTable.AddNew
        TargetTable.str1 = SourceTable.MidofItem
        TargetTable.str2 = SourceTable.FirstofDescA
        TargetTable.str3 = " ≈Ã„«·Ï „‘ —Ì«  «·√’‰«› ﬂ„Ì… - ﬁÌ„…"
        TargetTable.VAL1 = .Sales
        TargetTable.VAL2 = .RetSales
        TargetTable.VAL3 = .Sales - .RetSales
        TargetTable.VAL7 = SourceTable.salesvalue
        TargetTable.VAL8 = SourceTable.retsalesvalue
        TargetTable.VAL9 = SourceTable.salesvalue - SourceTable.retsalesvalue
        
        TargetTable.Date1 = Date1.Text
        TargetTable.date2 = date2.Text
        TargetTable.str9 = firsttitle
        TargetTable.str10 = Secondtitle
    
        TargetTable.Update
    End If
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm14.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.MDB"
Report1.Action = 1
End Sub
Private Function MyValid()
If Not IsDate(Date1.Text) Or Not IsDate(date2.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If DateValue(Date1.Text) > DateValue(date2.Text) Then
    MsgBox " !! «· «—ÌŒ «·√Ê· √ﬂ»— „‰ «·À«‰Ï "
    Exit Function
End If
'If xList.Visible And xList.BoundText = "" Then Exit Function
If xGroup.Visible And xGroup.BoundText = "" Then Exit Function
MyValid = True
End Function

