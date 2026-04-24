VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form SupRpt2 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   3405
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
   ScaleHeight     =   3405
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
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   750
      Width           =   1290
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   750
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2325
      Width           =   3690
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   390
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   165
         Width           =   1140
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4140
      Top             =   1440
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
      TabIndex        =   10
      Top             =   225
      Width           =   2340
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
      Top             =   750
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
      Left            =   1425
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   825
      Width           =   390
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
      TabIndex        =   6
      Top             =   285
      Width           =   480
   End
End
Attribute VB_Name = "SupRpt2"
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
If IsDate(xdate1.Text) And IsDate(xDate2.Text) Then
    Select Case publicFlag
    Case 3, 5
        Module1 ' Total invoices for supplier
    Case 7, 8
        Module2 ' invoices for supplier
    Case 9
        module3 ' items for supplier
    Case 11
        Module4 '  Cashed from Supplier
    Case 12
        module5 '  Cashed from Supplier
    Case 13, 14
        Module6 '  Cashed from Supplier
    End Select
End If
End Sub

Private Sub CmdUndo_Click()
xclient.Text = ""
xdate1.Text = ""
xDate2.Text = ""
xCOMP.BoundText = ""
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
If publicFlag = 13 Or publicFlag = 14 Then
    Label1.Visible = False
    xCode.Visible = False
    xCodeName.Visible = False
    xdate1.Top = xCode.Top
    xDate2.Top = xCode.Top
    Label3.Top = xCode.Top
    Label4.Top = xCode.Top
End If
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
xCodeName.Caption = IIf(SuplierTable.NoMatch, "", SuplierTable.DESCA)
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
Private Sub Module1() ' Total invoices for supplier
Dim sourcetable As Recordset
Dim TargetTable As Recordset
If publicFlag = 3 Then
    cString = " Select File4_11.Doc_ID,File4_11.SAL,File4_11.PAY,File4_11.Date" & _
              " From file4_11 " & _
              " Where File4_11.[Type] = '4' " & _
              " And [Date] Between " & DateSql(xdate1.Text) & _
              " And " & DateSql(xDate2.Text) & _
              " And File4_11.Code = " & MyParn(xCode.Text)
ElseIf publicFlag = 5 Then
    cString = " Select File4_11.Doc_ID,File4_11.SAL,File4_11.PAY,File4_11.Date" & _
              " From file4_11 " & _
              " Where File4_11.[Type] = '5' " & _
              " And [Date] Between " & DateSql(xdate1.Text) & _
              " And " & DateSql(xDate2.Text) & _
              " And File4_11.Code = " & MyParn(xCode.Text)
End If
tempdb.Execute "DELETE * FROM TEMP"
Set sourcetable = mydb.OpenRecordset(cString)
Set TargetTable = tempdb.CreateDynaset("TEMP")
If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = sourcetable!Doc_Id
    TargetTable.str2 = IIf(publicFlag = 3, " ›Ê« Ì— „‘ —Ì« ", "›Ê« Ì— „—œÊœ „‘ —Ì« ") & " ··„Ê—œ " & xCodeName.Caption
    TargetTable.str4 = " „‰ " & xdate1.Text & " ≈·Ï " & xDate2.Text
    TargetTable.val1 = IIf(publicFlag = 3, sourcetable.sal, sourcetable.PAY)
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = sourcetable.Date
'    TargetTable.str9 = firsttitle
'    TargetTable.str10 = Secondtitle

    TargetTable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\Supler04.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
Function RetFirstBalance()
Dim FirstBalanceTAble As Recordset

cField = myiif("", "Iif(IsNull(SAL),0,SAL)-Iif(IsNull(PAY),0,PAY)") & _
        " as firstBalance "

cString = "Select " & cField & _
          "From File4_11" & _
          " Where Code = " & MyParn(xCode.Text) & " And " & _
          " Date < " & DateSql(xdate1.Text) & _
          " Group by code"
Set FirstBalanceTAble = mydb.OpenRecordset(cString, dbOpenSnapshot)
RetFirstBalance = IIf(FirstBalanceTAble.RecordCount = 0, 0, FirstBalanceTAble.FirstBalance)
End Function
Sub Module2()   ' invoices for supplier
Dim sourcetable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub

tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
If publicFlag = 7 Then
    cString = "Select file7_20.*,File1_10.DescA" & _
              " From file7_20 inner join file1_10 on file7_20.item = file1_10.item " & _
              " Where CODE = " & MyParn(xCode.Text) & _
              " and Date between DateValue(" & MyParn(xdate1.Text) & ")" & _
              " and DateValue(" & MyParn(xDate2.Text) & ")" & _
              " order by Date,doc_no "
Else
    cString = "Select FILE6_11.*,File1_10.DescA" & _
              " From FILE6_11 inner join file1_10 on FILE6_11.item = file1_10.item " & _
              " Where CODE = " & MyParn(xCode.Text) & _
              " and Date between DateValue(" & MyParn(xdate1.Text) & ")" & _
              " and DateValue(" & MyParn(xDate2.Text) & ")" & _
              " order by Date,doc_no "
End If
Set sourcetable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì—"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = sourcetable.CODE
    TargetTable.str3 = sourcetable!doc_no
    TargetTable.str4 = sourcetable.Item
'    TargetTable.str5 = sourcetable![FILE1_10.DESCA]
    TargetTable.val4 = sourcetable.Quant
'    TargetTable.val5 = sourcetable.Number
    TargetTable.STR7 = IIf(publicFlag = 7, " ›’Ì·Ï „‘ —Ì« ", " ›’Ì·Ï „—œÊœ „‘ —Ì« ") & "  " & xCodeName.Caption
    TargetTable.str2 = " „‰ " & xdate1.Text & " ≈·Ï " & xDate2.Text
    If bopt2 Then
        TargetTable.val1 = sourcetable.price
        TargetTable.val3 = sourcetable.total
    End If
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = sourcetable!Date
'    TargetTable.str9 = firsttitle
'    TargetTable.str10 = Secondtitle
    TargetTable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\Supler05.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
Sub module3()
Dim sourcetable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cField1 = myiif( _
               "FILE1_11.Type = '2'", "[IN]") & _
               " As PurQuant"

cField3 = myiif( _
               "FILE1_11.Type = '7'", "[OUT]") & _
               " As RetPurQuant"


cField5 = myiif( _
               "FILE1_11.Type = '2'", "[Total]") & _
               " As PurValue"

cField6 = myiif( _
               "FILE1_11.Type = '7'", "[Total]") & _
               " As RetPurValue"

cString = "Select File1_11.Item as Item,File1_11.Code," & _
          "File1_10.DescA,File1_10.Group," & _
          "file1_50.DescA," & _
          cField1 & "," & cField3 & "," & _
          cField5 & "," & cField6 & _
          " From ((File1_11 Inner Join File1_10 on File1_11.Item = File1_10.Item)" & _
          " Inner Join File4_10 on File1_11.Code = File4_10.Code)" & _
          " Inner Join file1_50 on File1_10.Group = File1_50.Code" & _
         " Where Date >= " & DateSql(xdate1.Text) & _
         " and Date <= " & DateSql(xDate2.Text) & _
          " And File1_11.Code = " & MyParn(xCode.Text) & _
          " And (file1_11.Type = '2' or file1_11.Type = '7') " & _
          " Group By " & _
          " File1_11.Item,File1_11.Code," & _
          "File1_10.DescA,File1_10.[Group]," & _
          "File1_50.DescA"
          
Set sourcetable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With sourcetable
If .RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = sourcetable!CODE
    TargetTable.str2 = "≈Ã„«·Ï „‘ —Ì«  «·√’‰«› „‰  " & xCodeName.Caption
    TargetTable.str6 = " „‰ " & xdate1.Text & " ≈·Ï " & xDate2.Text
    TargetTable.str3 = sourcetable!Item
    TargetTable.str4 = sourcetable![FILE1_10.DESCA]
    TargetTable.val1 = .PurQuant
    TargetTable.val2 = .RetPurQuant
    TargetTable.val3 = !PurQuant - !RetPurQuant
    TargetTable.Str11 = ![file1_50.desca]
    
    TargetTable.val7 = .Purvalue
    TargetTable.val8 = .retPurvalue
    TargetTable.Val9 = .Purvalue - .retPurvalue
'    TargetTable.str9 = firsttitle
'    TargetTable.str10 = Secondtitle
    
    
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\supler06.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End With
End Sub
Sub Module4() ' Payment to suppliers
Dim sourcetable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "Select Doc_No,Date,inv_No,DescA,Value " & _
          " from file8_20 " & _
          " Where Code = " & MyParn(xCode.Text) & " and " & _
          " Date Between DateValue(" & MyParn(xdate1.Text) & ") and " & _
          " DateValue(" & MyParn(xDate2.Text) & ")" & _
          " Order by Date"
          
Set sourcetable = mydb.OpenRecordset(cString)
If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = sourcetable.doc_no
    TargetTable.str2 = sourcetable.inv_No
    TargetTable.str3 = sourcetable.DESCA
    TargetTable.str4 = xCodeName.Caption
    
    TargetTable.val1 = sourcetable!Value
    TargetTable.Date1 = Format(sourcetable!Date, "dd-mm-yy")
    TargetTable.date2 = xdate1.Text
    TargetTable.date3 = xDate2.Text
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle

    TargetTable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\Client11.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
Sub module5()
Dim sourcetable As Recordset
Dim TargetTable As Recordset
Dim FirstBalance As Double

If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
cString = "Select File4_11.*," & _
           "File4_10.DescA" & _
          " From File4_11 inner Join file4_10 on file4_10.Code = File4_11.Code" & _
          " Where File4_11.Code = " & MyParn(xCode.Text) & " And " & _
          " Date Between " & DateSql(xdate1.Text) & _
          " and " & DateSql(xDate2.Text) & _
          " Order By Date,Sal "

Set sourcetable = mydb.OpenRecordset(cString, dbOpenSnapshot)
Set TargetTable = tempdb.OpenRecordset("TEMP")

If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If

FirstBalance = RetFirstBalance
If FirstBalance <> 0 Then
    TargetTable.AddNew
'    TargetTable.str1 = xclient.Text
    TargetTable.str2 = sourcetable![FILE4_10.DESCA]
    TargetTable.str4 = "—’Ìœ " & Format(xdate1.Text, "dd-mm-yy")

    If FirstBalance > 0 Then
        TargetTable.val1 = Abs(FirstBalance)
    ElseIf FirstBalance < 0 Then
         TargetTable.val2 = Abs(FirstBalance)
    End If
    
    TargetTable.val3 = FirstBalance
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = xdate1.Text
    TargetTable.Update
End If

Do
    TargetTable.AddNew
    FirstBalance = FirstBalance + TurnValue(sourcetable.sal, Null, 0) - TurnValue(sourcetable.PAY, Null, 0)
    TargetTable.str1 = sourcetable.CODE
    TargetTable.str2 = sourcetable.[FILE4_10.DESCA]
    TargetTable.str3 = sourcetable.Doc_Id
    TargetTable.str4 = sourcetable.[FILE4_11.DESCA]
    
    TargetTable.val1 = TurnValue(sourcetable.sal, Null, 0)
    TargetTable.val2 = TurnValue(sourcetable.PAY, Null, 0)
    TargetTable.val3 = TurnValue(FirstBalance, Null, 0)
    
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = Format(sourcetable!Date, "dd-mm-yy")
'    TargetTable.str9 = firsttitle
'    TargetTable.str10 = Secondtitle

    TargetTable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\Supler07.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
Sub Module6()
Dim sourcetable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
If publicFlag = 13 Then
    cString = "Select File7_20.DOC_NO,File7_20.CODE,File7_20.DATE," & _
              "File4_10.DescA," & _
              "Sum(File7_20.Total) as SumOfTotal" & _
              " From File7_20 Inner Join File4_10 On File7_20.CODE = File4_10.Code" & _
              " Where DATE Between " & DateSql(xdate1.Text) & _
              " And " & DateSql(xDate2.Text) & _
              " Group By File7_20.DOC_NO,File7_20.CODE,File7_20.DATE," & _
              " File4_10.DescA"
ElseIf publicFlag = 14 Then
    cString = "Select FILE6_11.DOC_NO,FILE6_11.CODE,FILE6_11.DATE," & _
              "File4_10.DescA," & _
              "Sum(FILE6_11.Total) as SumOfTotal" & _
              " From FILE6_11 Inner Join File4_10 On FILE6_11.CODE = File4_10.Code" & _
              " Where DATE Between " & DateSql(xdate1.Text) & _
              " And " & DateSql(xDate2.Text) & _
              " Group By FILE6_11.DOC_NO,FILE6_11.CODE,FILE6_11.DATE," & _
              " File4_10.DescA"
End If
Set sourcetable = mydb.OpenRecordset(cString)
If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = sourcetable.doc_no
    TargetTable.str2 = sourcetable.CODE
    TargetTable.str3 = sourcetable.DESCA
    TargetTable.str4 = IIf(publicFlag = 13, "«Ã„«·Ï ›Ê« Ì— „‘ —Ì«  ", "≈Ã„«·Ï ›Ê« Ì— „—œÊœ „‘ —Ì« ")
    
    TargetTable.val1 = sourcetable.sumofTotal
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = Format(sourcetable!Date, "dd-mm-yy")
'    TargetTable.str9 = firsttitle
'    TargetTable.str10 = Secondtitle

    TargetTable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
myws.BeginTrans
myws.CommitTrans
REPORT1.ReportFileName = PublicPath & "\Reports\Supler08.RPT"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
