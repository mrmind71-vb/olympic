VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form SUPLReport3 
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   2115
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
   ScaleHeight     =   2115
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
   Begin VB.TextBox xclient 
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
      Top             =   1275
      Width           =   3390
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   225
         Width           =   1065
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   390
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   1065
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2175
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
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
   Begin VB.Label ClientName 
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
Attribute VB_Name = "SUPLReport3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClientTable As Recordset
Dim TempTable As Recordset
Sub clientCash()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "SELECT FILE4_10.CODE, FILE8_30.DOC_NO , FILE4_10.DESCA AS CustName , FILE8_30.DATE, FILE8_30.desca, FILE8_30.VALUE, FILE8_30.doc FROM FILE4_10 RIGHT JOIN FILE8_30 ON FILE4_10.CODE = FILE8_30.CODE " & _
          "WHERE FILE8_30.code is not null "
If xclient.Text <> "" Then cString = cString & " AND FILE8_30.CODE = " & MyParn(xclient.Text)
If IsDate(xdate1.Text) Then cString = cString & " AND FILE8_30.DATE >= " & DateSql(xdate1.Text)
If IsDate(xDate2.Text) Then cString = cString & " AND FILE8_30.DATE <= " & DateSql(xDate2.Text)
cString = cString & " GROUP BY FILE4_10.CODE, FILE8_30.DOC_NO ,FILE4_10.DESCA, FILE8_30.DATE, FILE8_30.desca, FILE8_30.VALUE, FILE8_30.doc "
          
Set SourceTable = mydb.OpenRecordset(cString)
If SourceTable.RecordCount > 0 Then
Do
    
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.CUSTNAME
    TargetTable.str2 = SourceTable.doc_no
    TargetTable.str4 = SourceTable.DESCA
    TargetTable.str3 = SourceTable.doc
    
    TargetTable.VAL1 = SourceTable.Value
    TargetTable.Date1 = SourceTable.Date
    TargetTable.date2 = IIf(IsDate(xdate1.Text), xdate1.Text, Null)
    TargetTable.date3 = IIf(IsDate(xDate2.Text), xDate2.Text, Null)
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End If

cString = " SELECT file5_23.DATE, FILE5_21.DATE_1, FILE5_21.SER_NO, file5_23.VALUE, FILE5_21.NAME1 FROM FILE5_21 INNER JOIN file5_23 ON FILE5_21.SER_NO = file5_23.ser_no " & _
          " WHERE FILE5_21.code is not null "
If xclient.Text <> "" Then cString = cString & " AND FILE5_21.CODE = " & MyParn(xclient.Text)
If IsDate(xdate1.Text) Then cString = cString & " AND FILE5_23.DATE >= " & DateSql(xdate1.Text)
If IsDate(xDate2.Text) Then cString = cString & " AND FILE5_23.DATE <= " & DateSql(xDate2.Text)
          
Set SourceTable = mydb.OpenRecordset(cString)
If SourceTable.RecordCount > 0 Then
Do
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.name1
    TargetTable.str2 = SourceTable.ser_no
    TargetTable.str4 = " √Ê—«ﬁ œ›⁄ Õﬁ " & Format(SourceTable.date_1, "DD-MM-YYYY")
    
    TargetTable.VAL2 = SourceTable.Value
    TargetTable.Date1 = SourceTable.Date
    TargetTable.date2 = IIf(IsDate(xdate1.Text), xdate1.Text, Null)
    TargetTable.date3 = IIf(IsDate(xDate2.Text), xDate2.Text, Null)
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF

End If
myws.BeginTrans
myws.CommitTrans

Report1.ReportFileName = PublicPath & "\Reports\Client11.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Sub clientitems()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")

CFIELD1 = myiif( _
               "Type = '6'  ", "[Out]") & _
               " As SalesQuant,"


cField3 = myiif( _
               "Type = '3'", "[in]") & _
               " As RetSalesQuant,"

cField5 = myiif( _
               "Type = '6'   ", "[Total]") & _
               " As SalesValue,"

cField6 = myiif( _
               "Type = '3'", "[Total]") & _
               " As RetSalesValue"

cString = "Select FILE1_10.PACK , File1_11.Item ,File1_11.Code," & _
          "File1_10.DescA,File1_10.[Group]," & _
          "File4_10.DescA," & _
          "File1_50.DescA," & _
          CFIELD1 & cField3 & cField5 & cField6 & _
          " From ((File1_11 Inner Join File1_10 on File1_11.Item = File1_10.Item)" & _
                " Inner Join File4_10 on File1_11.Code = File4_10.Code) " & _
                " Inner Join file1_50 On File1_10.[Group] = File1_50.Code WHERE FILE1_10.ITEM IS NOT NULL "
If xclient.Text <> "" Then cString = cString & " AND FILE1_11.CODE = " & MyParn(xclient.Text)
If IsDate(xdate1.Text) Then cString = cString & " AND FILE1_11.DATE >= " & DateSql(xdate1.Text)
If IsDate(xDate2.Text) Then cString = cString & " AND FILE1_11.DATE <= " & DateSql(xDate2.Text)
cString = cString & "Group By " & _
          " File1_11.item,File1_11.Code,FILE1_10.PACK , " & _
          "File1_10.DescA,File1_10.[Group]," & _
          "File4_10.DescA," & _
          "File1_50.DescA"
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
If .RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
Do
    If SourceTable.SALESQUANT > 0 Then
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.[file1_50.DescA]
    TargetTable.str2 = ": ’«›Ï „‘ —Ì«   «·√’‰«› ··„Ê—œ " & SourceTable.[file4_10.DescA]
    TargetTable.str3 = SourceTable.Item
    TargetTable.str4 = SourceTable.[FILE1_10.DESCA]
    TargetTable.Str11 = SALESQUANT
    TargetTable.str12 = RetSalesQuant
    TargetTable.STR13 = TurnValue(.SALESQUANT, Null, 0) - TurnValue(.RetSalesQuant, Null, 0)
    
    TargetTable.VAL4 = .SALESQUANT
    TargetTable.VAL5 = .RetSalesQuant
    TargetTable.VAL6 = .SALESQUANT - .RetSalesQuant
    
    TargetTable.str9 = ": ’«›Ï „‘ —Ì«   «·√’‰«› ··„Ê—œ "
    TargetTable.VAL7 = .salesvalue
    TargetTable.VAL8 = .retsalesvalue
    TargetTable.VAL9 = .salesvalue - .retsalesvalue
    
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle

    TargetTable.Update
    End If
    SourceTable.MoveNext
Loop Until SourceTable.EOF
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\Client8.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End With
End Sub
Sub ClientMove()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
Dim FirstBalance As Double

If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
cString = "Select File4_11.*," & _
           "File4_10.DescA" & _
          " From File4_11 inner Join file4_10 on file4_10.Code = File4_11.Code" & _
          " Where File4_11.Code = " & MyParn(xclient.Text) & " And " & _
          " Date >= DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
          " Order By Date,Pay "

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
Set TargetTable = tempdb.OpenRecordset("TEMP")

If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If

FirstBalance = RetFirstBalance
If FirstBalance <> 0 Then
    TargetTable.AddNew
    TargetTable.str1 = xclient.Text
    TargetTable.str2 = SourceTable.[file4_10.DescA]
    TargetTable.str4 = "—’Ìœ " & xdate1.Text

    If FirstBalance > 0 Then
        TargetTable.VAL2 = Abs(FirstBalance)
    ElseIf FirstBalance < 0 Then
         TargetTable.VAL1 = Abs(FirstBalance)
    End If
    
    TargetTable.VAL3 = FirstBalance
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = xdate1.Text
    TargetTable.Update
End If

Do
    TargetTable.AddNew
    FirstBalance = FirstBalance + TurnValue(SourceTable.sal, Null, 0) - TurnValue(SourceTable.PAY, Null, 0)
    TargetTable.str1 = SourceTable.CODE
    TargetTable.str2 = SourceTable.[file4_10.DescA]
    TargetTable.str3 = SourceTable.DOC_ID
    TargetTable.str4 = SourceTable.[FILE4_11.DESCA]
    
    TargetTable.VAL1 = TurnValue(SourceTable.PAY, Null, 0)
    TargetTable.VAL2 = TurnValue(SourceTable.sal, Null, 0)
    TargetTable.VAL3 = TurnValue(FirstBalance, Null, 0)
    
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = SourceTable.Date
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\Client10.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Function CountBalance(Client)
Dim ClientMove As Recordset
NRETURN = 0
Set ClientMove = mydb.CreateDynaset("select * from file4_11 where code = " & MyParn(Client))
If ClientMove.RecordCount = 0 Then
    CountBalance = 0
    Exit Function
End If
ClientMove.MoveFirst
Do
    NRETURN = NRETURN + TurnValue(ClientMove.sal, Null, 0) - TurnValue(ClientMove.PAY, Null, 0)
    ClientMove.MoveNext
Loop Until ClientMove.EOF
CountBalance = NRETURN
End Function
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Private Sub CmdApply_Click()
Select Case publicFlag
Case 1, 2
    ClientInvtotal
Case 3, 4
    Salesinv
Case 5
   clientitems
Case 6
    clientCash
Case 7
    ClientMove
End Select
End Sub
Sub Salesinv()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub

tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.OpenRecordset("TEMP")
If publicFlag = 3 Then
    cString = "Select file7_20.*,File1_10.DescA,File1_10.PACK " & _
              " From file7_20 inner join file1_10 on file7_20.item = file1_10.item " & _
              " Where CODE = " & MyParn(xclient.Text) & _
              " and Date between DateValue(" & MyParn(xdate1.Text) & ")" & _
              " and DateValue(" & MyParn(xDate2.Text) & ")" & _
              " order by Date,DOC_NO "
Else
    cString = "Select File6_11.*,File1_10.DescA ,File1_10.PACK " & _
              " From File6_11 inner join file1_10 on File6_11.ITEM = file1_10.item " & _
              " Where CODE = " & MyParn(xclient.Text) & _
              " and Date between DateValue(" & MyParn(xdate1.Text) & ")" & _
              " and DateValue(" & MyParn(xDate2.Text) & ")" & _
              " order by Date,DOC_NO "
End If
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì—"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.CODE
    TargetTable.str2 = ClientName.Caption
    TargetTable.str3 = SourceTable.doc_no
    TargetTable.str4 = SourceTable.Item
    TargetTable.STR5 = SourceTable.DESCA
    TargetTable.str7 = IIf(publicFlag = 3, " ﬁ—Ì—  ›’Ì·Ï ›Ê« Ì— „‘ —Ì«  ", " ﬁ—Ì—  ›’Ì·Ï „—œÊœ« ")
    
    TargetTable.VAL1 = SourceTable.price
    TargetTable.VAL3 = SourceTable.total
    TargetTable.str6 = SourceTable.Quant

   
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = SourceTable.Date
    
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\report3.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Sub RetSalesinv()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cCondition = IIf(xclient.Text = "", " Where ", " Where Client = " & MyParn(xclient.Text) & " and ")
cString = "Select file6_11.*,File1_10.DescA " & _
          "From file6_11 inner join file1_10 on file6_11.itemCode = file1_10.item " & _
           cCondition & _
          " invDate between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")"
If xCOMP.BoundText <> "" Then cString = cString & " and comp = " & MyParn(xCOMP.BoundText)
cString = cString & " order by invDate,Code "

Set SourceTable = mydb.CreateSnapshot(cString)
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì—"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.Client
    TargetTable.str2 = ClientName.Caption
    TargetTable.str3 = SourceTable.CODE
    TargetTable.str4 = SourceTable.ItemCode
    TargetTable.STR5 = SourceTable.DESCA
    TargetTable.str6 = SourceTable.Quant1
    TargetTable.str7 = SourceTable.dam2
    TargetTable.str8 = IIf(SourceTable.Expire, "„‰ ÂÏ «·’·«ÕÌ…", "€Ì— „‰ ÂÏ «·’·«ÕÌ…")
    
    TargetTable.VAL1 = SourceTable.price
    TargetTable.VAL2 = SourceTable.discount
    TargetTable.VAL3 = SourceTable.total
   
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = SourceTable.invDate
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\report4.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub

Private Sub CmdUndo_Click()
xclient.Text = ""
xdate1.Text = ""
xDate2.Text = ""
'xCOMP.BoundText = ""
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
Set ClientTable = mydb.OpenRecordset("SELECT * FROM file4_10")
Set TempTable = tempdb.OpenRecordset("TEMP")
Set FlagTable = mydb.OpenRecordset("File1_70")
End Sub
Function MYVALID()
If Not (IsDate(xdate1.Text) And IsDate(xDate2.Text)) Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If DateValue(xdate1.Text) > DateValue(xDate2.Text) Then
    MsgBox "«· «—ÌŒ «·√Ê· √ﬂ»— „‰ «· «—ÌŒ «·À«‰Ï"
    Exit Function
End If
MYVALID = True
End Function
Private Sub xClient_Change()
If xclient.Text = "" Then Exit Sub
ClientTable.FindFirst " CODE = " & MyParn(xclient.Text)
If ClientTable.NoMatch Then Exit Sub
ClientName.Caption = ClientTable.DESCA
End Sub
Private Sub xclient_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xclient.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„ From File4_10"
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1200
    GrdArray(2) = 4000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub ClientInvtotal()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub
If publicFlag = 1 Then
    cString = "Select File7_20.DOC_NO, " & _
              "File7_20.DATE," & _
              "Sum(File7_20.Total) as sumTotal " & _
              "From file7_20 " & _
              "Where file7_20.CODE = " & MyParn(xclient.Text) & _
              " And FILE7_20.DATE Between DateValue(" & MyParn(xdate1.Text) & ") and DateValue(" & MyParn(xDate2.Text) & ")" & _
              " Group by DOC_NO,FILE7_20.DATE"
Else
    cString = "Select File6_11.DOC_NO," & _
              "File6_11.DATE," & _
              "Sum(File6_11.Total) as sumTotal " & _
              "From file6_11 " & _
              "Where file6_11.CODE = " & MyParn(xclient.Text) & _
              " And DATE Between DateValue(" & MyParn(xdate1.Text) & ") and DateValue(" & MyParn(xDate2.Text) & ")" & _
              " Group by DOC_NO,DATE"
End If
tempdb.Execute "DELETE * FROM TEMP"
Set SourceTable = mydb.CreateSnapshot(cString)
Set TargetTable = tempdb.CreateDynaset("TEMP")
If SourceTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
Do
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.doc_no
    TargetTable.str2 = ClientName.Caption
    TargetTable.str4 = IIf(publicFlag = 1, "„ «»⁄… √Ã„«·Ï „‘ —Ì«  ", "„ «»⁄… «Ã„«·Ï „—œÊœ« ")
    
    TargetTable.VAL1 = SourceTable.SUMTOTAL
    TargetTable.Date1 = xdate1.Text
    TargetTable.date2 = xDate2.Text
    TargetTable.date3 = SourceTable.Date
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\Comp4.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Function RetFirstBalance()
Dim FirstBalanceTAble As Recordset

cField = myiif("", "Iif(IsNull(Sal),0,Sal)-Iif(IsNull(Pay),0,Pay)") & _
        " as firstBalance "

cString = "Select " & cField & _
          "From File4_11" & _
          " Where Code = " & MyParn(xclient.Text) & " And " & _
          "  Date < DateValue(" & MyParn(xdate1.Text) & ")" & _
          " Group by File4_11.code"
Set FirstBalanceTAble = mydb.OpenRecordset(cString, dbOpenSnapshot)
RetFirstBalance = IIf(FirstBalanceTAble.RecordCount = 0, 0, FirstBalanceTAble.FirstBalance)
End Function
