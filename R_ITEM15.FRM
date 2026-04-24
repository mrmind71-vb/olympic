VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Item15 
   Caption         =   " ﬁ«—Ì— «·√’‰«›"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
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
   ScaleHeight     =   1710
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   495
      Width           =   1290
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1290
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   900
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
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   900
      Width           =   3765
      Begin VB.CommandButton CmdClear 
         Caption         =   " ÃœÌœ"
         Height          =   390
         Left            =   1275
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2550
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1140
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "≈·Ï  «—ÌŒ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   515
      Width           =   765
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
      Height          =   240
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "R_Item15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim DataTable As Recordset
Dim nOption As Integer
Function MYVALID()
If Not IsDate(xdate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
MYVALID = True
End Function
Private Sub CmdClear_Click()
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If TypeOf ActiveControl Is DBCombo Then ActiveControl.BoundText = ""
End If
End Sub
Private Sub Form_Load()
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TempTable = tempdb.OpenRecordset("TEMP")
With TempTable
' «·’«œ—
cString = " SELECT FILE1_82.* " & _
          " FROM FILE1_82 " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " ORDER BY DATE,DOC_NO,ITEM "
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
        .Date1 = DataTable!Date
        .str1 = "’«œ—"
        .str2 = DataTable.DOC_NO
        .str3 = "„‰ „Œ“‰ " & DataTable.Store
        .str4 = DataTable.Item
        .STR5 = DataTable.DESCA
        .STR6 = "01"
        .val1 = DataTable.Quant
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

' «·Ê«—œ
cString = " SELECT FILE1_81.* " & _
          " FROM FILE1_81 " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " ORDER BY DATE,DOC_NO,ITEM "
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
        .Date1 = DataTable.Date
        .str1 = "«·Ê«—œ"
        .str2 = DataTable.DOC_NO
        .str3 = "≈·Ï  „Œ“‰ " & DataTable.Store
        .str4 = DataTable.Item
        .STR5 = DataTable.DESCA
        .STR6 = "02"
        .val1 = DataTable.Quant
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

'«·Â«·ﬂ
cString = " SELECT FILE1_80.* " & _
          " FROM FILE1_80 " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " ORDER BY DATE,DOC_NO,ITEM "
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
        .Date1 = DataTable.Date
        .str1 = "«·Â«·ﬂ"
        .str2 = DataTable.DOC_NO
        .str3 = "„‰  „Œ“‰ " & DataTable.Store
        .str4 = DataTable.Item
        .STR5 = DataTable.DESCA
        .STR6 = "03"
        .val1 = DataTable.Quant
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

'«· ÕÊÌ·« 
cString = " SELECT FILE1_60.* " & _
          " FROM FILE1_60 " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " ORDER  BY DATE,DOC_NO,ITEM "
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
        .Date1 = DataTable.Date
        .str1 = " ÕÊÌ·« "
        .str2 = DataTable.DOC_NO
        .str3 = "„‰ " & DataTable.store1 & "≈·Ï " & DataTable.Store2
        .str4 = DataTable.Item
        .STR5 = DataTable.DESCA
        .STR6 = "04"
        .val1 = DataTable.Quant
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

'„»Ì⁄« 
cString = " SELECT FILE6_20.*, FILE1_10.DESCA as itemDescA" & _
          " FROM FILE1_10 RIGHT JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM " & _
          " where [date] Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " ORDER BY DATE,DOC_NO,FILE6_20.ITEM "
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While Not DataTable.EOF
        If DataTable.Store <> "SS" Then
            .AddNew
            .Date1 = DataTable.[Date]
            .str1 = "„»Ì⁄« "
            .str2 = DataTable.DOC_NO
            .str3 = " „‰ „Œ“‰ " & DataTable.Store
            .str4 = DataTable.Item
            .STR5 = DataTable!ItemDescA
            .STR6 = "05"
            .val1 = DataTable.Quant
            .Update
         End If
         DataTable.MoveNext
    Loop
End If

'„‘ —Ì« 
cString = " SELECT FILE7_20.*, FILE1_10.DESCA as itemDesc" & _
          " FROM FILE1_10 RIGHT JOIN FILE7_20 ON FILE1_10.ITEM = FILE7_20.ITEM " & _
          " where [date] Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " ORDER   BY DATE ,Doc_NO,FILE7_20.ITEM "
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
        .Date1 = DataTable.Date
        .str1 = "„‘ —Ì« "
        .str2 = DataTable.DOC_NO
        .str3 = "›Ï „Œ“‰ " & DataTable.Store
        .str4 = DataTable.Item
        .STR5 = DataTable.itemdesc
        .STR6 = "06"
        .val1 = DataTable.Quant
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If


cString = " SELECT FILE8_10.*, FILE3_10.DESCA AS ClientDesc " & _
          " FROM FILE3_10 RIGHT JOIN FILE8_10 ON FILE3_10.CODE = FILE8_10.CODE " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " ORDER   BY date ,doc_no,FILE8_10.code"
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
        .Date1 = DataTable.Date
        .str1 = "„ﬁ»Ê÷«  „‰ ⁄„·«¡"
        .str2 = DataTable.DOC_NO
'       .Str3 = "›Ï „Œ“‰ " & DataTable.store
        .str4 = DataTable.code
        .STR5 = DataTable.ClientDesc
        .STR6 = "07"
        .val1 = DataTable.Value
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

cString = " SELECT FILE8_20.*, FILE4_10.DESCA AS ClientDesc " & _
          " FROM FILE8_20 LEFT JOIN FILE4_10 ON FILE8_20.CODE = FILE4_10.CODE " & _
          " where [date] Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " ORDER   BY [date] ,doc_no,FILE8_20.code"
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
        .Date1 = DataTable.[Date]
        .str1 = "„œ›Ê⁄«  ≈·Ï „Ê—œÌ‰"
        .str2 = DataTable.DOC_NO
'       .Str3 = "›Ï „Œ“‰ " & DataTable.store
        .str4 = DataTable.code
        .STR5 = DataTable.ClientDesc
        .STR6 = "08"
        .val1 = DataTable.Value
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

cString = " SELECT FILE8_30.*, FILE4_10.DESCA AS ClientDesc " & _
          " FROM FILE8_30 LEFT JOIN FILE4_10 ON FILE8_30.CODE = FILE4_10.CODE " & _
          " where [date] Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " ORDER   BY [date] ,doc_no,FILE8_30.code"
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
        .Date1 = DataTable.[Date]
        .str1 = "„ﬁ»Ê÷«  „‰  ≈·Ï „Ê—œÌ‰"
        .str2 = DataTable.DOC_NO
'       .Str3 = "›Ï „Œ“‰ " & DataTable.store
        .str4 = DataTable.code
        .STR5 = DataTable.ClientDesc
        .STR6 = "09"
        .val1 = DataTable.Value
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

cString = " SELECT FILE8_40.*, FILE3_10.DESCA AS ClientDesc " & _
          " FROM FILE3_10 RIGHT JOIN FILE8_40 ON FILE3_10.CODE = FILE8_40.CODE " & _
          " where date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " ORDER   BY date ,doc_no,FILE8_40.code"
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
        .Date1 = DataTable.Date
        .str1 = "„œ›Ê⁄«  „‰ ⁄„·«¡"
        .str2 = DataTable.DOC_NO
'       .Str3 = "›Ï „Œ“‰ " & DataTable.store
        .str4 = DataTable.code
        .STR5 = DataTable.ClientDesc
        .STR6 = "11"
        .val1 = DataTable.Value
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

cString = " SELECT FILE8_50.* " & _
          " FROM FILE8_50 " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " ORDER   BY date,doc_no"
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
        .Date1 = DataTable.Date
        .str1 = "„’—Ê›« "
        .str2 = DataTable.DOC_NO
'       .Str3 = "›Ï „Œ“‰ " & DataTable.store
'       .Str4 = DataTable.code
        .STR5 = DataTable.DESCA
        .STR6 = "12"
        .val1 = DataTable.Value
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

.MoveFirst
End With
If TempTable.RecordCount > 0 Then
    With TempTable
    Do While True
        .Edit
        TempTable.STR7 = " »Ì«‰ »«·»Ì«‰«  «·„”Ã·… „‰  «—ÌŒ " & xdate1.Text & "≈·Ï  «—ÌŒ" & xDate2.Text
        TempTable.str9 = firsttitle
        TempTable.str10 = Secondtitle
        .Update
        .MoveNext
        If .EOF Then Exit Do
    Loop
    End With

    Report1.ReportFileName = PublicPath & "\Reports\R_ITEM15.rpt"
    Report1.DataFiles(0) = App.Path & "\Temp.mdb"
    Report1.Action = 1
Else
    MsgBox "·«  ÊÃœ Õ—ﬂ… „”Ã·… ·Â–… «·› —… "
End If
End Sub
Sub ItemsLookup()
ActiveControl.Text = ""
Dim Generalarray(4)
Dim GrdArray(3)
    
Set Generalarray(1) = Me
Generalarray(2) = "Select Item as «·’‰›,DescA,pack as [«”„ «·’‰›] From file1_10 as [«·»⁄Ê…] "
Generalarray(3) = " Where DescA Like('*cFilter*')"
Generalarray(4) = "Order by Item"
    
GrdArray(1) = 1000
GrdArray(2) = 3500
GrdArray(3) = 1500

Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search.Caption = "«” ⁄·«„ "
Search.Show 1
End Sub
Private Sub xItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ItemsLookup
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
xDesc.Caption = GrdText(Search.Grid1, 1)
Unload Search
End Sub
