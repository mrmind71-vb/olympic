VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form RepItem1 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
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
   ScaleHeight     =   2145
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   " ÃœÌœ"
      Height          =   390
      Left            =   1425
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1695
      Width           =   1440
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "«” Ã«»…"
      Height          =   390
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1695
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1695
      Width           =   1365
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -825
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   750
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBCtls.DBCombo xStore1 
      Bindings        =   "Repitm1.frx":0000
      Height          =   315
      Left            =   2625
      TabIndex        =   1
      Top             =   600
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin Crystal.CrystalReport REPORT1 
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
   Begin MSDBCtls.DBCombo xStore2 
      Bindings        =   "Repitm1.frx":0014
      Height          =   315
      Left            =   2625
      TabIndex        =   2
      Top             =   1050
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï „Œ“‰"
      Height          =   315
      Left            =   4875
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1125
      Width           =   765
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Õ Ï ÌÊ„ :"
      Height          =   315
      Left            =   4875
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   150
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰ „Œ“‰ "
      Height          =   240
      Left            =   4875
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   675
      Width           =   690
   End
End
Attribute VB_Name = "RepItem1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StockTable As Recordset, uStockTable As Recordset
Dim storeTable As Recordset
Dim cStore As String
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
Select Case publicFlag
Case 1
    RepItem1
Case 3
    RepItem2
Case 8
    RepItem3
End Select
End Sub
Sub RepItem1()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Set TempTable = tempdb.OpenRecordset("Temp")

If xStore2.BoundText <> "" Then
    cStoreString = " and cNo Between " & _
                   MyParn(xStore1.BoundText) & " and " & _
                   MyParn(xStore2.BoundText)
    cReportTitle = "„‰ " & xStore1.Text & _
                   "≈·Ï " & xStore2.Text
Else
    cStoreString = IIf(xStore1.BoundText <> "", " and cNo = " & MyParn(xStore1.BoundText), "")
    cReportTitle = IIf(xStore1.BoundText <> "", xStore1.Text, "")
End If


cString = "Select " & _
          "File1_11.item" & " As MidOfItem," & _
          "Sum(iif(IsNull([in]),0,[in]))-Sum(iif(IsNull([out]),0,[out])) As Quant," & _
          "First(File1_10.DescA) as firstOfDescA ,File1_10.pack" & _
          " From File1_11 inner join File1_10 on File1_11.item = file1_10.item " & _
          " where cDate <= DateValue(" & MyParn(xDate.Text) & ")" & _
          cStoreString & _
          " Group By " & _
          "File1_11.item," & _
          "File1_10.Pack"
          
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
        If .Quant <> 0 Then
            TempTable.AddNew
            TempTable.str1 = .MidofItem
            TempTable.str2 = .FirstofDescA
            TempTable.str3 = .Quant
            TempTable.str4 = cReportTitle
            TempTable.Date1 = xDate.Text
            TempTable.str8 = firsttitle
            TempTable.str9 = Secondtitle
            TempTable.Update
        End If
       .MoveNext
    Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm1.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Sub RepItem2()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Set TempTable = tempdb.OpenRecordset("Temp")

If xStore2.BoundText <> "" Then
    cStoreString = " and cNo Between " & _
                   MyParn(xStore1.BoundText) & " and " & _
                   MyParn(xStore2.BoundText)
    cReportTitle = "„‰ " & xStore1.Text & _
                   "≈·Ï " & xStore2.Text
Else
    cStoreString = IIf(xStore1.BoundText <> "", " and cNo = " & MyParn(xStore1.BoundText), "")
    cReportTitle = IIf(xStore1.BoundText <> "", xStore1.Text, "")
End If


cString = "Select " & _
          "File1_11.item " & " As MidOfItem," & _
          "Sum(iif(IsNull([in]),0,[in]))-Sum(iif(IsNull([out]),0,[out])) As Quant," & _
          "File1_10.DescA ,File1_10.pack,File1_10.Cost" & _
          " From File1_11 inner join File1_10 on File1_11.item = file1_10.item " & _
          " where cDate <= DateValue(" & MyParn(xDate.Text) & ")" & _
          cStoreString & _
          " Group By " & _
          "File1_11.item," & _
          "File1_10.DescA,File1_10.Pack,File1_10.Cost"
          
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
        If .Quant <> 0 Then
            TempTable.AddNew
            TempTable.str1 = .MidofItem
            TempTable.str2 = .DESCA
            TempTable.str3 = .Quant
            TempTable.str4 = cReportTitle
            
            TempTable.val1 = .COST
            TempTable.VAL2 = .Quant / .Pack * .COST
            TempTable.Date1 = xDate.Text
            TempTable.str19 = firsttitle
            TempTable.str20 = Secondtitle
            TempTable.Update
        End If
       .MoveNext
    Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItm3.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Sub RepItem3()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Set TempTable = tempdb.OpenRecordset("Temp")
If xStore2.BoundText <> "" Then
    cStoreString = " and cNo Between " & _
                   MyParn(xStore1.BoundText) & " and " & _
                   MyParn(xStore2.BoundText)
    cReportTitle = "„‰ " & xStore1.Text & _
                   "≈·Ï " & xStore2.Text
Else
    cStoreString = ""
    cReportTitle = ""
End If

cString = "Select File1_11.Item ," & _
          "Sum(iif(IsNull([in]),0,[in]))-Sum(iif(IsNull([out]),0,[out])) As Quant," & _
          "File1_10.DescA ,File1_10.pack,File1_10.Cost,File1_10.E_date" & _
          " From File1_11 inner join File1_10 on File1_11.item = file1_10.item " & _
          " where E_Date <= DateValue(" & MyParn(xDate.Text) & ")" & _
          cStoreString & _
          " Group By " & _
          "File1_11.item," & _
          "File1_10.DescA ,File1_10.pack,File1_10.Cost,File1_10.E_date"

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
        If .Quant <> 0 Then
            TempTable.AddNew
            TempTable.str1 = .Item
            TempTable.str2 = .DESCA
            TempTable.str3 = .Quant
            TempTable.str4 = cReportTitle
            TempTable.val1 = .COST
            TempTable.VAL2 = .Quant * .COST / .Pack
            TempTable.str10 = IIf(Expire, "» «—ÌŒ «·’·«ÕÌ…", "»œÊ‰  «—ÌŒ «·’·«ÕÌ…")
            TempTable.Date1 = xDate.Text
            TempTable.date2 = .E_Date
            TempTable.str19 = firsttitle
            TempTable.str20 = Secondtitle
            TempTable.Update
        End If
       .MoveNext
    Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\RepItem9.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Command1_Click()
xDate.Text = ""
xStore1.Text = ""
xStore2.Text = ""
Expire.Value = False
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set StockTable = mydb.OpenRecordset("File1_30")
Set storeTable = mydb.OpenRecordset("File1_30")
storeTable.Index = "nfile1_30_1"
Data1.DatabaseName = MdbPath
Data1.RecordSource = "Select * From File1_70 where flag = 1 Order by Code"
xStore1.BoundColumn = "Code"
xStore1.ListField = "Desca"
Data2.DatabaseName = MdbPath
Data2.RecordSource = "Select * From File1_70 where flag = 1 Order by Code"
xStore2.BoundColumn = "Code"
xStore2.ListField = "Desca"

End Sub
Private Sub Store_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xStore1.BoundText = ""
End Sub
Private Function MYVALID()
If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If xStore1.BoundText = "" And xStore2.BoundText <> "" Then
    MsgBox "·« Ì’Õ «œ—«Ã «·„Œ“‰ «·À«‰Ï »œÊ‰ «·√Ê·"
    Exit Function
End If
MYVALID = True
End Function
Private Sub xStore2_Click(Area As Integer)
If KeyCode = 46 Then xStore2.BoundText = ""
End Sub
