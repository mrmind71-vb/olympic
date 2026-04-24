VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Item16 
   Caption         =   "ÊÞÇÑíÑ ÇáÃÕäÇÝ"
   ClientHeight    =   2235
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
   ScaleHeight     =   2235
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
      Left            =   4125
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   1725
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
      Left            =   4575
      Top             =   1275
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
      TabIndex        =   6
      Top             =   1335
      Width           =   3765
      Begin VB.CommandButton CmdClear 
         Caption         =   "ÊÌÏíÏ"
         Height          =   390
         Left            =   1200
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "ÎÑæÌ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "ÚÑÖ"
         Height          =   390
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1140
      End
   End
   Begin MSDBCtls.DBCombo xStore1 
      Bindings        =   "R_Item16.frx":0000
      Height          =   315
      Left            =   735
      TabIndex        =   2
      Top             =   840
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáãÎÒä"
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
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   960
      Width           =   570
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Åáì ÊÇÑíÎ :"
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
      TabIndex        =   8
      Top             =   515
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ãä ÊÇÑíÎ :"
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
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "R_Item16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim InTable As Recordset
Dim OutTable As Recordset
Dim TransFromTable As Recordset
Dim TransToTable As Recordset
Dim PurchTable As Recordset
Dim SalTable As Recordset
Dim movetable As Recordset
Dim compTable As Recordset
Dim itemTable As Recordset
Dim nOption As Integer
Function MYVALID()
If Not IsDate(xDate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
If xStore1.BoundText = "" Then Exit Function
MYVALID = True
End Function
Private Sub CmdClear_Click()
xDate1.Text = ""
xDate2.Text = ""
xStore1.BoundText = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xStore1.BoundText = ""
xDate1.Text = ""
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
xStore1.BoundText = ""
xDate1.Text = ""
xDate2.Text = ""
Set movetable = mydb.OpenRecordset("File1_11")
Set itemTable = mydb.OpenRecordset("SELECT * FROM FILE1_10 ORDER BY ITEM ")
Set TempTable = tempdb.OpenRecordset("Temp")
Data1.DatabaseName = MdbPath
Data1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
xStore1.ListField = "Desca"
xStore1.BoundColumn = "code"
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"

cString = "SELECT FILE1_80.item, Sum(FILE1_80.QUANT) AS SumIn, First(FILE1_10.DESCA) AS DescItem " & _
          "FROM FILE1_80 LEFT JOIN FILE1_10 ON FILE1_80.item = FILE1_10.ITEM " & _
          " where Date Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE1_80.item "
Set InTable = mydb.OpenRecordset(cString, dbOpenSnapshot)


With InTable
If .RecordCount > 0 Then
    Do While Not .EOF
        TempTable.AddNew
        TempTable.VAL1 = .SUMIn
        TempTable.str1 = .Item
        TempTable.str2 = .DESCITEM
        TempTable.str7 = " ÅÌãÇáì æÇÑÏ áÃÕäÇÝ áãÎÒä " & xStore1.Text
        TempTable.str8 = " ãä ÊÇÑíÎ " & xDate1.Text & " Åáì ÊÇÑíÎ " & xDate2.Text
        TempTable.str9 = firsttitle
        TempTable.str10 = Secondtitle
        TempTable.Update
        .MoveNext
    Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\R_Item16.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Sub ItemsLookup()
ActiveControl.Text = ""
Dim Generalarray(4)
Dim GrdArray(3)
    
Set Generalarray(1) = Me
Generalarray(2) = "Select Item as ÇáÕäÝ,DescA,pack as [ÇÓã ÇáÕäÝ] From file1_10 as [ÇáÈÚæÉ] "
Generalarray(3) = " Where DescA Like('*cFilter*')"
Generalarray(4) = "Order by Item"
    
GrdArray(1) = 1000
GrdArray(2) = 3500
GrdArray(3) = 1500

Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search.Caption = "ÇÓÊÚáÇã "
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
Function CountDBalance(pItem, pStore, pDate, lValue)
MyQuant = 0
movetable.Seek ">=", pItem, pStore
If myFound(movetable, Array(pItem, pStore)) Then
    If lValue Then
        Do While movetable.Item = pItem
            If movetable.cNo = pStore And DateValue(movetable.CDate) <= DateValue(pDate) And movetable.cType <> "" Then
                MyQuant = MyQuant + TurnValue(movetable.In, Null, 0) - TurnValue(movetable.OUT, Null, 0)
            End If
            movetable.MoveNext
            If movetable.EOF Then Exit Do
        Loop
    Else
        Do While movetable.Item = pItem
            If movetable.cNo = pStore And DateValue(movetable.CDate) < DateValue(pDate) And movetable.cType <> "" Then
                MyQuant = MyQuant + TurnValue(movetable.In, Null, 0) - TurnValue(movetable.OUT, Null, 0)
            End If
            movetable.MoveNext
            If movetable.EOF Then Exit Do
        Loop
    End If
End If
CountDBalance = MyQuant
End Function

