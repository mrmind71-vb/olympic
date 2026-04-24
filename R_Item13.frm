VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Item13 
   Caption         =   " ﬁ«—Ì— «·√’‰«›"
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
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   120
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
      Left            =   480
      Top             =   120
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
      Top             =   1335
      Width           =   4845
      Begin VB.CommandButton CmdApply2 
         Caption         =   "⁄—÷ ≈Ã„«·Ï ﬁÌ„…"
         Height          =   390
         Left            =   1297
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1620
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   1020
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷ ≈Ã„«·Ï ﬂ„Ì…"
         Height          =   390
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1620
      End
   End
   Begin MSDBCtls.DBCombo xStore1 
      Bindings        =   "R_Item13.frx":0000
      Height          =   315
      Left            =   720
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
      Caption         =   "«·„Œ“‰"
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
      TabIndex        =   8
      Top             =   960
      Width           =   570
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
Attribute VB_Name = "R_Item13"
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
Dim ITEM2table As Recordset
Dim Itmtable As Recordset
Dim compTable As Recordset
Public itemTable As Recordset
Dim nOption As Integer
Public DataTable As Recordset
Public LastCost As Recordset
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
Private Sub CmdApply2_Click()
If Not MYVALID Then Exit Sub
Dim nT1, nT2, nT3, nT4, nT5, nT6 As Double

Dim LastCost1 As Recordset
Dim LastCost2 As Recordset

cStr1 = " SELECT FILE7_20.ITEM, FILE7_20.Price, FILE7_20.DATE FROM FILE7_20  " & _
        " WHERE Date < DateValue(" & MyParn(xDate1.Text) & ")" & _
        " ORDER BY DATE "
Set LastCost1 = mydb.OpenRecordset(cStr1)

cStr1 = " SELECT FILE7_20.ITEM, FILE7_20.Price, FILE7_20.DATE FROM FILE7_20  " & _
        " WHERE Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
        " ORDER BY DATE "
Set LastCost2 = mydb.OpenRecordset(cStr1)

tempdb.Execute "DELETE * FROM TEMP"

cStr1 = " SELECT FILE1_10.GROUP, FILE1_10.COST AS ITEMCOST , FILE1_10.DESCA, FILE1_10.ITEM, Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT FROM FILE1_10 LEFT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM " & _
        " WHERE STORE   = " & MyParn(xStore1.BoundText) & _
        " and Date < DateValue(" & MyParn(xDate1.Text) & ")" & _
        " GROUP BY FILE1_10.GROUP, FILE1_10.DESCA, FILE1_10.ITEM , FILE1_10.COST ORDER BY FILE1_10.ITEM"
Set itemTable = mydb.OpenRecordset(cStr1)

cStr1 = " SELECT  Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT , ITEM FROM FILE1_11 " & _
        " WHERE STORE   = " & MyParn(xStore1.BoundText) & _
        " and Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
        " GROUP BY FILE1_11.ITEM ORDER BY ITEM"
Set ITEM2table = mydb.OpenRecordset(cStr1)

cString = "SELECT FILE1_81.item, Sum(FILE1_81.TOTAL) AS SumIn " & _
          "FROM FILE1_81 " & _
          " where Date Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE1_81.item "
Set InTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE1_82.item, Sum(FILE1_82.TOTAL) AS SumOut " & _
          "FROM FILE1_82 " & _
          " where Date Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE1_82.item "
Set OutTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE1_60.STORE1, FILE1_60.item, Sum(FILE1_60.TOTAL) AS SumTransFrom " & _
          "FROM FILE1_60 " & _
          " where Date Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE1 = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE1_60.STORE1, FILE1_60.item "
Set TransFromTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE1_60.STORE2, FILE1_60.item, Sum(FILE1_60.TOTAL) AS SumTransTo " & _
          "FROM FILE1_60 " & _
          " where Date Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE2 = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE1_60.STORE2, FILE1_60.item "
Set TransToTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE6_20.STORE, FILE6_20.ITEM, Sum(FILE6_20.QUANT * FILE6_20.COST ) AS SumSal " & _
          "FROM FILE6_20 " & _
          " where DATE Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE6_20.STORE, FILE6_20.ITEM "
Set SalTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE7_20.STORE, FILE7_20.ITEM, Sum(FILE7_20.TOTAL) AS SumPurch " & _
          "FROM FILE7_20 " & _
          " where DATE Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE7_20.STORE, FILE7_20.ITEM "
Set PurchTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE0_10.item, FILE0_10.Store, Sum(FILE0_10.Differ * FILE0_10.COST ) AS SumDiffer " & _
          "FROM FILE0_10 " & _
          " where date Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE = " & MyParn(xStore1.BoundText) & _
          " AND CLOSED = TRUE " & _
          " GROUP BY FILE0_10.item, FILE0_10.Store "
Set compTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

With Itmtable
If .RecordCount > 0 Then
    Do While Not .EOF
        TempTable.AddNew
        
        itemTable.FindFirst " ITEM = " & MyParn(.Item)
        If Not itemTable.NoMatch Then
            LastCost1.FindLast " ITEM = " & MyParn(.Item)
            If LastCost1.NoMatch Then
                TempTable.val1 = (TurnValue(itemTable.SumIn, Null, 0) - TurnValue(itemTable.SUMout, Null, 0)) * .[cost]
            Else
                TempTable.val1 = (TurnValue(itemTable.SumIn, Null, 0) - TurnValue(itemTable.SUMout, Null, 0)) * LastCost1.price
            End If
        End If
        
        ITEM2table.FindFirst " ITEM = " & MyParn(.Item)
        If Not ITEM2table.NoMatch Then
            LastCost2.FindLast " ITEM = " & MyParn(.Item)
            If LastCost2.NoMatch Then
                TempTable.val2 = (TurnValue(ITEM2table.SumIn, Null, 0) - TurnValue(ITEM2table.SUMout, Null, 0)) * .cost
            Else
                TempTable.val2 = (TurnValue(ITEM2table.SumIn, Null, 0) - TurnValue(ITEM2table.SUMout, Null, 0)) * LastCost2.price
            End If
        Else
            TempTable.val2 = 0
        End If
        InTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not InTable.NoMatch Then TempTable.VAL3 = InTable.SumIn
        
        OutTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not OutTable.NoMatch Then TempTable.VAL4 = OutTable.SUMout

        TransFromTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not TransFromTable.NoMatch Then TempTable.VAL5 = TransFromTable.SumTransFrom

        TransToTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not TransToTable.NoMatch Then TempTable.VAL6 = TransToTable.SumTransTO

        SalTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not SalTable.NoMatch Then TempTable.VAL7 = SalTable.SUMSAL

        PurchTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not PurchTable.NoMatch Then TempTable.VAL8 = PurchTable.SumPurch
        
        compTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not compTable.NoMatch Then TempTable.VAL9 = compTable.SumDiffer
        
        
        TempTable.str1 = .Item
        TempTable.str2 = .DESCA

        TempTable.str7 = " »Ì«‰ »≈Ã„«·Ï Õ—ﬂ… «·√’‰«› ·„Œ“‰ " & xStore1.Text
        TempTable.str8 = " „‰  «—ÌŒ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
        TempTable.str9 = firsttitle
        TempTable.str10 = Secondtitle
        TempTable.Update
        .MoveNext
    Loop
End If
End With

TempTable.MoveFirst
If TempTable.RecordCount > 0 Then
    Do While True
        If (TempTable.val1 + TempTable.val2 + TempTable.VAL3 + TempTable.VAL4 + TempTable.VAL5 + TempTable.VAL6 + TempTable.VAL7 + TempTable.VAL8) = 0 Then
            TempTable.Delete
        End If
        TempTable.MoveNext
        If TempTable.EOF Then Exit Do
    Loop
End If

Report1.ReportFileName = PublicPath & "\Reports\R_Item13.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
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

Set TempTable = tempdb.OpenRecordset("Temp")
Set Itmtable = mydb.OpenRecordset("SELECT * FROM FILE1_10")

Data1.DatabaseName = MdbPath
Data1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
xStore1.ListField = "Desca"
xStore1.BoundColumn = "code"
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
Dim nT1, nT2, nT3, nT4, nT5, nT6 As Double
tempdb.Execute "DELETE * FROM TEMP"

cStr1 = " SELECT FILE1_10.GROUP, FILE1_10.DESCA, FILE1_10.ITEM, Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT FROM FILE1_10 LEFT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM " & _
        " WHERE STORE   = " & MyParn(xStore1.BoundText) & _
        " and Date < DateValue(" & MyParn(xDate1.Text) & ")" & _
        " GROUP BY FILE1_10.GROUP, FILE1_10.DESCA, FILE1_10.ITEM ORDER BY FILE1_10.ITEM"
Set itemTable = mydb.OpenRecordset(cStr1)

cStr1 = " SELECT  Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT , ITEM FROM FILE1_11 " & _
        " WHERE STORE   = " & MyParn(xStore1.BoundText) & _
        " and Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
        " GROUP BY FILE1_11.ITEM ORDER BY ITEM"
Set ITEM2table = mydb.OpenRecordset(cStr1)

cString = "SELECT FILE1_81.item, Sum(FILE1_81.QUANT) AS SumIn " & _
          "FROM FILE1_81 " & _
          " where Date Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE1_81.item "
Set InTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE1_82.item, Sum(FILE1_82.QUANT) AS SumOut " & _
          "FROM FILE1_82 " & _
          " where Date Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE1_82.item "
Set OutTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE1_60.STORE1, FILE1_60.item, Sum(FILE1_60.QUANT) AS SumTransFrom " & _
          "FROM FILE1_60 " & _
          " where Date Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE1 = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE1_60.STORE1, FILE1_60.item "
Set TransFromTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE1_60.STORE2, FILE1_60.item, Sum(FILE1_60.QUANT) AS SumTransTo " & _
          "FROM FILE1_60 " & _
          " where Date Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE2 = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE1_60.STORE2, FILE1_60.item "
Set TransToTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE6_20.STORE, FILE6_20.ITEM, Sum(FILE6_20.QUANT) AS SumSal " & _
          "FROM FILE6_20 " & _
          " where DATE Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE6_20.STORE, FILE6_20.ITEM "
Set SalTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE7_20.STORE, FILE7_20.ITEM, Sum(FILE7_20.QUANT) AS SumPurch " & _
          "FROM FILE7_20 " & _
          " where DATE Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE = " & MyParn(xStore1.BoundText) & _
          " GROUP BY FILE7_20.STORE, FILE7_20.ITEM "
Set PurchTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE0_10.item, FILE0_10.Store, Sum(FILE0_10.Differ) AS SumDiffer " & _
          "FROM FILE0_10 " & _
          " where date Between DateValue(" & MyParn(xDate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " and STORE = " & MyParn(xStore1.BoundText) & _
          " AND CLOSED = TRUE " & _
          " GROUP BY FILE0_10.item, FILE0_10.Store "
Set compTable = mydb.OpenRecordset(cString, dbOpenSnapshot)


With itemTable
If .RecordCount > 0 Then
    Do While Not .EOF
        TempTable.AddNew
        TempTable.val1 = TurnValue(.SumIn, Null, 0) - TurnValue(.SUMout, Null, 0)
        ITEM2table.FindFirst " ITEM = " & MyParn(.Item)
        If Not ITEM2table.NoMatch Then
            TempTable.val2 = TurnValue(ITEM2table.SumIn, Null, 0) - TurnValue(ITEM2table.SUMout, Null, 0)
        Else
            TempTable.val2 = 0
        End If
        InTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not InTable.NoMatch Then TempTable.VAL3 = InTable.SumIn
        
        OutTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not OutTable.NoMatch Then TempTable.VAL4 = OutTable.SUMout

        TransFromTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not TransFromTable.NoMatch Then TempTable.VAL5 = TransFromTable.SumTransFrom

        TransToTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not TransToTable.NoMatch Then TempTable.VAL6 = TransToTable.SumTransTO

        SalTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not SalTable.NoMatch Then TempTable.VAL7 = SalTable.SUMSAL

        PurchTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not PurchTable.NoMatch Then TempTable.VAL8 = PurchTable.SumPurch
        
        compTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not compTable.NoMatch Then TempTable.VAL9 = compTable.SumDiffer
        
        
        TempTable.str1 = itemTable.Item
        TempTable.str2 = itemTable.DESCA

        TempTable.str7 = " »Ì«‰ »≈Ã„«·Ï Õ—ﬂ… «·√’‰«› ·„Œ“‰ " & xStore1.Text
        TempTable.str8 = " „‰  «—ÌŒ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
        TempTable.str9 = firsttitle
        TempTable.str10 = Secondtitle
        TempTable.Update
        .MoveNext
    Loop
End If
End With

TempTable.MoveFirst
If TempTable.RecordCount > 0 Then
    Do While True
        If (TempTable.val1 + TempTable.val2 + TempTable.VAL3 + TempTable.VAL4 + TempTable.VAL5 + TempTable.VAL6 + TempTable.VAL7 + TempTable.VAL8) = 0 Then
            TempTable.Delete
        End If
        TempTable.MoveNext
        If TempTable.EOF Then Exit Do
    Loop
End If

Report1.ReportFileName = PublicPath & "\Reports\R_Item13.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
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
Function CountDBalance(pItem, pStore, pDate, lValue)
MyQuant = 0
movetable.FindFirst "ITEM = " & MyParn(pItem) & " AND STORE = " & MyParn(pStore)
If Not movetable.NoMatch Then
    If lValue Then
        Do While movetable.Item = pItem
            If movetable.Store = pStore And DateValue(movetable!Date) <= DateValue(pDate) And movetable!Type <> "" Then
                MyQuant = MyQuant + TurnValue(movetable.In, Null, 0) - TurnValue(movetable.OUT, Null, 0)
            End If
            movetable.MoveNext
            If movetable.EOF Then Exit Do
        Loop
    Else
        Do While movetable.Item = pItem
            If movetable.Store = pStore And DateValue(movetable.Date) < DateValue(pDate) And movetable!Type <> "" Then
                MyQuant = MyQuant + TurnValue(movetable.In, Null, 0) - TurnValue(movetable.OUT, Null, 0)
            End If
            movetable.MoveNext
            If movetable.EOF Then Exit Do
        Loop
    End If
End If
CountDBalance = MyQuant
End Function
