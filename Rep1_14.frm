VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Rep1_14 
   Caption         =   "ÊÞÇÑíÑ ÇáÃÕäÇÝ"
   ClientHeight    =   1830
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
   ScaleHeight     =   1830
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
      Left            =   2775
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   525
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
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   975
      Width           =   3765
      Begin VB.CommandButton CmdClear 
         Caption         =   "ÊÌÏíÏ"
         Height          =   390
         Left            =   1200
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "ÎÑæÌ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "ÚÑÖ"
         Height          =   390
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1140
      End
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "Rep1_14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim InTable As Recordset
Dim OutTable As Recordset
Dim PurchTable As Recordset
Dim SalTable As Recordset
Dim movetable As Recordset
Dim itemTable As Recordset
Dim nOption As Integer
Function MYVALID() As Boolean
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
Private Sub cmdundo_Click()
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
Set movetable = mydb.OpenRecordset("SELECT * FROM File1_11 ORDER BY ITEM , DATE ")
Set itemTable = mydb.OpenRecordset("SELECT * FROM FILE1_10 ORDER BY ITEM ")
Set TempTable = tempdb.OpenRecordset("Temp", dbOpenDynaset)
End Sub
Private Sub CmdApply2_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
TempTable.Requery
cString = "SELECT FILE1_81.item, Sum(FILE1_81.QUANT) AS SumIn " & _
          "FROM FILE1_81 " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " GROUP BY FILE1_81.item "
Set InTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE1_82.item, Sum(FILE1_82!Quant) AS SumOut " & _
          "FROM FILE1_82 " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " GROUP BY FILE1_82.item "
Set OutTable = mydb.OpenRecordset(cString, dbOpenSnapshot)


cString = "SELECT FILE6_20.ITEM, Sum(FILE6_20.QUANT) AS SumSal " & _
          "FROM FILE6_20 " & _
          " where DATE Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " GROUP BY FILE6_20.ITEM "
Set SalTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE7_20.ITEM, Sum(FILE7_20.QUANT) AS SumPurch " & _
          "FROM FILE7_20 " & _
          " where DATE Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " GROUP BY FILE7_20.ITEM "
Set PurchTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

With itemTable
If .RecordCount > 0 Then
    Do While Not .EOF
        TempTable.AddNew
        TempTable.VAL1 = ItemDBalance(.Item, xdate1.Text, False)
        TempTable.VAL2 = ItemDBalance(.Item, xDate2.Text, True)
        
        InTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not InTable.NoMatch Then TempTable.VAL3 = InTable.SUMIn
        
        OutTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not OutTable.NoMatch Then TempTable.VAL4 = InTable.SUMIn

        SalTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not SalTable.NoMatch Then TempTable.VAL7 = SalTable.SUMSAL

        PurchTable.FindFirst "ITEM =  " & MyParn(.Item)
        If Not PurchTable.NoMatch Then TempTable.VAL8 = PurchTable.SumPurch
        
        TempTable.str1 = itemTable.Item
        TempTable.str2 = itemTable.DESCA

        TempTable.str7 = " ÈíÇä ÈÅÌãÇáì ÍÑßÉ ÇáÃÕäÇÝ "
        TempTable.str8 = " ãä ÊÇÑíÎ " & xdate1.Text & " Åáì ÊÇÑíÎ " & xDate2.Text
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
        If (TempTable.VAL1 + TempTable.VAL2 + TempTable.VAL3 + TempTable.VAL4 + TempTable.VAL7 + TempTable.VAL8) = 0 Then
            TempTable.Delete
        End If
        TempTable.MoveNext
        If TempTable.EOF Then Exit Do
    Loop
End If

Report1.ReportFileName = PublicPath & "\Reports\R_ITEM14.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdApply_Click()
sDate = " ([Date] Between " & DateSql(xdate1.Text) & _
         " and " & DateSql(xDate2.Text) & ")"

cField1 = myiif("[Date] < " & DateSql(xdate1.Text) _
          , "val(Format([in]))- val(Format([out]))") & _
           " As FirstBalance"

cField2 = myiif(sDate & _
            " and Type = '6' " _
          , "val(Format([out]))") & _
           " As Sale "

cField3 = myiif(sDate & _
            " and Type = '3' " _
          , "val(Format([in]))") & _
           " As RetSale "

cField4 = myiif(sDate & _
            " and Type = '2' " _
          , "val(Format([in]))") & _
           " As Purchase "

cField5 = myiif(sDate & _
            " and Type = '7' " _
          , "val(Format([Out]))") & _
           " As RetPurchase "

cField6 = myiif(sDate & _
            " and Type = '4' " _
          , "val(Format([in]))") & _
           " As Input "

cField7 = myiif(sDate & _
            " and Type = '8' " _
          , "val(Format([Out]))") & _
           " As OutPut"

cField8 = myiif(sDate & _
            " and Type = '9' " _
          , "val(Format([Out]))") & _
           " As Damage"

cField9 = myiif("Date <= " & DateSql(xDate2.Text) & "" _
          , "val(Format([in]))- val(Format([out]))") & _
           " As LastBalance"

cString = "SELECT First(File1_10.item) as FirstOfItem,  First(File1_10.DescA) as firstOfDesca," & _
           cField1 & "," & cField2 & "," & cField3 & "," & _
           cField4 & "," & cField5 & "," & cField6 & "," & _
           cField7 & "," & cField8 & "," & cField9 & _
          " FROM FILE1_10 Right JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM " & _
          " GROUP BY FILE1_11.item "

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

tempdb.Execute "DELETE * FROM TEMP"
TempTable.Requery

With SourceTable
    Do While Not .EOF
        TempTable.AddNew
        TempTable.str1 = SourceTable!firstOfItem
        TempTable.str2 = SourceTable!FirstofDescA
        TempTable.str7 = " ÈíÇä ÈÅÌãÇáì ÍÑßÉ ÇáÃÕäÇÝ "
        TempTable.str8 = " ãä ÊÇÑíÎ " & xdate1.Text & " Åáì ÊÇÑíÎ " & xDate2.Text
        TempTable.str9 = firsttitle
        TempTable.str10 = Secondtitle
    
    
        TempTable.VAL1 = SourceTable!FirstBalance
        TempTable.VAL2 = SourceTable!Output
        TempTable.VAL3 = SourceTable!input
        TempTable.VAL4 = SourceTable!damage
        TempTable.VAL5 = SourceTable!Sale
        TempTable.VAL6 = SourceTable!RetSale
        TempTable.VAL7 = SourceTable!Purchase
        TempTable.VAL8 = SourceTable!RetPurchase
        TempTable.VAL9 = SourceTable!LastBalance
        TempTable.Update
        .MoveNext
    Loop
End With

Report1.ReportFileName = PublicPath & "\Reports\R_ITEM14.rpt"
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
Function ItemDBalance(pItem, pDate, lValue)
MyQuant = 0
movetable.Seek ">=", pItem
If myFound(movetable, Array(pItem)) Then
    If lValue Then
        Do While movetable.Item = pItem
            If DateValue(movetable.Date) <= DateValue(pDate) And movetable!Type <> "" Then
                MyQuant = MyQuant + TurnValue(movetable.In, Null, 0) - TurnValue(movetable.OUT, Null, 0)
            End If
            movetable.MoveNext
            If movetable.EOF Then Exit Do
        Loop
    Else
        Do While movetable.Item = pItem
            If DateValue(movetable.Date) < DateValue(pDate) And movetable!Type <> "" Then
                MyQuant = MyQuant + TurnValue(movetable.In, Null, 0) - TurnValue(movetable.OUT, Null, 0)
            End If
            movetable.MoveNext
            If movetable.EOF Then Exit Do
        Loop
    End If
End If
ItemDBalance = MyQuant
End Function

