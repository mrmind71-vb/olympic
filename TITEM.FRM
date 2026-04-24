VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form TItem 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ï „»Ì⁄«  ⁄„Ì· "
   ClientHeight    =   6435
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   9750
   Begin VB.OptionButton xInd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "«· — Ì» ⁄·Ï «·ﬂÊœ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2175
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   675
      Value           =   -1  'True
      Width           =   1740
   End
   Begin VB.OptionButton xInd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "«· — Ì» ⁄·Ï «·ﬁÌ„…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   675
      Width           =   1740
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   150
      Width           =   1515
   End
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1650
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   150
      Width           =   1515
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4050
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4650
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   75
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSFlexGridLib.MSFlexGrid invGrid 
      Height          =   5040
      Left            =   225
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   8890
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   32896
      ForeColorFixed  =   16777215
      BackColorSel    =   65535
      BackColorBkg    =   16777215
      Enabled         =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   5265
      Left            =   75
      Top             =   1050
      Width           =   9540
   End
   Begin VB.Label xDesc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Mudir MT"
         Size            =   10.5
         Charset         =   178
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6075
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   225
      Width           =   2790
   End
   Begin VB.Label xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Mudir MT"
         Size            =   10.5
         Charset         =   178
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   225
      Width           =   690
   End
End
Attribute VB_Name = "TItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim invtable As Recordset
Dim TempTable As Recordset
Dim cString As String
Dim nInd As Byte
Sub Fillgrd()
I = 1
invGrid.Rows = 1
If formMode <> EmptyMode Then invGrid.Rows = 1
invtable.MoveFirst
Do
   invGrid.AddItem ""
   invGrid.TextArray(faIndex(I, 0, invGrid)) = TurnValue(invtable.Item, Null, "")
   invGrid.TextArray(faIndex(I, 1, invGrid)) = TurnValue(invtable.DescaITEM, Null, "")
   invGrid.TextArray(faIndex(I, 2, invGrid)) = TurnValue(invtable.SUMQUANT, Null, "")
   invGrid.TextArray(faIndex(I, 3, invGrid)) = TurnValue(Format(invtable.SUMTOTAL, "#0.00"), Null, "")
   invGrid.TextArray(faIndex(I, 4, invGrid)) = TurnValue(Format(invtable.SUMTOTAL - invtable.SUMCOST, "#0.00"), Null, "")
   invtable.MoveNext
   I = I + 1
Loop Until invtable.EOF
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Cmd_Print_Click()
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")
With invtable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .Item
    TempTable.str2 = .DescaITEM
    TempTable.val1 = .SUMQUANT
    TempTable.val2 = .SUMTOTAL
    TempTable.VAL3 = .SUMTOTAL - .SUMCOST
    TempTable.STR6 = "»Ì«‰ »≈Ã„«·Ï ﬁÌ„… „»Ì⁄«  √’‰«› «·„Ã„Ê⁄…«·›—⁄Ì… ·„" & xDesc.Caption & " „‰  «—ÌŒ " & TSalItem.xdate1.Text & " ≈·Ï  «—ÌŒ " & TSalItem.xDate2.Text
'    TempTable.str19 = firsttitle
'    TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\Tsalitem.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
    Unload Me
    Set TItem = Nothing
End Sub
Private Sub Form_Load()
Dim cString As String
tempdb.Execute "Delete * From Temp"
'Me.Picture = LoadPicture(App.Path & "\mainback.jpg")
Set mydb = OpenDatabase(MdbPath)
xCode.Caption = TGrItem.xCode.Caption
xDesc.Caption = TGrItem.xDesc.Caption
invGrid.FormatString = "«·ﬂÊœ|" & "«·»Ì«‰|" & "≈Ã„«·Ï ﬂ„Ì…|" & "≈Ã„«·Ï ﬁÌ„…|" & "≈Ã„«·Ï —»Õ"
invGrid.Cols = 5
invGrid.ColWidth(0) = 1000
invGrid.ColWidth(1) = 3000
invGrid.ColWidth(2) = 1500
invGrid.ColWidth(3) = 1500
invGrid.ColWidth(4) = 1500

cString = "SELECT FILE1_10.ITEM AS Item , First (FILE1_10.DESCA) AS DESCAItem, Sum(FILE6_20.TOTAL) AS SumTOTAL, Sum(FILE6_20.QUANT) AS SumQUANT , Sum(FILE6_20.QUANT * FILE6_20.COST) AS SumCOST " & _
          " FROM (FILE1_10 RIGHT JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE " & _
          " where file1_50.CODE = " & MyParn(TGrItem.xCode.Caption) & _
          " And FILE6_20.DATE Between DateValue(" & MyParn(TSalItem.xdate1.Text) & ") and DateValue(" & MyParn(TSalItem.xDate2.Text) & ")" & _
          " GROUP BY FILE1_10.ITEM"
If nInd = 1 Then
    cString = cString & " ORDER BY ITEM "
Else
    cString = cString & " ORDER BY Sum(FILE6_20.TOTAL) DESC "
End If
Set invtable = mydb.CreateSnapshot(cString)
If invtable.RecordCount > 0 Then
    Fillgrd
Else
    invGrid.Rows = 1
End If
End Sub
Private Sub InvGrid_Click()
Dim DataTable As Recordset
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")

cString = "SELECT FILE6_20.DOC_NO, FILE6_20.DATE, FILE6_20.QUANT, FILE6_20.PRICE, FILE6_20.TOTAL , FILE3_10.DESCA , FILE6_20.COST " & _
          "FROM FILE3_10 RIGHT JOIN FILE6_20 ON FILE3_10.CODE = FILE6_20.CODE " & _
          "Where file6_20.item = " & MyParn(GrdText(invGrid, 0)) & _
          " And FILE6_20.DATE Between DateValue(" & MyParn(TSalItem.xdate1.Text) & ") and DateValue(" & MyParn(TSalItem.xDate2.Text) & ")" & _
          " ORDER BY FILE6_20.date"
Set DataTable = mydb.CreateSnapshot(cString)

With DataTable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .DOC_NO
    TempTable.str2 = .DESCA
    TempTable.val1 = .Quant
    TempTable.val2 = .price
    TempTable.VAL3 = .TOTAL
    TempTable.VAL4 = .cost
    TempTable.VAL5 = .TOTAL - (.cost * .Quant)
    TempTable.Date1 = .[Date]
    TempTable.STR6 = "„»Ì⁄«  ’‰›  " & GrdText(invGrid, 0) & "  " & GrdText(invGrid, 1)
    TempTable.STR7 = " „‰  «—ÌŒ " & TSalItem.xdate1.Text & " ≈·Ï  «—ÌŒ " & TSalItem.xDate2.Text
'   TempTable.str19 = firsttitle
'   TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\itemsal.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Private Sub xInd_Click(Index As Integer)
    If xInd(Index).Value = True Then nInd = Index
    cString = "SELECT FILE1_10.ITEM AS Item , First (FILE1_10.DESCA) AS DESCAItem, Sum(FILE6_20.TOTAL) AS SumTOTAL, Sum(FILE6_20.QUANT) AS SumQUANT , Sum(FILE6_20.QUANT * FILE6_20.COST) AS SumCOST " & _
              " FROM (FILE1_10 RIGHT JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE " & _
              " where file1_50.CODE = " & MyParn(TGrItem.xCode.Caption) & _
              " And FILE6_20.DATE Between DateValue(" & MyParn(TSalItem.xdate1.Text) & ") and DateValue(" & MyParn(TSalItem.xDate2.Text) & ")" & _
              " GROUP BY FILE1_10.ITEM"
    If nInd = 1 Then
        cString = cString & " ORDER BY FILE1_10.ITEM "
    Else
        cString = cString & " ORDER BY Sum(FILE6_20.TOTAL) DESC "
    End If
    Set invtable = mydb.CreateSnapshot(cString)
    If invtable.RecordCount > 0 Then
        Fillgrd
    Else
        invGrid.Rows = 1
    End If
End Sub
