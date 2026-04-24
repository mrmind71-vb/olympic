VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form TItemCust 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ï „»Ì⁄«  ⁄„Ì· "
   ClientHeight    =   5010
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   9390
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
   ScaleHeight     =   5010
   ScaleWidth      =   9390
   WindowState     =   2  'Maximized
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
      TabIndex        =   7
      Top             =   675
      Width           =   1740
   End
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
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   675
      Value           =   -1  'True
      Width           =   1740
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
      Left            =   2325
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   150
      Width           =   1515
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   1515
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   5550
      Top             =   75
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
      Height          =   345
      Left            =   4275
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   75
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSFlexGridLib.MSFlexGrid invGrid 
      Height          =   5265
      Left            =   225
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1275
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   9287
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
      Left            =   8625
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   225
      Width           =   690
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
      Left            =   5700
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   225
      Width           =   2790
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   5640
      Left            =   75
      Top             =   1050
      Width           =   9240
   End
   Begin VB.Label xClientName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1995
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1305
      Width           =   4425
   End
End
Attribute VB_Name = "TItemCust"
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
   invGrid.TextArray(faIndex(I, 1, invGrid)) = TurnValue(invtable.F_DescA, Null, "")
   invGrid.TextArray(faIndex(I, 2, invGrid)) = TurnValue(invtable.SUMQUANT, Null, "")
   invGrid.TextArray(faIndex(I, 3, invGrid)) = TurnValue(invtable.SUMSAL, Null, "")
   If lRepCust Then
   invGrid.TextArray(faIndex(I, 4, invGrid)) = TurnValue(Format(invtable.SUMSAL - invtable.SUMCOST, "#0.00"), Null, "")
   End If
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
Set invtable = mydb.CreateSnapshot(cString)
If invtable.RecordCount = 0 Then Exit Sub
invtable.MoveFirst
With invtable
Do
'  If .SUMSAL > 0 Then
    TempTable.AddNew
    TempTable.str1 = .Item
    TempTable.str4 = .F_DescA
    TempTable.val1 = .SUMQUANT
    TempTable.val2 = .SUMSAL
    If lRepCust Then
    TempTable.VAL3 = .SUMSAL - .SUMCOST
    TempTable.VAL4 = (.SUMSAL - .SUMCOST) * 100 / .SUMSAL
    TempTable.VAL5 = .F_COST
    TempTable.VAL6 = .SUMSAL / .SUMQUANT
    End If
    If lRepCust Then
    TempTable.str9 = ": ’«›Ï „»Ì⁄«  «·√’‰«› ··⁄„Ì·"
    Else
    TempTable.str9 = ": ’«›Ï „‘ —Ì«  √’‰«› „‰ «·„Ê—œ"
    End If
    TempTable.str10 = " „‰  «—ÌŒ " & TSalCust.xdate1.Text & " ≈·Ï  «—ÌŒ " & TSalCust.xdate1.Text
'   TargetTable.str19 = firsttitle
'   TargetTable.str20 = Secondtitle
    TempTable.Update
'    End If
    .MoveNext
Loop Until .EOF
End With
Report1.ReportFileName = PublicPath & "\Reports\TITMCUST.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TItemCust = Nothing
End Sub
Private Sub Form_Load()
If lRepCust Then
    Me.Caption = "≈Ã„«·Ï „»Ì⁄«  ⁄„Ì· "
Else
    Me.Caption = "≈Ã„«·Ï „‘ —Ì«  „Ê—œ "
End If
xCode.Caption = TSalCust.xCode.Caption
xDesc.Caption = TSalCust.xDesc.Caption
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
If lRepCust Then
    cString = "SELECT FILE1_10.ITEM, First(FILE1_10.DESCA) AS F_DESCA,  First(FILE1_10.COST) AS F_COST, " & _
              "Sum(FILE6_20.TOTAL) AS SumSAL , Sum(FILE6_20.QUANT) AS SumQuant , Sum(FILE6_20.COST * FILE6_20.QUANT) AS SumCOST " & _
              "FROM FILE6_20 LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM " & _
              " Where Date >= DateValue(" & MyParn(TSalCust.xdate1.Text) & ")" & _
              " and Date <= DateValue(" & MyParn(TSalCust.xDate2.Text) & ")" & _
              " And File6_20.Code = " & MyParn(TSalCust.xCode.Caption) & _
              " Group By File1_10.item"
    If nInd = 1 Then
        cString = cString & " ORDER BY FILE1_10.ITEM"
    Else
        cString = cString & " ORDER BY Sum(FILE6_20.TOTAL) DESC "
    End If
Else
    cString = "SELECT FILE1_10.ITEM, First(FILE1_10.DESCA) AS F_DESCA, " & _
              "Sum(FILE7_20.TOTAL) AS SumSAL , Sum(FILE7_20.QUANT) AS SumQuant " & _
              "FROM FILE7_20 LEFT JOIN FILE1_10 ON FILE7_20.ITEM = FILE1_10.ITEM " & _
              " Where Date_INV >= DateValue(" & MyParn(TSalCust.xdate1.Text) & ")" & _
              " and Date_INV <= DateValue(" & MyParn(TSalCust.xDate2.Text) & ")" & _
              " And File7_20.Code = " & MyParn(TSalCust.xCode.Caption) & _
              " Group By File1_10.item"
    If nInd = 1 Then
        cString = cString & " ORDER BY FILE1_10.ITEM"
    Else
        cString = cString & " ORDER BY Sum(FILE7_20.TOTAL) DESC "
    End If
End If
Set invtable = mydb.CreateSnapshot(cString)
If lRepCust Then
    invGrid.Cols = 5
    invGrid.FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹›|" & " ﬂ„Ì… „»Ì⁄« |" & " ﬁÌ„… «·„»Ì⁄« |" & " —»‹‹‹Õ"
    invGrid.ColWidth(0) = 1200
    invGrid.ColWidth(1) = 2500
    invGrid.ColWidth(2) = 1500
    invGrid.ColWidth(3) = 1500
    invGrid.ColWidth(4) = 1500

Else
    invGrid.Cols = 4
    invGrid.FormatString = "ﬂÊœ «·’‰›|" & "«·’‰‹‹‹‹›|" & "≈Ã„«·Ï ﬂ„Ì… „‘ —Ì« |" & "≈Ã„«·Ï ﬁÌ„… «·„‘ —Ì« "
    invGrid.ColWidth(0) = 1200
    invGrid.ColWidth(1) = 2500
    invGrid.ColWidth(2) = 2000
    invGrid.ColWidth(3) = 2000
End If
For I = 0 To invGrid.Cols - 1
    invGrid.ColAlignment(I) = 1
Next

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
If lRepCust Then
    cString = " SELECT FILE6_20.DOC_NO, FILE6_20.DATE AS DATEINV , FILE6_20.CODE, FILE6_20.ITEM, FILE6_20.QUANT, FILE6_20.PRICE, FILE6_20.TOTAL , FILE6_20.COST " & _
              " FROM FILE6_20 " & _
              " Where file6_20.code = " & MyParn(xCode.Caption) & _
              " and file6_20.item = " & MyParn(GrdText(invGrid, 0)) & _
              " And FILE6_20.DATE Between " & DateSql(TSalCust.xdate1.Text) & " and " & DateSql(TSalCust.xDate2.Text) & _
              " ORDER BY FILE6_20.date"
Else
    cString = "SELECT FILE7_20.inv_no AS DOC_NO , FILE7_20.DATE_inv AS DATEINV , FILE7_20.CODE, FILE7_20.ITEM, FILE7_20.QUANT, FILE7_20.PRICE, FILE7_20.TOTAL " & _
              "FROM FILE7_20 " & _
              "Where file7_20.code = " & MyParn(xCode.Caption) & _
              "and file7_20.item = " & MyParn(GrdText(invGrid, 0)) & _
              " And FILE7_20.DATE_inv Between DateValue(" & MyParn(TSalCust.xdate1.Text) & ") and DateValue(" & MyParn(TSalCust.xDate2.Text) & ")" & _
              " ORDER BY FILE7_20.date_inv"
End If
Set DataTable = mydb.CreateSnapshot(cString)
With DataTable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    If lRepCust Then
        TempTable.str1 = " „ «»⁄… „»Ì⁄«  «·’‰› " & MyParn(GrdText(invGrid, 1))
        TempTable.str2 = " ··⁄„Ì· " & xDesc.Caption
        TempTable.VAL4 = .cost
        TempTable.VAL5 = (.TOTAL - (.cost * .Quant))
    Else
        TempTable.str1 = " „ «»⁄… „‘ —Ì«  «·’‰› " & MyParn(GrdText(invGrid, 1))
        TempTable.str2 = " ··„Ê—œ  " & xDesc.Caption
    End If
    TempTable.str4 " Œ·«· «·› —… „‰ " & TSalCust.xdate1.Text & " ≈·Ï " & TSalCust.xDate2.Text
    TempTable.str3 = .DOC_NO
    TempTable.val2 = .Quant
    TempTable.VAL3 = .price
    TempTable.val1 = .TOTAL
    
    TempTable.date3 = .DateINV
'   TempTable.str19 = firsttitle
'   TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\TItemC.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Private Sub xInd_Click(Index As Integer)
    If xInd(Index).Value = True Then nInd = Index
    If lRepCust Then
        cString = "SELECT FILE1_10.ITEM, First(FILE1_10.DESCA) AS F_DESCA,  First(FILE1_10.COST) AS F_COST, " & _
                  "Sum(FILE6_20.TOTAL) AS SumSAL , Sum(FILE6_20.QUANT) AS SumQuant , Sum(FILE6_20.COST * FILE6_20.QUANT) AS SumCOST " & _
                  "FROM FILE6_20 LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM " & _
                  " Where Date >= DateValue(" & MyParn(TSalCust.xdate1.Text) & ")" & _
                  " and Date <= DateValue(" & MyParn(TSalCust.xDate2.Text) & ")" & _
                  " And File6_20.Code = " & MyParn(TSalCust.xCode.Caption) & _
                  " Group By File1_10.item"
        If nInd = 1 Then
            cString = cString & " ORDER BY FILE1_10.ITEM"
        Else
            cString = cString & " ORDER BY Sum(FILE6_20.TOTAL) DESC "
        End If
    Else
        cString = "SELECT FILE1_10.ITEM, First(FILE1_10.DESCA) AS F_DESCA, " & _
                  "Sum(FILE7_20.TOTAL) AS SumSAL , Sum(FILE7_20.QUANT) AS SumQuant " & _
                  "FROM FILE7_20 LEFT JOIN FILE1_10 ON FILE7_20.ITEM = FILE1_10.ITEM " & _
                  " Where Date_INV >= DateValue(" & MyParn(TSalCust.xdate1.Text) & ")" & _
                  " and Date_INV <= DateValue(" & MyParn(TSalCust.xDate2.Text) & ")" & _
                  " And File7_20.Code = " & MyParn(TSalCust.xCode.Caption) & _
                  " Group By File1_10.item"
        If nInd = 1 Then
            cString = cString & " ORDER BY FILE1_10.ITEM"
        Else
            cString = cString & " ORDER BY Sum(FILE7_20.TOTAL) DESC "
        End If
    End If
    Set invtable = mydb.CreateSnapshot(cString)
    If lRepCust Then
        invGrid.Cols = 5
        invGrid.FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹›|" & " ﬂ„Ì… „»Ì⁄« |" & " ﬁÌ„… «·„»Ì⁄« |" & " —»‹‹‹Õ"
        invGrid.ColWidth(0) = 1200
        invGrid.ColWidth(1) = 2500
        invGrid.ColWidth(2) = 1500
        invGrid.ColWidth(3) = 1500
        invGrid.ColWidth(4) = 1500
    
    Else
        invGrid.Cols = 4
        invGrid.FormatString = "ﬂÊœ «·’‰›|" & "«·’‰‹‹‹‹›|" & "≈Ã„«·Ï ﬂ„Ì… „‘ —Ì« |" & "≈Ã„«·Ï ﬁÌ„… «·„‘ —Ì« "
        invGrid.ColWidth(0) = 1200
        invGrid.ColWidth(1) = 2500
        invGrid.ColWidth(2) = 2000
        invGrid.ColWidth(3) = 2000
    End If
    For I = 0 To invGrid.Cols - 1
        invGrid.ColAlignment(I) = 1
    Next
    
    If invtable.RecordCount > 0 Then
        Fillgrd
    Else
        invGrid.Rows = 1
    End If
End Sub
