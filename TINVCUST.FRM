VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form TInvCust 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ï „»Ì⁄«  ⁄„Ì· "
   ClientHeight    =   4545
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   9480
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
   ScaleHeight     =   4545
   ScaleWidth      =   9480
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
      Left            =   1725
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
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
      Left            =   150
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   1515
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   2160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSFlexGridLib.MSFlexGrid invGrid 
      Height          =   5565
      Left            =   300
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   825
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   9816
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
      Left            =   8700
      RightToLeft     =   -1  'True
      TabIndex        =   4
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
      Left            =   5775
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   225
      Width           =   2790
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   5865
      Left            =   150
      Top             =   675
      Width           =   9240
   End
End
Attribute VB_Name = "TInvCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim invtable As Recordset
Dim TempTable As Recordset
Dim cString As String
Sub Fillgrd()
I = 1
invGrid.Rows = 1
If formMode <> EmptyMode Then invGrid.Rows = 1
invtable.MoveFirst
Do
   invGrid.AddItem ""
   If lRepCust Then
        invGrid.TextArray(faIndex(I, 0, invGrid)) = TurnValue(invtable.DOC_NO, Null, "")
        invGrid.TextArray(faIndex(I, 1, invGrid)) = TurnValue(Format(invtable.Date, "DD-MM-YYYY"), Null, "")
        invGrid.TextArray(faIndex(I, 2, invGrid)) = TurnValue(Format(invtable.SUMTOTAL, "#0.00"), Null, "")
        invGrid.TextArray(faIndex(I, 3, invGrid)) = TurnValue(Format(invtable.SUMTOTAL - invtable.SUMCOST, "#0.00"), Null, "")
   Else
        invGrid.TextArray(faIndex(I, 0, invGrid)) = TurnValue(invtable.inv_No, Null, "")
        invGrid.TextArray(faIndex(I, 1, invGrid)) = TurnValue(Format(invtable.Date_INV, "DD-MM-YYYY"), Null, "")
        invGrid.TextArray(faIndex(I, 2, invGrid)) = TurnValue(invtable.SUMTOTAL, Null, "")
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
With invtable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    If lRepCust Then
        TempTable.str1 = .DOC_NO
        TempTable.Date1 = .Date
        TempTable.val1 = .SUMTOTAL
        TempTable.val2 = .SUMTOTAL - .SUMCOST
        TempTable.STR6 = "»Ì«‰ »≈Ã„«·Ï ﬁÌ„… „»Ì⁄«   «·⁄„Ì· " & xDesc.Caption & " „‰  «—ÌŒ " & TSalCust.xdate1.Text & " ≈·Ï  «—ÌŒ " & TSalCust.xDate2.Text
    Else
        TempTable.str1 = .inv_No
        TempTable.Date1 = .Date_INV
        TempTable.val1 = .SUMTOTAL
        TempTable.STR6 = "»Ì«‰ »≈Ã„«·Ï ﬁÌ„… „‘ —Ì«  „‰ «·„Ê—œ " & xDesc.Caption & " „‰  «—ÌŒ " & TSalCust.xdate1.Text & " ≈·Ï  «—ÌŒ " & TSalCust.xDate2.Text
    End If
'    TempTable.str19 = firsttitle
'    TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\TInvCust.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TInvCust = Nothing
End Sub
Private Sub Form_Load()
If lRepCust Then
    Me.Caption = "≈Ã„«·Ï „»Ì⁄«  ⁄„Ì· "
Else
    Me.Caption = " ≈Ã„«·Ï „‘ —Ì«  „Ê—œ"
End If
Dim cString As String
tempdb.Execute "Delete * From Temp"

'Me.Picture = LoadPicture(App.Path & "\mainback.jpg")
Set mydb = OpenDatabase(MdbPath)
xCode.Caption = TSalCust.xCode.Caption
xDesc.Caption = TSalCust.xDesc.Caption
If lRepCust Then
    invGrid.FormatString = "—ﬁ„ ›« Ê—…|" & " «—ÌŒ|" & "≈Ã„«·Ï ›« Ê—…|" & "≈Ã„«·Ï —»Õ"
    invGrid.Cols = 4
    invGrid.ColWidth(0) = 1500
    invGrid.ColWidth(1) = 2000
    invGrid.ColWidth(2) = 2000
    invGrid.ColWidth(3) = 2000
    
    cString = "Select File6_20.DOC_NO, " & _
              "File6_20.DATE," & _
              "Sum(File6_20.Total) as sumTotal ," & _
              "Sum(File6_20.COST * File6_20.QUANT ) as sumCOST " & _
              "From file6_20 " & _
              "Where file6_20.CODE = " & MyParn(TSalCust.xCode.Caption) & _
              " And FILE6_20.DATE Between DateValue(" & MyParn(TSalCust.xdate1.Text) & ") and DateValue(" & MyParn(TSalCust.xDate2.Text) & ")" & _
              " Group by DOC_NO,FILE6_20.DATE"
Else
    invGrid.FormatString = "—ﬁ„ ›« Ê—…|" & " «—ÌŒ|" & "≈Ã„«·Ï ›« Ê—…"
    invGrid.Cols = 3
    invGrid.ColWidth(0) = 1000
    invGrid.ColWidth(1) = 2000
    invGrid.ColWidth(2) = 2000
    
    cString = "Select File7_20.inv_no, " & _
              "File7_20.[DATE_inv]," & _
              "Sum(File7_20.Total) as sumTotal " & _
              "From file7_20 " & _
              "Where file7_20.CODE = " & MyParn(TSalCust.xCode.Caption) & _
              " And FILE7_20.DATE_inv Between DateValue(" & MyParn(TSalCust.xdate1.Text) & ") and DateValue(" & MyParn(TSalCust.xDate2.Text) & ")" & _
              " Group by inv_no,FILE7_20.DATE_inv"
End If
Set invtable = mydb.CreateSnapshot(cString)

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
    cString = "SELECT FILE6_20.*, FILE1_10.DESCA " & _
              "FROM FILE1_10 RIGHT JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM " & _
              "Where file6_20.DOC_NO = " & MyParn(GrdText(invGrid, 0)) & _
              " ORDER BY FILE6_20.ITEM"
Else
    cString = "SELECT FILE7_20.*, FILE1_10.DESCA " & _
              "FROM FILE1_10 RIGHT JOIN FILE7_20 ON FILE1_10.ITEM = FILE7_20.ITEM " & _
              "Where file7_20.inv_no = " & MyParn(GrdText(invGrid, 0)) & _
              " ORDER BY FILE7_20.ITEM"
End If
Set DataTable = mydb.CreateSnapshot(cString)

With DataTable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .Item
    TempTable.str2 = .DESCA
    TempTable.val1 = .Quant
    TempTable.val2 = .price
    TempTable.VAL3 = .TOTAL
    If lRepCust Then
    TempTable.VAL4 = .cost
    TempTable.VAL5 = (.TOTAL - (.cost * .Quant))
    End If
    TempTable.STR6 = "›« Ê—… —ﬁ„ " & GrdText(invGrid, 0) & "» «—ÌŒ  " & GrdText(invGrid, 1)
'   TempTable.str19 = firsttitle
'   TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\ListInv.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
