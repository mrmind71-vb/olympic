VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpSales20 
   Caption         =   "«—»«Õ «’‰«› ÌÊ„Ï"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
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
   ScaleHeight     =   1875
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   420
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1395
      Width           =   1275
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1395
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   6180
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1680
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   210
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   1350
         TabIndex        =   7
         Top             =   975
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Œ“‰ :"
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1065
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì :"
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   675
         Width           =   465
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   765
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "rpSales20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If TypeOf ActiveControl Is DataCombo And KeyCode = 46 Then ActiveControl.BoundText = ""
End Sub
Private Sub cmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xdate1.Text) And Trim(xdate1.Text) <> "" Then
     MsgBox "«· «—ÌŒ «·«Ê· €Ì— ”·Ì„"
    Exit Function
End If
If Not IsDate(xDate2.Text) And Trim(xDate2.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(1), sDoc_no As String
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable


cString = "SELECT year(INV_TOTAL.DATE) as g_year,month(INV_TOTAL.date) as g_month , Sum(INV_TOTAL.TOTAL) as sumofTotal,SUM(INV_TOTAL.DISCOUNT) as SumofDiscount , Sum(INV_TOTAL.COST) AS sumofCost FROM INV_TOTAL"
If IsDate(xdate1.Text) Then
    cString = cString & turn(cString) & " date >= " & DateSq(xdate1.Text)
    aHeader(0) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & " date <= " & DateSq(xDate2.Text)
    aHeader(0) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If Trim(xStore.Text) <> "" Then
    cString = cString & turnFound(cString) & " store = " & MyParn(xStore.BoundText)
    aHeader(1) = "[" & "„⁄—÷ : " & xStore.Text & "]"
End If

cString = cString & " GROUP BY year(INV_TOTAL.date),month(INV_TOTAL.date)"

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    temptable.AddNew
'    If Check1.Value <> 0 Then
'        temptable!str1 = sourcetable!Desca
'        temptable!str2 = sourcetable!Store
'    End If
    temptable!val1 = sourcetable!g_Year
    temptable!val2 = sourcetable!g_Month
    temptable!Val3 = Val(sourcetable!sumofTotal & "") - Val(sourcetable!sumOfDiscount & "")
    temptable!Val7 = Val(sourcetable!sumofTotal & "") - Val(sourcetable!sumOfDiscount & "") - Val(sourcetable!SumOfCOST & "")
    temptable!str21 = TurnValue(retHeader(aHeader, 0, 2))
    temptable.Update
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\sales20.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub
Sub myProc()
ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Unload Search3
End Sub

Private Sub xGroup1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(2, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "Select code,Desca From file1_52"
    Generalarray(2) = "Order by Desca"
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    listarray(0, 0) = "«·«”„"
    listarray(0, 1) = "(%%DescA%%) "
    
    GrdArray(0, 0) = "«·ﬂÊœ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "«·≈”„"
    GrdArray(1, 1) = 3000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Search3.Show 1
End If
End Sub
Private Sub xGroup1_LostFocus()
xgroup1Desca.Caption = ""
If xGroup1.Text = "" Then Exit Sub
xgroup1Desca.Caption = GetDesca("select desca from FILE1_52 where code = " & MyParn(xGroup1.Text)) & ""
End Sub

Private Sub xitem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(3, 5)
    Dim GrdArray(3, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "Select File1_10.item,File1_10.Desca,File1_10.Length,File1_10.Width From file1_10"
    Generalarray(2) = "Order by file1_10.Desca"
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
    listarray(0, 1) = "(FILE1_10.ITEM LIKE '%cFilter%' or  %%FILE1_10.DescA%%) "
    
    listarray(1, 0) = "«·„Ã„Ê⁄…"
    listarray(1, 1) = "(FILE1_10.[GROUP] IN (SELECT CODE FROM FILE1_50 WHERE %%DESCA%%))"
    
    listarray(2, 0) = "„Ã„Ê⁄… «·»Ì⁄-«·‰Ê⁄"
    listarray(2, 1) = "(FILE1_10.GROUP1 IN (SELECT CODE FROM FILE1_52 WHERE %%DESCA%%) OR FILE1_10.TYPE IN (SELECT CODE FROM FILE1_53 WHERE %%DESCA%%) )"
    
    listarray(3, 0) = "«·ÿÊ·-«·⁄—÷"
    listarray(3, 1) = "(FILE1_10.LENGTH LIKE cFilter or  FILE1_10.WIDTH LIKE cFilter) "
    
    GrdArray(0, 0) = "ﬂÊœ «·’‰›"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "≈”„ «·’‰›"
    GrdArray(1, 1) = 3000
    
    GrdArray(2, 0) = "«·ÿÊ·"
    GrdArray(2, 1) = 1000
    
    GrdArray(3, 0) = "«·⁄—÷"
    GrdArray(3, 1) = 1000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "«” ⁄·«„ «·«’‰«›"
    Search3.Show 1
End If
End Sub

Private Sub xitem_LostFocus()
xitemDesca.Caption = ""
If XITEM.Text = "" Then Exit Sub
xitemDesca.Caption = GetDesca("select desca from FILE1_10 where item = " & MyParn(XITEM.Text)) & ""
End Sub

Private Sub xType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(2, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "Select code,Desca From file1_53"
    Generalarray(2) = "Order by Desca"
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    listarray(0, 0) = "«·«”„"
    listarray(0, 1) = "(%%DescA%%) "
    
    GrdArray(0, 0) = "«·ﬂÊœ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "«·‰Ê⁄"
    GrdArray(1, 1) = 3000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Search3.Show 1
End If
End Sub

Private Sub xType_LostFocus()
xtypeDesca.Caption = ""
If xType.Text = "" Then Exit Sub
xtypeDesca.Caption = GetDesca("select desca from FILE1_53 where code = " & MyParn(xType.Text)) & ""
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(2, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "Select code,Desca From file0_40"
    Generalarray(2) = "Order by Desca"
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    listarray(0, 0) = "«·«”„"
    listarray(0, 1) = "(%%DescA%%) "
    
    GrdArray(0, 0) = "«·ﬂÊœ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "«·≈”„"
    GrdArray(1, 1) = 3000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Search3.Show 1
End If
End Sub
Private Sub xCode_LostFocus()
xCodeDesca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCodeDesca.Caption = GetDesca("select desca from FILE0_40 where code = " & MyParn(xCode.Text)) & ""
End Sub

Private Sub xMAN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(2, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "Select code,Desca From file0_14"
    Generalarray(2) = "Order by Desca"
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    listarray(0, 0) = "«·«”„"
    listarray(0, 1) = "(%%DescA%%) "
    
    GrdArray(0, 0) = "«·ﬂÊœ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "«·≈”„"
    GrdArray(1, 1) = 3000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Search3.Show 1
End If
End Sub
Private Sub xman_LostFocus()
xManDescA.Caption = ""
If xMan.Text = "" Then Exit Sub
xManDescA.Caption = GetDesca("select desca from FILE0_14 where code = " & MyParn(xMan.Text)) & ""
End Sub
Private Sub Form_Load()

openCon con

data1.ConnectionString = strCon
data1.RecordSource = "FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
