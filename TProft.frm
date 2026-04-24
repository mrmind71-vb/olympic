VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Tproft_shop 
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   345
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   225
      Width           =   3480
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   1365
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ãä äÇÑíÎ"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Åáì ÊÇÑíÎ"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4275
      Top             =   525
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
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   345
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1725
      Width           =   3480
      Begin VB.CommandButton Cmd_Exit 
         Caption         =   "ÎÜÜÜÑæÌ"
         Height          =   465
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton Cmd_Print 
         Caption         =   "ØÈÇÚÉ ÇáãæÞÝ"
         Height          =   465
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1800
      End
   End
End
Attribute VB_Name = "Tproft_shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_EXIT_Click()
    Unload Me
End Sub
Private Sub CMD_PRINT_Click()
Dim temptable As New ADODB.Recordset, aHeader(1)
Dim n1 As Double, n2 As Double, n3 As Double, n31 As Double, n4 As Double, n6 As Double
Dim n61 As Double, n7 As Double, n13 As Double, n14 As Double, n15 As Double, n16 As Double
Dim n17 As Double, n12 As Double

contemp.Execute "Delete * From Temp"
temptable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
contemp.BeginTrans

If Not (IsDate(xDate1.Text) And IsDate(xDate2.Text)) Then
    MsgBox "ÇáÊÇÑíÎ ÛíÑ ÕÍíÍ"
    Exit Sub
End If

cWhere = " date >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSql(xDate2.Text)

n2 = Val(GetDesca("select sum(Val(( FILE1_11.IN - FILE1_11.[OUT] ) & '')* Val(FILE1_11.PRICE & '')*(1-(Val(FILE1_11.DISCOUNT & '')/100))) as tall from file1_11 where (FILE1_11.TYPE = '2' OR FILE1_11.TYPE = '7' ) AND " & cWhere) & "")

n1 = Val(GetDesca("SELECT Sum(([IN]-[OUT])*[COST]*(1-(Val(DISC & '')/100)))  FROM FILE1_11 LEFT JOIN MyItemCost ON FILE1_11.ITEM = MyItemCost.ITEM  Where DATE < " & DateSq(xDate1.Text)) & "")


n3 = Val(GetDesca("SELECT Sum(Val((SALES.OUT-SALES.[IN]) & '')*Val(SALES.PRICE & '')*(1-(Val(SALES.DISCOUNT & '')/100))) AS TSALES FROM SALES INNER JOIN FILE3_10 ON SALES.code = FILE3_10.CODE where NOT FILE3_10.CUST AND " & cWhere) & "")
n31 = Val(GetDesca("SELECT Sum(DISCOUNT )  FROM FILE6_20H INNER JOIN FILE3_10 ON FILE6_20H.code = FILE3_10.CODE where NOT FILE3_10.CUST AND " & cWhere) & "")
n4 = Val(GetDesca("SELECT Sum(Val((SALES.OUT-SALES.[IN]) & '')*Val(MYITEMCOST.COST & '')*(1-(Val(MYITEMCOST.DISC & '')/100))) AS TSALES FROM (SALES INNER JOIN FILE3_10 ON SALES.code = FILE3_10.CODE) LEFT JOIN MyItemCost ON SALES.ITEM = MyItemCost.ITEM where NOT FILE3_10.CUST AND " & cWhere) & "")

n6 = Val(GetDesca("SELECT Sum(Val((SALES.OUT-SALES.[IN]) & '')*Val(SALES.PRICE & '')*(1-(Val(SALES.DISCOUNT & '')/100))) AS TSALES FROM SALES INNER JOIN FILE3_10 ON SALES.code = FILE3_10.CODE where FILE3_10.CUST AND " & cWhere) & "")
n61 = Val(GetDesca("SELECT Sum(DISCOUNT )  FROM FILE6_20H INNER JOIN FILE3_10 ON FILE6_20H.code = FILE3_10.CODE where FILE3_10.CUST AND " & cWhere) & "")
n7 = Val(GetDesca("SELECT Sum(Val((SALES.OUT-SALES.[IN]) & '')*Val(MYITEMCOST.COST & '')*(1-(Val(MYITEMCOST.DISC & '')/100))) AS TSALES FROM (SALES INNER JOIN FILE3_10 ON SALES.code = FILE3_10.CODE) LEFT JOIN MyItemCost ON SALES.ITEM = MyItemCost.ITEM WHERE FILE3_10.cust AND  " & cWhere) & "")

n13 = Val(GetDesca("SELECT Sum(file3_11.sal - file3_11.pay ) AS fbal FROM file3_11 WHERE FILE3_11.date < " & DateSq(xDate1.Text)) & "")
n14 = Val(GetDesca("SELECT Sum(file3_11.sal - file3_11.pay ) AS fbal FROM file3_11 WHERE FILE3_11.date <= " & DateSq(xDate2.Text)) & "")
n15 = Val(GetDesca("SELECT Sum(file3_11.pay ) AS tpay FROM file3_11 WHERE file3_11.[type] = '7' and  FILE3_11.date >= " & DateSq(xDate1.Text) & " and file3_11.date <= " & DateSq(xDate2.Text)) & "")
n16 = Val(GetDesca("SELECT Sum(file3_11.pay ) AS tpay FROM file3_11 WHERE (file3_11.[type] = '7' or file3_11.type = '10' ) and  FILE3_11.date >= " & DateSq(xDate1.Text) & " and file3_11.date <= " & DateSq(xDate2.Text)) & "")

n17 = Val(GetDesca("SELECT Sum(file3_11.pay ) AS tpay FROM file3_11 WHERE  file3_11.type = '10'  and  FILE3_11.date >= " & DateSq(xDate1.Text) & " and file3_11.date <= " & DateSq(xDate2.Text)) & "")

n12 = Val(GetDesca("SELECT Sum(([IN]-[OUT])*[COST]*(1-(Val(DISC & '')/100)))  FROM FILE1_11 LEFT JOIN MyItemCost ON FILE1_11.ITEM = MyItemCost.ITEM  Where DATE <= " & DateSq(xDate2.Text)) & "")
temptable.AddNew
temptable!Str11 = "ãæÞÝ ÅÌãÇáì ÇáãÍá"

temptable!str12 = " ãä " & DateFix(xDate1.Text) & " Åáì " & DateFix(xDate2.Text)


temptable!val1 = n1

temptable!val2 = Val(n2 & "")

temptable!val3 = Val(n3 & "") - Val(n31 & "")
temptable!val4 = Val(n4 & "")
temptable!val5 = Val(n3 & "") - Val(n31 & "") - Val(n4 & "")

temptable!Val6 = n6 - n61
temptable!Val7 = n7
temptable!Val8 = n6 - n61 - n7

temptable!val9 = Val(temptable!val3 & "") + Val(temptable!Val6 & "")
temptable!Val10 = Val(temptable!val4 & "") + Val(temptable!Val7 & "")
temptable!Val11 = Val(temptable!val5 & "") + Val(temptable!Val8 & "")

temptable!Val12 = n12

temptable!Val13 = Val(n13 & "")
temptable!VAL14 = Val(n14 & "")
temptable!val15 = Val(n15 & "")
temptable!val16 = Val(n16 & "")

temptable!val17 = Val(n17 & "")
temptable!val18 = Val(temptable!val9 & "") - Val(n17 & "")

temptable.Update
contemp.CommitTrans

mainfrm.Report1.ReportFileName = App.Path & "\Reports\T_PROFT.rpt"
mainfrm.Report1.DataFiles(0) = tempPath
mainfrm.Report1.Action = 1

temptable.Close
Set temptable = Nothing
End Sub
