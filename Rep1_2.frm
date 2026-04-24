VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Rep1_2 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   3720
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
   ScaleHeight     =   3720
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox x1 
      Alignment       =   1  'Right Justify
      Caption         =   "”⁄— Ã„·…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4785
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2640
      Width           =   1140
   End
   Begin VB.CheckBox x4 
      Alignment       =   1  'Right Justify
      Caption         =   "”⁄— „Õ·« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2640
      Width           =   1140
   End
   Begin VB.CheckBox x3 
      Alignment       =   1  'Right Justify
      Caption         =   "”⁄— „” Â·ﬂ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1665
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox x2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ã„·… „ Ê”ÿ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox xBal 
      Alignment       =   1  'Right Justify
      Caption         =   "√’‰«› ·Â« —’Ìœ ›ﬁÿ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2280
      Width           =   1845
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1800
      Width           =   1515
   End
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   " ›—Ì€"
      Height          =   390
      Left            =   2175
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3180
      Width           =   1515
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1425
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "Rep1_2.frx":0000
      Height          =   315
      Left            =   825
      TabIndex        =   0
      Top             =   225
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
   End
   Begin MSDBCtls.DBCombo xGroup1 
      Bindings        =   "Rep1_2.frx":0014
      DataSource      =   "Data1"
      Height          =   315
      Left            =   825
      TabIndex        =   1
      Top             =   618
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "«” Ã«»…"
      Height          =   390
      Left            =   3765
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3195
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3180
      Width           =   1515
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   120
      Top             =   600
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
   Begin MSDBCtls.DBCombo xGroup2 
      Bindings        =   "Rep1_2.frx":0028
      Height          =   315
      Left            =   825
      TabIndex        =   2
      Top             =   1011
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDBCtls.DBCombo xStore 
      Bindings        =   "Rep1_2.frx":003C
      DataSource      =   "Data3"
      Height          =   315
      Left            =   825
      TabIndex        =   9
      Top             =   1404
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4425
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1800
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„Œ“‰"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4425
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1425
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰ „Ã„Ê⁄… "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4425
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   675
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï „Ã„Ê⁄…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4425
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1050
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄… —∆Ì”Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4425
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   300
      Width           =   1095
   End
End
Attribute VB_Name = "Rep1_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
RepItem1
End Sub
Private Sub RepItem1()
Dim TempTable As Recordset
Dim cHead As String
tempdb.Execute "DELETE * FROM TEMP"
Set TempTable = tempdb.OpenRecordset("Temp")
cString = "SELECT FILE1_10.ITEM, FILE1_50.DESCA AS grdesc, FILE1_70.DESCA AS MGrdesc, FILE1_50.M_GROUP, Sum(FILE1_11.OUT) AS TOUT, Sum(FILE1_11.[IN]) AS TIN, FILE1_10.DESCA, FILE1_10.COST1, FILE1_10.COST2 , FILE1_10.COST4 , FILE1_10.PRICE , FILE1_10.PACK, FILE1_10.FACTCODE , FILE1_10.GROUP " & _
            " FROM FILE1_11 RIGHT JOIN ((FILE1_50 INNER JOIN FILE1_10 ON FILE1_50.CODE = FILE1_10.GROUP) INNER JOIN FILE1_70 ON FILE1_50.M_GROUP =  FILE1_70.CODE) ON FILE1_11.ITEM = FILE1_10.ITEM " & _
            " Where File1_70.Flag = 2 "
If xGroup.Text <> "" And (xGroup1.Text = "" Or xGroup2.Text = "") Then cString = cString & " and File1_50.M_GROUP = " & MyParn(xGroup.BoundText)
If xGroup1.BoundText <> "" Then cString = cString & " and File1_10.Group >= " & MyParn(xGroup1.BoundText)
If xGroup2.BoundText <> "" Then cString = cString & " and File1_10.Group <= " & MyParn(xGroup2.BoundText)
If xStore.BoundText <> "" Then cString = cString & " and File1_11.STORE = " & MyParn(xStore.BoundText)
If IsDate(xDate.Text) Then cString = cString & " AND FILE1_11.DATE <= " & DateSql(xDate.Text)
cHead = "√—’œ… «·‘—ﬂ… "
If IsDate(xDate.Text) Then cHead = cHead & " · «—ÌŒ  " & Format(xDate.Text, "DD-MM-YYYY")
If xStore.BoundText <> "" Then cHead = cHead & " ·„Œ“‰ " & xStore.Text

cString = cString & " GROUP BY FILE1_10.ITEM, FILE1_50.DESCA, FILE1_70.DESCA, FILE1_50.M_GROUP, FILE1_10.GROUP, FILE1_10.ITEM, FILE1_70.FLAG, FILE1_10.DESCA, FILE1_10.COST1, FILE1_10.PACK, FILE1_10.FACTCODE , FILE1_10.GROUP , FILE1_10.COST2 , FILE1_10.COST4 , FILE1_10.PRICE  "

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
i = 1
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
            TempTable.AddNew
            TempTable.str6 = .MGrdesc
            TempTable.STR5 = .m_Group
            TempTable.str1 = .Item
            TempTable.str2 = .DESCA
            TempTable.str3 = !Group
            TempTable.str4 = .GRDESC
            nBal = TurnValue(.TIN, Null, 0) - TurnValue(.TOUT, Null, 0)
            TempTable.VAL6 = nBal
            
            If x1.Value Then
                TempTable.VAL2 = .COST1
                TempTable.VAL3 = .COST1 * nBal
                TempTable.Str11 = "”⁄— Ã„·…"
            End If
            If x2.Value Then
                TempTable.VAL2 = .COST2
                TempTable.VAL3 = .COST2 * nBal
                TempTable.Str11 = "Ã„·… „ Ê”ÿ"
            End If
            If x3.Value Then
                TempTable.VAL2 = .price
                TempTable.VAL3 = .price * nBal
                TempTable.Str11 = "”⁄— „” Â·ﬂ"
            End If
            If x4.Value Then
                TempTable.VAL2 = .COST4
                TempTable.VAL3 = .COST4 * nBal
                TempTable.Str11 = "”⁄— „Õ·« "
            End If
            TempTable.STR15 = .FACTCODE
            TempTable.str8 = firsttitle
            TempTable.str9 = cHead
            If nBal <> 0 Then TempTable.Update
       .MoveNext
    Loop
End If
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\Item02.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Private Sub RepItem2()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Set TempTable = tempdb.OpenRecordset("Temp")
Dim LastCost As Recordset
Set LastCost = mydb.OpenRecordset(cStr1)

cString = "SELECT FILE1_10.ITEM, First(FILE1_10.PRICE) AS ItemPRICE, " & _
          " First(FILE1_10.COST) AS ItemCOST, First(FILE1_10.DESCA) AS ItemDESCA,  " & _
          "Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT, " & _
          " First(FILE1_50.M_GROUP) AS MGrCode, First(FILE1_50.DESCA) AS GrDESC,  " & _
          " First(FILE1_10.GROUP) AS GrCode, First(FILE1_70.DESCA) AS MGrDesc" & _
          " FROM ((FILE1_10 RIGHT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT " & _
          " JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_70 ON " & _
          " FILE1_50.M_GROUP = FILE1_70.CODE"

sQuery = " Where File1_70.Flag = 2"

If xGroup.Text <> "" And (xGroup1.Text = "" Or xGroup2.Text = "") Then
    sQuery = myQuery(sQuery) & "File1_50.M_GROUP = " & MyParn(xGroup.BoundText)
End If

If xGroup1.BoundText <> "" Then
    sQuery = myQuery(sQuery) & "File1_10.Group >= " & MyParn(xGroup1.BoundText)
End If

If xGroup2.BoundText <> "" Then
    sQuery = myQuery(sQuery) & "File1_10.Group <= " & MyParn(xGroup2.BoundText)
End If

If xStore.BoundText <> "" Then
    sQuery = myQuery(sQuery) & "File1_11.store = " & MyParn(xStore.BoundText)
End If

If IsDate(xDate.Text) Then
    sQuery = myQuery(sQuery) & " Date <= DateValue(" & MyParn(xDate.Text) & ")"
End If

cString = cString & sQuery & " GROUP BY File1_10.ITEM "
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
i = 1
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
        If Format((TurnValue(SourceTable.SUMIN, Null, 0) - TurnValue(SourceTable.SUMOUT, Null, 0)), "##0.00") <> "0.00" Then
            TempTable.AddNew
            TempTable.str6 = .MGrdesc
            TempTable.STR5 = .MGrCode
            TempTable.str1 = .Item
            TempTable.str2 = .ItemDescA
            TempTable.str3 = .GrCode
            TempTable.str4 = .GRDESC
            
            TempTable.VAL1 = .ITEMcost
            TempTable.VAL2 = .ITEMcost * TempTable.VAL3
            TempTable.str7 = " ﬁ—Ì— «—’œ… «·‘—ﬂ…"
            If xStore.BoundText <> "" Then
                TempTable.str7 = TempTable.str7 & " ·„Œ“‰ " & xStore.Text
            End If
            If IsDate(xDate.Text) Then
                TempTable.str10 = "«·√—’œ… Õ Ï  «—ÌŒ " & Format(xDate.Text)
            End If
            TempTable.str8 = firsttitle
            TempTable.str9 = Secondtitle
            TempTable.Update
        End If
       .MoveNext
    Loop
End If
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\repItm_2.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Private Sub CmdClear_Click()
xGroup.BoundText = ""
xGroup1.BoundText = ""
xGroup2.BoundText = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Command1_Click()

End Sub
Private Sub Form_Load()
Data1.DatabaseName = MdbPath
Data1.RecordSource = "Select Code,DescA From File1_70 Where Flag = 2"
Data2.DatabaseName = MdbPath
Data2.RecordSource = "Select Code,DescA From File1_50 "
xGroup.BoundColumn = "Code"
xGroup.ListField = "DescA"
xGroup1.BoundColumn = "Code"
xGroup1.ListField = "Desca"
xGroup2.BoundColumn = "Code"
xGroup2.ListField = "Desca"

Data3.DatabaseName = MdbPath
Data3.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
xStore.ListField = "Desca"
xStore.BoundColumn = "code"
If publicFlag = 1 Then
    Label2(2).Visible = False
    xDate.Visible = False
End If
x1.Visible = bopt2
x2.Visible = bopt2
x4.Visible = bopt2
End Sub

Private Sub xGroup_Change()
Data2.RecordSource = "Select * From File1_50" & IIf(xGroup.BoundText <> "", " where M_GROUP = " & MyParn(xGroup.BoundText), "")
Data2.Refresh
End Sub
Private Sub REPITEM6()
Dim SourceTable As Recordset
Dim TargetTable As Recordset
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
CFIELD1 = myiif("cType = '6' or ctype = '0' ", "[OUT]") & " as Sales,"
CFIELD2 = myiif("cType = '3'", "[IN]") & " as RetSales, "
cField3 = myiif("cType = '6'or ctype = '0' ", "[TOTAL]") & " as SalesValue, "
cField4 = myiif("cType = '3'", "[TOTAL]") & " as RetSalesValue "

cString = "Select File1_11.Item as MidOfItem," & _
          "First(File1_10.DescA) as FirstOfDescA,File1_10.Pack," & _
          CFIELD1 & CFIELD2 & cField3 & cField4 & _
          " From File1_11 Inner Join file1_10 on file1_11.Item = file1_10.Item " & _
          " Where cDate Between DateValue(" & MyParn(Date1.Text) & ")" & _
          " and DateValue(" & MyParn(date2.Text) & ")" & _
          " Group By File1_11.Item,File1_10.Pack"

Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "«·   ÊÃœ »Ì‰«  ›Ï «· ﬁ—Ì— ø"
    Exit Sub
End If
With SourceTable
Do
    If .Sales + .RetSales <> 0 Then
    TargetTable.AddNew
    TargetTable.str1 = SourceTable.MidofItem
    TargetTable.str2 = SourceTable.FirstofDescA
    TargetTable.str3 = " ≈Ã„«·Ï „»Ì⁄«  «·√’‰«› ﬂ„Ì… - ﬁÌ„…"
    TargetTable.VAL1 = .Sales
    TargetTable.VAL2 = .RetSales
    TargetTable.VAL3 = .Sales - .RetSales
    TargetTable.VAL7 = SourceTable.salesvalue
    TargetTable.VAL8 = SourceTable.retsalesvalue
    TargetTable.VAL9 = SourceTable.salesvalue - SourceTable.retsalesvalue
    
    TargetTable.Date1 = Date1.Text
    TargetTable.date2 = date2.Text
    TargetTable.str9 = Mid(firsttitle, 1, 50)
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    End If
    SourceTable.MoveNext

Loop Until SourceTable.EOF
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\RepItm7.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub

