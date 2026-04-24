VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form REP1_24 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   3555
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
   ScaleHeight     =   3555
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1845
      Width           =   1665
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2250
      Width           =   1665
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
      Top             =   1800
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   " ›—Ì€"
      Height          =   390
      Left            =   2130
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3090
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "REP1_24.frx":0000
      Height          =   315
      Left            =   795
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
      Height          =   465
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   1440
   End
   Begin MSDBCtls.DBCombo xGroup1 
      Bindings        =   "REP1_24.frx":0014
      DataSource      =   "Data1"
      Height          =   315
      Left            =   795
      TabIndex        =   1
      Top             =   630
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
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3090
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   555
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3090
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
      Bindings        =   "REP1_24.frx":0028
      Height          =   315
      Left            =   795
      TabIndex        =   2
      Top             =   1035
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
      Bindings        =   "REP1_24.frx":003C
      DataSource      =   "Data3"
      Height          =   315
      Left            =   825
      TabIndex        =   9
      Top             =   1440
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1830
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·Ï  «—ÌŒ :"
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
      Index           =   2
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2325
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„Œ“‰"
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
      Index           =   1
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "„‰ „Ã„Ê⁄… "
      Height          =   315
      Index           =   0
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "≈·Ï „Ã„Ê⁄…"
      Height          =   255
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1050
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "„Ã„Ê⁄… —∆Ì”Ì…"
      Height          =   195
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   300
      Width           =   1110
   End
End
Attribute VB_Name = "REP1_24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
    RepItem2
End Sub
Private Sub RepItem2()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Set TempTable = tempdb.OpenRecordset("Temp")
cString = "SELECT Min(FILE6_20.PRICE) AS MinPRICE, Max(FILE6_20.PRICE) AS MaxPRICE, Sum(FILE6_20.TOTAL) AS  " & _
          " SumTOTAL, Sum(FILE6_20.QUANT) AS SumQUANT, FILE1_10.ITEM, FILE1_10.DESCA, First(FILE1_50.M_GROUP)   " & _
          " AS MGrCode, First(FILE1_50.DESCA) AS GrDESC, First(FILE1_10.GROUP) AS GrCode, First(FILE1_70.DESCA) AS  MGrDesc " & _
          " FROM FILE6_20 LEFT JOIN ((FILE1_10 LEFT JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN  " & _
          " FILE1_70 ON FILE1_50.M_GROUP = FILE1_70.CODE) ON FILE6_20.ITEM = FILE1_10.ITEM " & _
          " Where File1_70.Flag = 2 "

If xStore.BoundText <> "" Then
    sQuery = myQuery(sQuery) & "File6_20.store = " & MyParn(xStore.BoundText)
End If

If IsDate(Date1.Text) Then
    sQuery = myQuery(sQuery) & " file6_20.Date >= DateValue(" & MyParn(Date1.Text)
End If

If IsDate(date2.Text) Then
    sQuery = myQuery(sQuery) & " file6_20.Date <= DateValue(" & MyParn(date2.Text)
End If

cString = cString & sQuery & " GROUP BY FILE1_10.ITEM, FILE1_10.DESCA "
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
i = 1
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
        TempTable.AddNew
        TempTable.str6 = .MGrdesc
        TempTable.STR5 = .MGrCode
        TempTable.str1 = .Item
        TempTable.str2 = .DESCA
        TempTable.str3 = .GrCode
        TempTable.str4 = .GRDESC
        
        TempTable.VAL1 = .SUMQUANT
        TempTable.VAL2 = .SUMTOTAL
        TempTable.VAL3 = .MAXPRICE
        TempTable.VAL4 = .MINPRICE
        
        TempTable.str7 = " ﬁ—Ì— ≈Ã„«·Ï „»Ì⁄«  ( „ﬁ«—‰… «”⁄«— )"
        If xStore.BoundText <> "" Then
            TempTable.str7 = TempTable.str7 & " ·„Œ“‰ " & xStore.Text
        End If
        TempTable.str8 = firsttitle
        TempTable.str9 = Secondtitle
        TempTable.Update
       .MoveNext
    Loop
End If
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\ritem_24.rpt"
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

