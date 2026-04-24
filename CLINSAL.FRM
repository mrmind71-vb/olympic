VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form ClinSal 
   Caption         =   "ÊÞÇÑíÑ ÇáÚãáÇÁ"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2370
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   1230
   End
   Begin VB.TextBox xVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   540
      Width           =   1230
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1155
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1230
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3840
      Top             =   1920
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
      Height          =   645
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   2385
      Begin VB.CommandButton CmdExit 
         Caption         =   "ÎÑæÌ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "ÚÑÖ"
         Height          =   375
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ãÈíÚÇÊ áÃÞá ãä"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ãÈíÚÇÊ ÃßÈÑ ãä"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Çáì :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Height          =   195
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   195
      Width           =   765
   End
End
Attribute VB_Name = "ClinSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TempTable As Recordset
Dim CodeTable As Recordset
Dim MoveClientTable As Recordset
Dim ClientTable As Recordset
Dim nOption As Integer
Function MYVALID()
If publicFlag = 1 Then
    If Not IsDate(xdate1.Text) Then Exit Function
    If Not IsDate(xDate2.Text) Then Exit Function
End If
If xVal1.Text = Null Then Exit Function
If xVal2.Text = Null Then Exit Function
MYVALID = True
End Function

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdundo_Click()
xdate1.Text = ""
xDate2.Text = ""
xVal1.Text = ""
xVal2.Text = ""
End Sub
Private Sub Form_Load()

Set TempTable = tempdb.OpenRecordset("TEMP")
Set ClientTable = mydb.OpenRecordset("FILE3_10")
Set CodeTable = mydb.OpenRecordset("FILE1_70")
Set MoveClientTable = mydb.OpenRecordset("SELECT * FROM FILE3_11 ORDER BY CODE , [TYPE] , [DATE] ")


If publicFlag = 2 Then
    xdate1.Visible = False
    xDate2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
End If
If publicFlag = 1 Then
    Label1.Caption = "ãÈíÚÇÊ ÃßÈÑ ãä"
    Label2.Caption = "ãÈíÚÇÊ ÃÞá  ãä"
Else
    Label1.Caption = "ÑÕíÏ ÃÚáì ãä"
    Label2.Caption = "ÑÕíÏ ÃÞá  ãä"
End If
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "Delete * From Temp"
If publicFlag = 1 Then
    cString = "SELECT FILE3_10.GROUP,FILE3_10.CODE, First(FILE3_10.DESCA) AS F_DESCA, Sum(FILE3_11.SAL) AS SumSAL , FILE1_70.DESCA  " & _
              " FROM (FILE3_10 INNER JOIN FILE3_11 ON FILE3_10.CODE = FILE3_11.CODE) LEFT JOIN FILE1_70 ON FILE3_10.group = FILE1_70.CODE " & _
              " WHERE FILE1_70.FLAG = 3 AND FILE3_11.TYPE ='4' " & _
              " AND Date >= DateValue('" & xdate1.Text & "')" & _
              " AND Date <= DateValue('" & xDate2.Text & "')" & _
              " GROUP BY FILE3_10.CODE,FILE3_10.GROUP , FILE1_70.DESCA " & _
              " ORDER BY FILE3_10.Group ,FILE3_10.CODE "
    Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
Else
    cString = "SELECT FILE3_10.GROUP,FILE3_10.CODE, First(FILE3_10.DESCA) AS F_DESCA, Sum(FILE3_11.SAL) AS SumSAL , Sum(FILE3_11.PAY) AS SumPay , FILE1_70.DESCA  " & _
              " FROM (FILE3_10 INNER JOIN FILE3_11 ON FILE3_10.CODE = FILE3_11.CODE) LEFT JOIN FILE1_70 ON FILE3_10.group = FILE1_70.CODE " & _
              " WHERE FILE1_70.FLAG = 3  " & _
              " GROUP BY FILE3_10.CODE,FILE3_10.GROUP , FILE1_70.DESCA " & _
              " ORDER BY FILE3_10.Group ,FILE3_10.CODE "
    Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
End If
With SourceTable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .CODE
    TempTable.str2 = .F_DescA
    If publicFlag = 1 Then
        TempTable.str3 = " ãä ÊÇÑíÎ " & xdate1.Text & " Åáì ÊÇÑíÎ " & xDate2.Text
        TempTable.VAL1 = TurnValue(.SUMSAL, Null, 0)
    Else
        TempTable.VAL1 = TurnValue(.SUMSAL, Null, 0) - TurnValue(.SUMPAY, Null, 0)
    End If
    TempTable.STR5 = .DESCA
    If publicFlag = 1 Then
        TempTable.str6 = "ÚãáÇÁ áåã ãÈíÚÇÊ ÃßÈÑ ãä  " & xVal1.Text & " æ ÃÞá ãä  " & xVal2.Text
        TempTable.str7 = "ÞíãÉ ÇáãÈíÚÇÊ"
    Else
        TempTable.str6 = "ÚãáÇÁ áåã ÃÑÕÏÉ ÃÚáì ãä  " & xVal1.Text & " æ ÃÞá ãä  " & xVal2.Text
        TempTable.str7 = "ÇáÑÕíÏ"
    End If
    TempTable.STR4 = .Group
    TempTable.str19 = firsttitle
    TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With

With TempTable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    If TempTable.VAL1 < Val(xVal1.Text) Or TempTable.VAL1 > Val(xVal2.Text) Then
        TempTable.Delete
    End If
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\ClSal.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Function myDateValue(pDate)
myDateValue = "#" & DateValue(pDate) & "#"
End Function
