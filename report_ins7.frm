VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form report_insfrm7 
   Caption         =   "«ﬁ”«ÿ ·«  ”«ÊÌ ﬁÌ„… ⁄ﬁœ «·⁄÷Ê «·„ﬁ”ÿ"
   ClientHeight    =   2145
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2145
   ScaleWidth      =   4710
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1485
      Width           =   1500
   End
   Begin VB.CommandButton cmdClear 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1665
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   1485
      Width           =   1500
   End
   Begin VB.CommandButton CmdApply 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3195
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   1485
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   90
      Width           =   4515
      Begin VB.TextBox xDate_end1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "D"
         Top             =   585
         Width           =   1410
      End
      Begin VB.TextBox xDate_End2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   585
         Width           =   1410
      End
      Begin VB.TextBox xDate_Begin2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox xdate_begin1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "D"
         Top             =   225
         Width           =   1410
      End
      Begin MSDataListLib.DataCombo xStatus 
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   945
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Õ«·… «·⁄÷Ê"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   990
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ ‰Â«Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ »œ«Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   1320
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1485
      Top             =   -450
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   -360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   -360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   0
      Top             =   -360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
Attribute VB_Name = "report_insfrm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty, oSearchRegion As New Search
Private Sub cmdApply_Click()
doprint
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Function doprint(Optional pReport As String = "")
Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset, cOr As String
Dim aHeader(11)
cString = "Select file2_10.CODE,FILE2_10.DESCA,dbo.f_last_year_date_install(FILE2_10.CODE) AS LAST_DATE_PAID,INSTALL_CODES.DESCA AS INSTALL_DESCA,FILE2_10.DATE_END,FILE2_10.INSTALL_COUNT,FILE2_10.VALUE,FILE2_10.INSTALL_VALUE,SUM(INSTALL_BALANCE.VALUE) AS install_total,SUM(INSTALL_BALANCE.INS_COUNT) AS ins_count  " & _
          " From File2_10 LEFT JOIN INSTALL_BALANCE ON FILE2_10.CODE = INSTALL_BALANCE.CODE LEFT JOIN INSTALL_CODES ON FILE2_10.INSTALL_TYPE = INSTALL_CODES.CODE "

If xStatus.MatchedWithList Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.STATUS = " & addvalue(xStatus.BoundText)
End If

If IsDate(xdate_begin1.Text) Then
    aHeader(2) = " «—ÌŒ »œ«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(myFormat_p(xdate_begin1.Text), myFormat_p(xDate_Begin2.Text))
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.DATE_BEGIN >= " & DateSq(xdate_begin1.Text)
End If

If IsDate(xDate_Begin2.Text) Then
    aHeader(2) = " «—ÌŒ »œ«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(myFormat_p(xdate_begin1.Text), myFormat_p(xDate_Begin2.Text))
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.DATE_BEGIN <= " & DateSq(xDate_Begin2.Text)
End If

If IsDate(xDate_end1.Text) Then
    aHeader(3) = " «—ÌŒ ‰Â«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(xDate_end1.Text, xDate_End2.Text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.date_end >= " & DateSq(xDate_end1.Text)
End If

If IsDate(xDate_End2.Text) Then
    aHeader(3) = " «—ÌŒ ‰Â«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(xDate_end1.Text, xDate_End2.Text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.date_end <= " & DateSq(xDate_End2.Text)
End If


If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " group by file2_10.CODE,FILE2_10.DESCA,FILE2_10.DATE_BEGIN,FILE2_10.DATE_END,FILE2_10.INSTALL_COUNT,FILE2_10.VALUE,FILE2_10.INSTALL_VALUE,INSTALL_CODES.DESCA,FILE2_10.PHONE,FILE2_10.ADDRESS"
cHaving = "FILE2_10.Value <> Sum(INSTALL_BALANCE.Value)"

If cHaving <> "" Then cString = cString & " HAVING " & cHaving

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!val1 = sourcetable!CODE
    temptable!str1 = ArbString(sourcetable!CODE)
    temptable!str2 = sourcetable!Desca
    'temptable!str3 = TurnValue(ArbString(myFormat_p(sourcetable!date_begin)))
    temptable!str3 = sourcetable!install_desca
    temptable!str5 = TurnValue(ArbString(myFormat_p(sourcetable!LAST_DATE_PAID)))
    'temptable!str6 = TurnValue(ArbString(sourcetable!Address & ""))
    'temptable!str7 = TurnValue(sourcetable!Phone)
    temptable!VAL3 = mRound(sourcetable!ins_count)
    temptable!val4 = mRound(sourcetable!Value)
    temptable!VAL5 = mRound(sourcetable!Install_total)
    temptable!VAL6 = mRound(sourcetable!Value) - mRound(sourcetable!Install_total)
    temptable!str10 = TurnValue(Me.Caption)
    temptable!str11 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!str12 = TurnValue(retHeader(aHeader, 3, 5))
    temptable.Update
    sourcetable.MoveNext
Loop

temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    If pReport = "" Then
        Report1.ReportFileName = sPath_App & "\REPORTS\REPORT_INS7.rpt"
    Else
        Report1.ReportFileName = sPath_App & "\REPORTS\" & pReport
    End If
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set sourcetable = Nothing
End With
End Function
Private Sub Form_Load()
openCon con

Set data1.Recordset = myRecordSet("select * from status_Codes", con)
Set xStatus.RowSource = data1
xStatus.ListField = "Desca"
xStatus.BoundColumn = "Code"

FixRpImage Me

LoadText Me
'xStatus.BoundText = 1

End Sub

Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
If Not xType.MatchedWithList Then xType.BoundText = ""
End Sub
Private Sub xJob_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    'Job_Lookup Me, oSearchJob
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
Set report_insfrm5 = Nothing
End Sub

Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xdate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xDate_end1_GotFocus()
myGotFocus xDate_end1
End Sub
Private Sub xDate_end1_LostFocus()
myLostFocus xDate_end1
myValidDate xDate_end1
End Sub
Private Sub xDate_End2_GotFocus()
myGotFocus xDate_End2
End Sub
Private Sub xDate_End2_LostFocus()
myLostFocus xDate_End2
myValidDate xDate_End2
End Sub
Private Sub xdate_begin2_GotFocus()
myGotFocus xDate_Begin2
End Sub
Private Sub xdate_begin2_LostFocus()
myLostFocus xDate_Begin2
myValidDate xDate_Begin2
End Sub
Private Sub xDate_begin1_GotFocus()
myGotFocus xdate_begin1
End Sub
Private Sub xDate_begin1_LostFocus()
myLostFocus xdate_begin1
myValidDate xdate_begin1
End Sub
Private Sub xStatus_GotFocus()
myGotFocus xStatus
End Sub
Private Sub xStatus_LostFocus()
myLostFocus xStatus
If Not xStatus.MatchedWithList Then xStatus.BoundText = ""
End Sub
