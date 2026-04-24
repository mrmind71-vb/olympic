VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form report_insfrm10 
   Caption         =   "„ «»⁄… „ﬁœ„ «·«ﬁ”«ÿ ··⁄÷ÊÌ… «·„ﬁ”ÿ…"
   ClientHeight    =   2775
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   7065
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
   ScaleHeight     =   2775
   ScaleWidth      =   7065
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   90
      Width           =   6810
      Begin VB.TextBox xCode 
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
         Left            =   4095
         MaxLength       =   9
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "N"
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
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
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   1035
      End
   End
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
      Left            =   2430
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   2115
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
      Left            =   3960
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   2115
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
      Left            =   5490
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   2115
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   6810
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "D"
         Top             =   540
         Width           =   1410
      End
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
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   540
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "D"
         Top             =   180
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
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   180
         Width           =   1410
      End
      Begin MSDataListLib.DataCombo xStatus 
         Height          =   330
         Left            =   2655
         TabIndex        =   5
         Top             =   900
         Width           =   2850
         _ExtentX        =   5027
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
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   945
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
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   585
         Width           =   915
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
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   1140
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
Attribute VB_Name = "report_insfrm10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchMember As New Search
Private Sub cmdApply_Click()
doprint
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Function doprint(Optional pReport As String = "")
Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset, cOr As String
Dim aHeader(11)
cString = "SELECT  FILE6_21.DATE_DUE,FILE6_21.CODE,FILE2_10.DESCA,FILE6_21.VALUE,FILE6_21.DATE_PAID,FILE6_21.TAX + FILE6_21.TAX_DIFF,FILE6_21.VALUE +  FILE6_21.TAX +  FILE6_21.TAX_DIFF AS TOTAL,FILE2_10.PHONE" & _
          " FROM FILE6_21 INNER JOIN FILE2_10 ON FILE6_21.CODE = FILE2_10.CODE AND FILE6_21.ID =  [dbo].[f_install_first](FILE2_10.CODE)"


If xStatus.MatchedWithList Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.STATUS = " & addvalue(xStatus.BoundText)
End If

If IsDate(xdate_begin1.text) Then
    aHeader(0) = " «—ÌŒ »œ«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(myFormat_p(xdate_begin1.text), myFormat_p(xDate_Begin2.text))
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.DATE_BEGIN >= " & DateSq(xdate_begin1.text)
End If


If IsDate(xDate_Begin2.text) Then
    aHeader(0) = " «—ÌŒ »œ«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(myFormat_p(xdate_begin1.text), myFormat_p(xDate_Begin2.text))
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.DATE_BEGIN <= " & DateSq(xDate_Begin2.text)
End If

If IsDate(xDate_end1.text) Then
    aHeader(1) = " «—ÌŒ ‰Â«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(xDate_end1.text, xDate_End2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.date_end >= " & DateSq(xDate_end1.text)
End If

If IsDate(xDate_End2.text) Then
    aHeader(2) = " «—ÌŒ ‰Â«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(xDate_end1.text, xDate_End2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.date_end <= " & DateSq(xDate_End2.text)
End If

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!val1 = sourcetable!CODE
    temptable!str1 = ArbString(sourcetable!CODE)
    temptable!str2 = sourcetable!Desca
    temptable!str3 = TurnValue(myFormat_p(sourcetable!DATE_DUE))
    temptable!str4 = sourcetable!Phone
    
    temptable!val2 = mRound(sourcetable!Value)
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
    Report1.ReportFileName = sPath_App & "\REPORTS\REPORT_INS10.rpt"
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
xStatus.BoundText = 1
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

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then MemberLookupInstall Me, oSearchMember
End Sub

Private Sub xCode_LostFocus()
xCodeDesca.Caption = ""

If Not ValidNum(xCode.text) Then Exit Sub
Dim aMember As Variant
aMember = Member_Load_install(xCode.text, , con)
xCodeDesca.Caption = retFlag(aMember, "Desca") & ""
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
Public Sub myProc()
If ActiveControl.Name = xCode.Name Then
    xCode.text = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
    xCodeDesca.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 1)
    Unload oSearchMember
End If
End Sub

