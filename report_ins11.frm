VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form report_insfrm11 
   Caption         =   "«»‰«¡ «⁄÷«¡ „ﬁ”ÿÌ‰ Ê’·Ê« «·Ì ”‰ „⁄Ì‰"
   ClientHeight    =   2565
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   4920
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
   ScaleHeight     =   2565
   ScaleWidth      =   4920
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
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1890
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
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   1890
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
      TabIndex        =   3
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   1890
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Width           =   4605
      Begin VB.TextBox xDate 
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
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   585
         Width           =   1410
      End
      Begin VB.TextBox xAge 
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
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "D"
         Top             =   225
         Width           =   1410
      End
      Begin MSDataListLib.DataCombo xGender 
         Height          =   330
         Left            =   270
         TabIndex        =   2
         Top             =   945
         Width           =   2985
         _ExtentX        =   5265
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
      Begin MSDataListLib.DataCombo xStatus 
         Height          =   330
         Left            =   270
         TabIndex        =   10
         Top             =   1305
         Width           =   2985
         _ExtentX        =   5265
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
         Index           =   0
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "„‰  «—ÌŒ"
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
         Index           =   3
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "«·‰Ê⁄"
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
         Index           =   1
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   990
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "«·”‰"
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
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
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
Attribute VB_Name = "report_insfrm11"
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
Private Function doprint()
Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset, cOr As String
Dim aHeader(11)
cString = "Select FILE2_10.CODE,FILE2_10.DESCA,FILE2_11.DESCA AS REL_DESCA,FILE2_11.DATE_BIRTH,FILE2_10.DATE_BEGIN,FILE2_10.DATE_END,FILE2_10.MOBIL " & _
          " From FILE2_10 INNER JOIN FILE2_11 ON FILE2_10.CODE = FILE2_11.MEMBER "

cWhere = "FILE2_11.RELATION = 2"

If ValidNum(xAge.text) Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_11.DATE_BIRTH <= " & DateSq(DateAdd("yyyy", -1 * mRound(xAge.text), myFormat(xDate.text)))
    aHeader(0) = "«»‰«¡ Ê’·Ê« «·Ì ”‰ " & xAge.text & " ”‰…"
End If

If xGender.MatchedWithList Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_11.GENDER = " & addvalue(xGender.BoundText)
    aHeader(1) = aHeader(0) & turn(aHeader(0), " ") & " (" & xGender.text & ")"
End If


If xStatus.MatchedWithList Then
    aHeader(2) = "Õ«·… «·⁄÷ÊÌ… : " & xStatus.text
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.STATUS = " & addvalue(xStatus.BoundText)
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If


sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!val1 = sourcetable!CODE
    temptable!str1 = ArbString(sourcetable!CODE)
    temptable!str2 = sourcetable!rel_Desca
    temptable!str3 = sourcetable!Desca
    temptable!str4 = TurnValue(ArbString(myFormat_p(sourcetable!DATE_BIRTH)))
    If IsDate(sourcetable!DATE_BIRTH) And IsDate(xDate.text) Then
        temptable!str5 = TurnValue(ArbString(AgeString(myFormat(sourcetable!DATE_BIRTH), myFormat(xDate.text))))
    End If
    temptable!str6 = TurnValue(ArbString(myFormat_p(sourcetable!DATE_BEGIN)))
    temptable!str7 = TurnValue(ArbString(myFormat_p(sourcetable!DATE_END)))
    temptable!str8 = TurnValue(ArbString(sourcetable!MOBIL))
    temptable!str10 = TurnValue(Me.Caption)
    temptable!str11 = TurnValue(retHeader(aHeader, 0, 5))
    temptable!str12 = TurnValue(retHeader(aHeader, 5, 5))
    temptable.Update
    sourcetable.MoveNext
Loop

temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    
    Report1.ReportFileName = sPath_App & "\REPORTS\REPORT_INS11.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set sourcetable = Nothing
End With
End Function
Private Sub Form_Load()
openCon con

Set data1.Recordset = myRecordSet("SELECT CODE,DESCA FROM GENDER_CODES ORDER BY CODE", con)
Set xGender.RowSource = data1
xGender.ListField = "Desca"
xGender.BoundColumn = "Code"

Set data2.Recordset = myRecordSet("select * from status_Codes", con)
Set xStatus.RowSource = data2
xStatus.ListField = "Desca"
xStatus.BoundColumn = "Code"

xDate.text = myFormat_p(Date)

xStatus.BoundText = 1

FixRpImage Me
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xdate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xGender_GotFocus()
myGotFocus xGender
End Sub
Private Sub xGender_LostFocus()
myLostFocus xGender
If Not xGender.MatchedWithList Then xGender.BoundText = ""
End Sub
Private Sub xStatus_GotFocus()
myGotFocus xStatus
End Sub
Private Sub xStatus_LostFocus()
myLostFocus xStatus
If Not xStatus.MatchedWithList Then xStatus.BoundText = ""
End Sub

Private Sub xAge_GotFocus()
myGotFocus xAge
End Sub
Private Sub xAge_LostFocus()
myLostFocus xAge
End Sub


