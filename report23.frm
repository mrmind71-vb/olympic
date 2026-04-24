VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form reportfrm23 
   Caption         =   "«⁄÷«¡ ›«’·Ì «·⁄÷ÊÌ…"
   ClientHeight    =   2235
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   4845
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
   ScaleHeight     =   2235
   ScaleWidth      =   4845
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
      Left            =   90
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
      Left            =   1620
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
      Left            =   3150
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   1485
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   4605
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   225
         Width           =   1680
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
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "D"
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox xDate1 
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
         TabIndex        =   3
         Tag             =   "D"
         Top             =   585
         Width           =   1410
      End
      Begin VB.TextBox xDate2 
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "D"
         Top             =   585
         Width           =   1680
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   420
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   945
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   741
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— «·„Ê”„"
         ButtonStyle     =   3
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ «·«· Õ«ﬁ"
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
         TabIndex        =   11
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ ”œ«œ „‰"
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
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "›Ì „Ê”„"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1035
         Width           =   960
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
Attribute VB_Name = "reportfrm23"
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
cString = "Select file1_10.*,TYPE_CODES.DESCA AS TYPE_DESCA,SES_NO " & _
          " From File1_10 LEFT JOIN TYPE_CODES ON FILE1_10.TYPE = TYPE_CODES.CODE"

If IsNumeric(cmdYear(0).Tag) Then
     aHeader(1) = "«·–Ì‰ ”œœÊ« „Ê”„ " & cmdYear(0).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) >= " & cmdYear(0).Tag
End If


If IsDate(xDate1.Text) Then
    aHeader(2) = "„”œœ " & BetweenString(xDate1.Text, xDate2.Text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE WHERE DATE >= " & DateSq(xDate1.Text) & " AND (NOT FORM_NO IS NULL) AND PAID_TYPES.IS_PAID = 1 )"
End If

If IsDate(xDate2.Text) Then
    aHeader(2) = "„”œœ " & BetweenString(xDate1.Text, xDate2.Text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H WHERE DATE <= " & DateSq(xDate2.Text) & " AND (NOT FORM_NO IS NULL))"
End If

If IsDate(xdate_begin1.Text) Then
    aHeader(3) = " «—ÌŒ «· Õ«ﬁ „‰ " & BetweenString(xdate_begin1.Text, xDate_Begin2.Text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.DATE_BEGIN >= " & DateSq(xdate_begin1.Text)
End If

If IsDate(xDate_Begin2.Text) Then
    aHeader(3) = " «—ÌŒ «· Õ«ﬁ «·Ì " & BetweenString(xdate_begin1.Text, xDate_Begin2.Text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.DATE_BEGIN <= " & DateSq(xDate_Begin2.Text)
End If

If xSafe.Value = 0 Then
    cWhere = cWhere & turn(cWhere, " and ") & " (dbo.f_save(FILE1_10.CODE) = 0)"
    aHeader(4) = "»œÊ‰ Õ«›ŸÌ «·⁄÷ÊÌ…"
Else
    'aHeader(5) = "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ…"
End If

If xDrop.Value = 0 Then
    cWhere = cWhere & turn(cWhere, " and ") & " (FILE1_10.[DROP] = 0)"
   ' aHeader(6) = "»œÊ‰ ”«ﬁÿÌ «·⁄÷ÊÌ…"
Else
    aHeader(3) = "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ…"
End If

If xDied.Value = 0 Then
   cWhere2 = cWhere & turn(cWhere, " and ") & " (Died = 0)"
  '  aHeader(10) = "»œÊ‰ «·„ Ê›ÌÌ‰"
Else
    aHeader(3) = "⁄—÷ «·„ Ê›ÌÌ‰"
End If

If cWhere2 <> "" Then
    cString = cString & " WHERE " & cWhere2
ElseIf cWhere <> "" Then
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
    temptable!str2 = sourcetable!desca
    temptable!str3 = TurnValue(ArbString(myFormat_p(sourcetable!date_begin)))
    temptable!str4 = TurnValue(ArbString(sourcetable!SES_NO))
    temptable!str5 = sourcetable!TYPE_desca
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
    
    Report1.ReportFileName = sPath_App & "\REPORTS\REPORT8.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set sourcetable = Nothing
End With
End Function

Private Sub cmdYear_Click(Index As Integer)
Years_LookupAll Me, oSearchYear, , cmdYear(Index).Tag <> ""
End Sub
Private Sub Form_Load()
openCon con

FixRpImage Me
End Sub

Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDate_begin1_GotFocus()
myGotFocus xdate_begin1
End Sub
Private Sub xDate_begin1_LostFocus()
myLostFocus xdate_begin1
myValidDate xdate_begin1
End Sub
Private Sub xdate_begin2_GotFocus()
myGotFocus xDate_Begin2
End Sub
Private Sub xdate_begin2_LostFocus()
myLostFocus xDate_Begin2
myValidDate xDate_Begin2
End Sub

Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
If Not xType.MatchedWithList Then xType.BoundText = ""
End Sub
Sub myProc()
If ActiveControl.Name = cmdYear(0).Name Then
    ActiveControl.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
    ActiveControl.Caption = IIf(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0) = "", "«Œ «— «·„Ê”„", oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
    oSearchYear.Hide
ElseIf ActiveControl.Name = XRegion.Name Then
    XRegion.BoundText = oSearchRegion.grid1.TextMatrix(oSearchRegion.grid1.Row, 0)
    oSearchRegion.Hide
End If
End Sub

Private Sub xJob_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    'Job_Lookup Me, oSearchJob
End If
End Sub
Private Sub xAddress_GotFocus()
myGotFocus xAddress
End Sub
Private Sub xAddress_LostFocus()
myLostFocus xAddress
myValidDate xAddress
End Sub
Private Sub xRegion_GotFocus()
myGotFocus XRegion
End Sub

Private Sub XRegion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then region_Lookup Me, oSearchRegion
End Sub

Private Sub xRegion_LostFocus()
myLostFocus XRegion
If Not XRegion.MatchedWithList Then XRegion.BoundText = ""
End Sub
