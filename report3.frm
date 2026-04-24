VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form reportfrm3 
   Caption         =   "«⁄÷«¡ »œÊ‰ »Ì«‰«  «”«”Ì…"
   ClientHeight    =   3225
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   9780
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
   ScaleHeight     =   3225
   ScaleWidth      =   9780
   Begin VB.Frame Frame3 
      Height          =   2355
      Left            =   7065
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   135
      Width           =   2580
      Begin VB.CheckBox xMiss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰  «—ÌŒ „Ì·«œ"
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
         Height          =   315
         Index           =   0
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   180
         Width           =   1785
      End
      Begin VB.CheckBox xMiss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰  «—ÌŒ «· Õ«ﬁ"
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
         Height          =   315
         Index           =   1
         Left            =   585
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   495
         Width           =   1830
      End
      Begin VB.CheckBox xMiss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "⁄÷Ê »œÊ‰ ›∆…"
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
         Height          =   270
         Index           =   2
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   855
         Width           =   1785
      End
      Begin VB.CheckBox xMiss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ ⁄‰Ê«‰"
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
         Height          =   315
         Index           =   3
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1170
         Width           =   1965
      End
      Begin VB.CheckBox xMiss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ —ﬁ„  ·Ì›Ê‰"
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
         Height          =   315
         Index           =   4
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1560
         Width           =   1650
      End
      Begin VB.CheckBox xMiss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ „Ê»«Ì·"
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
         Height          =   315
         Index           =   5
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1935
         Width           =   1695
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
      Left            =   90
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   2520
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
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   2520
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
      TabIndex        =   10
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   2520
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Height          =   1230
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1260
      Width           =   2400
      Begin VB.CheckBox xSafe 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ…"
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
         Height          =   390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   765
         Value           =   1  'Checked
         Width           =   2130
      End
      Begin VB.CheckBox xDrop 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ…"
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
         Height          =   435
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   405
         Width           =   2040
      End
      Begin VB.CheckBox xDied 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ «·„ Ê›ÌÌ‰"
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
         Height          =   345
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   135
         Value           =   1  'Checked
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   4515
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   180
         Width           =   1680
      End
      Begin VB.TextBox xcode1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   180
         Width           =   1410
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   900
         Width           =   1410
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   900
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   540
         Width           =   3120
         _ExtentX        =   5503
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
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   1260
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   661
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
         Caption         =   "„‰ —ﬁ„"
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
         Index           =   7
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   225
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "›∆… «·⁄÷ÊÌ…"
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
         TabIndex        =   16
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "„”œœ „‰"
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
         Index           =   1
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "”œœ „Ê”„"
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
         TabIndex        =   14
         Top             =   1305
         Width           =   960
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1485
      Top             =   -90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   5940
      Top             =   2925
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
Attribute VB_Name = "reportfrm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty
Private Sub cmdApply_Click()
doprint
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Function doprint()
Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset, cOr As String
Dim aHeader(11)
cString = "Select file1_10.*,TYPE_CODES.DescA as type_Desca,dbo.f_last_year_date(FILE1_10.CODE) AS DATE_LAST " & _
          " From File1_10 left join TYPE_codes on File1_10.TYPE = TYPE_codes.Code"


If xMiss(0).Value = 1 Then
    cOr = "FILE1_10.DATE_BIRTH IS NULL"
    aHeader(0) = xMiss(0).Caption
End If

If xMiss(1).Value = 1 Then
    cOr = cOr & turn(cOr, " or ") & "FILE1_10.DATE_BEGIN IS NULL"
    aHeader(1) = xMiss(1).Caption
End If

If xMiss(2).Value = 1 Then
    cOr = cOr & turn(cOr, " or ") & "FILE1_10.TYPE IS NULL"
    aHeader(2) = xMiss(2).Caption
End If

If xMiss(3).Value = 1 Then
    cOr = cOr & turn(cOr, " or ") & "FILE1_10.ADDRESS IS NULL"
    aHeader(3) = xMiss(3).Caption
End If

If xMiss(4).Value = 1 Then
    cOr = cOr & turn(cOr, " or ") & "FILE1_10.PHONE IS NULL"
    aHeader(4) = xMiss(4).Caption
End If

If xMiss(5).Value = 1 Then
    cOr = cOr & turn(cOr, " or ") & "FILE1_10.MOBIL IS NULL"
    aHeader(5) = xMiss(5).Caption
End If

If cOr <> "" Then cWhere = "(" & cOr & ")"

If ValidNum(xcode1.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE1_10.code " & IIf(ValidNum(xCode2.text), " >= ", " = ") & addvalue(xcode1.text)
    aHeader(6) = IIf(ValidNum(xCode2.text), BetweenString(xcode1.text, xCode2.text, "„‰ —ﬁ„ ⁄÷ÊÌ… : ", "Õ Ì —ﬁ„ ⁄÷ÊÌ… : "), "—ﬁ„ ⁄÷ÊÌ… :" & xcode1.text)
End If

If ValidNum(xCode2.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE1_10.code <= " & addvalue(xCode2.text)
    aHeader(6) = BetweenString(xcode1.text, xCode2.text, "„‰ —ﬁ„ ⁄÷ÊÌ… : ", "Õ Ì —ﬁ„ ⁄÷ÊÌ… : ")
End If

If IsNumeric(cmdYear(0).Tag) Then
     aHeader(7) = "«·–Ì‰ ”œœÊ« „Ê”„ " & cmdYear(0).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) >= " & cmdYear(0).Tag
End If

If xType.MatchedWithList Then
    aHeader(8) = "›∆… «·⁄÷ÊÌ… : " & xType.text
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.TYPE = " & addvalue(xType.BoundText)
End If

If IsDate(xDate1.text) Then
    aHeader(9) = "„”œœ " & BetweenString(xDate1.text, xDate2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE WHERE DATE >= " & DateSq(xDate1.text) & " AND (NOT FORM_NO IS NULL) AND PAID_TYPES.IS_PAID = 1 )"
End If

If IsDate(xDate2.text) Then
    aHeader(9) = "„”œœ " & BetweenString(xDate1.text, xDate2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H WHERE DATE <= " & DateSq(xDate2.text) & " AND (NOT FORM_NO IS NULL))"
End If

If xSafe.Value = 0 Then
    cWhere = cWhere & turn(cWhere, " and ") & " (dbo.f_save(FILE1_10.CODE) = 0)"
    'aHeader(5) = "»œÊ‰ Õ«›ŸÌ «·⁄÷ÊÌ…"
Else
    'aHeader(5) = "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ…"
End If

If xDrop.Value = 0 Then
    cWhere = cWhere & turn(cWhere, " and ") & " (FILE1_10.[DROP] = 0)"
   ' aHeader(6) = "»œÊ‰ ”«ﬁÿÌ «·⁄÷ÊÌ…"
Else
   ' aHeader(6) = "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ…"
End If

If xDied.Value = 0 Then
   cWhere2 = cWhere & turn(cWhere, " and ") & " (Died = 0)"
  '  aHeader(10) = "»œÊ‰ «·„ Ê›ÌÌ‰"
Else
 '   aHeader(10) = "⁄—÷ «·„ Ê›ÌÌ‰"
End If


If cWhere2 <> "" Then
    cString = cString & " WHERE " & cWhere2
ElseIf cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

Me.MousePointer = 11
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!val1 = sourcetable!code
    temptable!str1 = sourcetable!code
    temptable!str2 = sourcetable!Desca
    temptable!str3 = TurnValue(sourcetable!TYPE_desca)
    temptable!str4 = TurnValue(ArbStr(myFormat_p(sourcetable!date_begin)))
    temptable!str5 = TurnValue(ArbStr(myFormat_p(sourcetable!DATE_LAST)))

    temptable!str10 = TurnValue(Me.Caption)
    temptable!str11 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!str12 = TurnValue(retHeader(aHeader, 3, 5))
    temptable.Update
    sourcetable.MoveNext
Loop

temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    Me.MousePointer = 0
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    
    Report1.ReportFileName = sPath_App & "\REPORTS\REPORT3.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
    Me.MousePointer = 0
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

Set data1.Recordset = myRecordSet("SELECT CODE,DESCA FROM TYPE_CODES ORDER BY CODE", con)
Set xType.RowSource = data1
xType.ListField = "Desca"
xType.BoundColumn = "Code"

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
End If
End Sub

Private Sub xCode1_GotFocus()
myGotFocus xcode1
End Sub
Private Sub xCode1_LostFocus()
myLostFocus xcode1
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xCode2
End Sub

