VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form MemberRep1 
   Caption         =   "»Ì«‰«  √⁄÷«¡ «·Ã„⁄Ì… «·⁄„Ê„Ì… ··‰«œÏ"
   ClientHeight    =   5190
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   9315
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
   ScaleHeight     =   5190
   ScaleWidth      =   9315
   Begin VB.Frame Frame5 
      Height          =   1545
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2835
      Width           =   9060
      Begin VB.TextBox xHeader 
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
         Height          =   960
         Left            =   45
         MaxLength       =   200
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   495
         Width           =   6855
      End
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   4005
         TabIndex        =   5
         Top             =   135
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«Ã „«⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7065
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄‰Ê«‰ «· ﬁ—Ì—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7065
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   540
         Width           =   1020
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "«Œ— ”œ«œ ›Ì"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   5355
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1755
      Width           =   3885
      Begin VB.TextBox xdate2 
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
         Left            =   225
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox xdate1 
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
         Left            =   225
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   630
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ·«· «·› —… „‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1050
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1755
      Width           =   5145
      Begin VB.CommandButton Command1 
         Caption         =   "√⁄÷«¡ ·«  ‰ÿ»ﬁ ⁄·ÌÂ„ «·‘—Êÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2610
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   585
         Width           =   2490
      End
      Begin VB.CommandButton Command2 
         Caption         =   "√⁄÷«¡ ·„ Ì„— ⁄·ÌÂ„ ”‰…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   135
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   " «—ÌŒ «· Õ«ﬁ  «»⁄ ﬁ»· «·⁄÷Ê"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   585
         Width           =   2490
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ﬂ· «·«⁄÷«¡ Ê«·“ÊÃ« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   135
         Width           =   2490
      End
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4410
      Width           =   6405
      Begin VB.CommandButton cmdPrint_pdf 
         Caption         =   "ÿ»«⁄… pdf"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1590
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton cmdPrint_test 
         Caption         =   "ÿ»«⁄…  ﬁ—Ì— „—«Ã⁄…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3195
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   1500
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "ÿ»«⁄… «·«⁄÷«¡"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4860
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   5355
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   45
      Width           =   3885
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
         Left            =   225
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox xCode1 
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
         Left            =   225
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1500
      End
      Begin VB.TextBox xDate 
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
         Left            =   225
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1260
         Width           =   1500
      End
      Begin MSDataListLib.DataCombo xYear_code 
         Height          =   330
         Left            =   225
         TabIndex        =   28
         Top             =   900
         Width           =   1500
         _ExtentX        =   2646
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·Ã„⁄Ì… «·⁄„Ê„Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1305
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ ⁄÷ÊÌ… „‰ —ﬁ„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï —ﬁ„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ «‘ —«ﬂ«  Õ Ì ”‰…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   945
         Width           =   1815
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1950
      Top             =   15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "MemberRep1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdOk_Click()
If Not MYVALID Then Exit Sub
doprint
End Sub
Private Sub doprint(Optional bTest As Boolean, Optional bBdf As Boolean = False)
Dim cString As String, cWhere As String
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT FILE1_10.CODE,FILE1_10.DESCA From File1_10 INNER JOIN FILE6_20H ON FILE1_10.DOC_NO = FILE6_20H.DOC_NO"
cWhere = " FILE1_10.DIED = 0"
cWhere = cWhere & " AND " & " FILE6_20H.TYPE = 1"
cWhere = cWhere & " AND " & " FILE6_20H.YEAR_CODE >= " & addvalue(xYear_code.BoundText)
cWhere = cWhere & " AND " & "(FILE1_10.DateBegin <= " & dateSql(DateAdd("yyyy", -1, xDate.Text))

If ValidNum(xCode1.Text) Then
    cWhere = cWhere & " AND " & " FILE1_10.code >= " & addvalue(xCode1.Text)
End If

If xCode2.Text <> "" Then
    cWhere = cWhere & " AND " & " FILE1_10.code >= " & addvalue(xCode1.Text)
End If

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Do Until sourcetable.EOF
    temptable.AddNew
    If Trim(xHeader.Text) <> "" Then
        temptable!str11 = TurnValue(xHeader.Text)
    Else
        temptable!str11 = "«·√⁄÷«¡ «·–Ì‰ ·Â„ Õﬁ Õ÷Ê— «·Ã„⁄Ì… «·⁄„Ê„Ì… «·„‰⁄ﬁœ… ›Ì " & myFormat_p(xDate.Text)
    End If
    temptable!str2 = !Desca
    temptable!str1 = !CODE
    temptable!val1 = !CODE
    temptable!val2 = 0
    temptable.Update
    sourcetable.MoveNext
Loop
End With

cString = "SELECT FILE1_11.Member,file1_11.code,File1_11.Desca,File1_11.Title " & _
          " from file1_11 inner join file1_10 on file1_11.member = file1_10.code INNER JOIN FILE6_20H ON FILE1_10.DOC_NO = FILE6_20H.DOC_NO"
cWhere = "FILE1_11.RELATION = 1"
cWhere = cWhere & " AND " & " FILE6_20H.TYPE = 1"
cWhere = cWhere & " AND " & " FILE6_20H.YEAR_CODE >= " & addvalue(xYear_code.BoundText)
cWhere = cWhere & " AND " & "(FILE1_11.DateBegin <= " & dateSql(DateAdd("yyyy", -1, xDate.Text))
          
If ValidNum(xCode1.Text) Then
    cWhere = cWhere & " AND " & " FILE1_10.code >= " & addvalue(xCode1.Text)
End If

If xCode2.Text <> "" Then
    cWhere = cWhere & " AND " & " FILE1_10.code >= " & addvalue(xCode1.Text)
End If

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
Do Until sourcetable.EOF
        temptable.AddNew
    If Trim(xHeader.Text) <> "" Then
        temptable!str11 = TurnValue(xHeader.Text)
    Else
        temptable!str11 = "«·√⁄÷«¡ «·–Ì‰ ·Â„ Õﬁ Õ÷Ê— «·Ã„⁄Ì… «·⁄„Ê„Ì… «·„‰⁄ﬁœ… ›Ì " & myFormat_p(xDate.Text)
    End If
        
        temptable!val1 = !Member
        temptable!str1 = !Member
        temptable!str2 = !Desca
        temptable!val2 = !CODE
        temptable.Update
    sourcetable.MoveNext
Loop
End With
If (temptable.EOF And temptable.BOF) Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
Else
    If bTest Then
       Report1.ReportFileName = MainPath & "\RPT\rpt1_1_1.rpt"
    ElseIf bBdf Then
        Report1.ReportFileName = MainPath & "\RPT\rpt1_1_pdf.rpt"
    Else
        Report1.ReportFileName = MainPath & "\RPT\rpt1_1.rpt"
    End If
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If
End Sub
Private Sub Report03_2()
Dim I As Integer
cString = "Select Code,Desca,Title,dateBegin From File1_10 Where Not isNull(DateBegin)"
Set sourcetable = mydb.OpenRecordset(cString)
tempdb.Execute "delete * from temp"
Set temptable = tempdb.OpenRecordset("SELECT  * FROM TEMP ", dbOpenDynaset)
With sourcetable
'SourceTable.MoveLast
'nRecordcount = SourceTable.RecordCount
sourcetable.MoveFirst
Do
'    If mydatediff( Format(xDate.Text, "dd-mm-yyyy"), Format()) >= 1 Then
'        TempTable.AddNew
'        TempTable.str11 = "«·Ã„⁄Ì… «·⁄„Ê„Ì… Õ Ì" & xDate.Text
'        TempTable.STR1 = TurnValue(!Title, Null, "") & IIf(NotIsNull(!Title), "/", "") & !Code
'        TempTable.STR2 = !Desca
'        TempTable!Val1 = 0
'        TempTable.Update
'    End If
'    SourceTable.MoveNext
Loop Until sourcetable.EOF
End With

cString = "Select Code,Desca,Title,dateBegin From File1_11 Where Not isNull(DateBegin) and Not isNull(DateBirth)"
Set sourcetable = mydb.OpenRecordset(cString)
Set temptable = tempdb.OpenRecordset("SELECT  * FROM TEMP ", dbOpenDynaset)
With sourcetable
sourcetable.MoveFirst
Do
    If myDateDiff(Format(xDate.Text, "dd-mm-yyyy"), Format(!datebegin, "dd-mm-yyyy")) >= 1 And _
    myDateDiff(Format(xDate.Text, "dd-mm-yyyy"), Format(!DateBirth, "dd-mm-yyyy")) >= 21 Then
        temptable.AddNew
        temptable.str11 = "«·Ã„⁄Ì… «·⁄„Ê„Ì… Õ Ì" & xDate.Text
        temptable.str1 = TurnValue(!Title, Null, "") & IIf(Not IsNull(!Title), "/", "") & !CODE
        temptable.str2 = !Desca
        temptable!val1 = 1
        temptable.Update
    End If
    sourcetable.MoveNext
Loop Until sourcetable.EOF
End With
Report1.WindowShowExportBtn = False
Report1.ReportFileName = MainPath & "\rpt05.rpt"
Report1.DataFiles(0) = tempFile
Report1.Action = 1: tempdb.Execute "Delete * from temp"
End Sub

Private Sub cmdPrint_pdf_Click()
Report03_1 , True
End Sub

Private Sub cmdPrint_test_Click()
Report03_1 True
End Sub

Private Sub Command1_Click()
Report03_3
End Sub

Private Sub Command2_Click()
Report03_4
End Sub

Private Sub Command3_Click()
Report03_5
End Sub

Private Sub Command4_Click()
Report03_6
End Sub

Private Sub Command5_Click()
Report03_1 True
End Sub

Private Sub Command6_Click()
End Sub
Private Sub Form_Load()
openCon con
LoadText Me
End Sub
Private Function MYVALID() As Boolean
If (Not Val(xYear.Text) > 0) Or Not IsNumeric(xYear.Text) Then
    MsgBox "«·”‰… €Ì— ’«·Õ…"
    Exit Function
End If
If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If Not IsNumeric(xCode1.Text) And Not MyEmpty(xCode1.Text) Then
    MsgBox "«·—ﬁ„ €Ì— ’ÕÌÕ"
    Exit Function
End If
If Not IsNumeric(xCode2.Text) And Not MyEmpty(xCode2.Text) Then
    MsgBox "«·—ﬁ„ €Ì— ’ÕÌÕ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub Report03_3()
If Not MYVALID Then Exit Sub
cString = "Select Code,dESce,Desca,Title,dateBegin,DateLast,Died,isSave From File1_10 "

cOr = cOr & turnFound(cOr, " ( ", " or ", " ( ") & "  isSave"
cOr = cOr & turnFound(cOr, " ( ", " or ", " ( ") & "  Died"
cOr = cOr & turnFound(cOr, " ( ", " or ", " ( ") & " dateLast < " & dateSql(paidDate(xYear.Text))
cOr = cOr & turnFound(cOr, " ( ", " or ", " ( ") & " DateBegin > " & dateSql(DateAdd("yyyy", -1, xDate.Text))
cOr = cOr & turnFound(cOr, " ( ", " ) ", "")
If xCode1.Text <> "" Then
    cString = cString & turnFound(cString) & " code >= " & xCode1.Text
End If

If xCode2.Text <> "" Then
    cString = cString & turnFound(cString) & " code <= " & xCode2.Text
End If

cString = cString & turnFound(cString) & cOr

Set sourcetable = mydb.OpenRecordset(cString)
tempdb.Execute "delete * from temp"
Set temptable = tempdb.OpenRecordset("SELECT  * FROM TEMP ", dbOpenDynaset)
If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
    Exit Sub
End If
With sourcetable
sourcetable.MoveFirst
Do
   temptable.AddNew
    temptable.str11 = "«·√⁄÷«¡ «·–Ì‰ ·Ì” ·Â„ Õﬁ Õ÷Ê— «·Ã„⁄Ì… «·⁄„Ê„Ì… «·„‰⁄ﬁœ… ›Ì " & Format(xDate.Text, "yyyy/m/d")
    temptable.str1 = Chr(254) & !CODE
    temptable.str2 = !Desca
    temptable.Str3 = !DESCE
    temptable.str4 = "⁄÷Ê"
    temptable.str5 = IIf(!died, "‰⁄„", "·«")
    temptable.STR6 = IIf(!IsSave, "‰⁄„", "·«")
    
    temptable!date1 = !datebegin
    temptable!DATE2 = !datelast

    temptable!val1 = !CODE
    temptable!Str20 = Chr(254)
    temptable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
End With

cString = "Select Member,fILE1_11.dESCe,File1_11.Desca,File1_11.Title,File1_11.dateBegin,File1_11.datebirth,dateLast,isSave From File1_10 Inner Join File1_11 On File1_10.Code = File1_11.Member "
cString = cString & " where Relation = '1'"

If xCode1.Text <> "" Then
    cString = cString & turnFound(cString) & " Member >= " & xCode1.Text
End If

If xCode2.Text <> "" Then
    cString = cString & turnFound(cString) & " Member <= " & xCode2.Text
End If
cOr = ""
cOr = cOr & turnFound(cOr, " ( ", " or ", " ( ") & "  isSave"
cOr = cOr & turnFound(cOr, " ( ", " or ", " ( ") & "  Died"
cOr = cOr & turnFound(cOr, " ( ", " or ", " ( ") & " dateLast < " & dateSql(paidDate(xYear.Text))
cOr = cOr & turnFound(cOr, " ( ", " or ", " ( ") & " File1_11.DateBegin > " & dateSql(DateAdd("yyyy", -1, xDate.Text))
cOr = cOr & turnFound(cOr, " ( ", " ) ", "")

cString = cString & turnFound(cString) & cOr

Set sourcetable = mydb.OpenRecordset(cString)
Set temptable = tempdb.OpenRecordset("SELECT  * FROM TEMP ", dbOpenDynaset)

With sourcetable
sourcetable.MoveFirst
Do
    temptable.AddNew
    temptable.str11 = "«·√⁄÷«¡ «·–Ì‰ ·Ì” ·Â„ Õﬁ Õ÷Ê— «·Ã„⁄Ì… «·⁄„Ê„Ì… «·„‰⁄ﬁœ… ›Ì " & Format(xDate.Text, "yyyy/m/d")
    temptable.str2 = !Desca
    temptable.Str3 = !DESCE
    temptable!val1 = !Member
    temptable.str4 = "“ÊÃ «Ê “ÊÃ…"
    temptable.STR6 = IIf(!IsSave, "‰⁄„", "·«")
    temptable!date1 = !datebegin
    temptable!DATE2 = !datelast

    temptable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
End With
Report1.WindowShowExportBtn = False
Report1.ReportFileName = MainPath & "\RPT\rpt1_2.rpt"
Report1.DataFiles(0) = tempFile
Report1.Action = 1: tempdb.Execute "Delete * from temp"
End Sub
Private Function TurnOr(cString, bLast)
If InStr(1, LCase(cString), " or ") = 0 Then
    TurnOr = " where "
    Exit Function
End If

If InStr(1, LCase(cString), " and ") <> 0 Then
    TurnOr = " and ( "
    Exit Function
End If

If InStr(1, LCase(cString), " or ") <> 0 Then
    TurnOr = " Or "
    Exit Function
End If
End Function
Private Sub Report03_4()
If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Sub
End If
    
cString = "Select Code,dESce,Desca,Title,dateBegin,DateLast From File1_10 "
cString = cString & turnFound(cString) & " DateBegin >= " & dateSql(DateAdd("yyyy", -1, xDate.Text))

If xCode1.Text <> "" Then
    cString = cString & turnFound(cString) & " code >= " & xCode1.Text
End If

If xCode2.Text <> "" Then
    cString = cString & turnFound(cString) & " code <= " & xCode2.Text
End If

Set sourcetable = mydb.OpenRecordset(cString & cWhere)
tempdb.Execute "delete * from temp"
Set temptable = tempdb.OpenRecordset("SELECT  * FROM TEMP ", dbOpenDynaset)
If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
    Exit Sub
End If
With sourcetable
sourcetable.MoveFirst
Do
    temptable.AddNew
    temptable.str11 = "«·«⁄÷«¡ «·–Ì‰ ·„ Ì„— ⁄·ÌÂ„ ”‰…" & Format(xDate.Text, "yyyy/m/d")
    temptable.str2 = !Desca
    temptable.Str3 = !DESCE
    temptable!val1 = !CODE
    temptable!val2 = 0
    temptable!Val3 = !CODE
    temptable!Str3 = "⁄÷Ê"
    temptable!Str20 = Chr(254)
    temptable!date1 = !datebegin
    temptable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
End With

cString = "Select Member,fILE1_11.dESCe,File1_11.Desca,File1_11.Title,File1_11.dateBegin,File1_11.datebirth,dateLast,isSave From File1_10 Inner Join File1_11 On File1_10.Code = File1_11.Member "
cString = cString & " where Relation = '1'"
cString = cString & turnFound(cString) & " File1_11.DateBegin >= " & dateSql(DateAdd("yyyy", -1, xDate.Text))
       

If xCode1.Text <> "" Then
    cString = cString & turnFound(cString) & " Member >= " & xCode1.Text
End If

If xCode2.Text <> "" Then
    cString = cString & turnFound(cString) & " Member <= " & xCode2.Text
End If

Set sourcetable = mydb.OpenRecordset(cString)
Set temptable = tempdb.OpenRecordset("SELECT  * FROM TEMP ", dbOpenDynaset)

With sourcetable
sourcetable.MoveFirst
Do
    temptable.AddNew
    temptable.str11 = "«·«⁄÷«¡ «·–Ì‰ ·„ Ì„— ⁄·ÌÂ„ ”‰…" & Format(xDate.Text, "yyyy/m/d")
    temptable.str2 = !Desca
    temptable.Str3 = !DESCE
    temptable!Str3 = " «»⁄"
    temptable!val1 = !Member
    temptable!val2 = 1
    temptable!Val3 = 0
    temptable!date1 = !datebegin
    temptable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
End With
Report1.WindowShowExportBtn = False
Report1.ReportFileName = MainPath & "\RPT\rep1_3.rpt"
Report1.DataFiles(0) = tempFile
Report1.Action = 1: tempdb.Execute "Delete * from temp"
End Sub
Private Sub Report03_5()
tempdb.Execute "delete * from temp"
Set temptable = tempdb.OpenRecordset("SELECT  * FROM TEMP ", dbOpenDynaset)
cString = "Select Member,fILE1_11.dESCe,FILE1_10.DATEBEGIN,File1_11.Desca,File1_11.Title,File1_11.dateBegin,File1_11.datebirth,dateLast,isSave From File1_10 Inner Join File1_11 On File1_10.Code = File1_11.Member "
cString = cString & " where Relation = '1'"
cString = cString & turnFound(cString) & " File1_11.DateBegin >= " & dateSql(DateAdd("yyyy", -1, xDate.Text))
cString = cString & turnFound(cString) & " File1_10.DateBegin > File1_11.DateBegin"
       

If xCode1.Text <> "" Then
    cString = cString & turnFound(cString) & " Member >= " & xCode1.Text
End If

If xCode2.Text <> "" Then
    cString = cString & turnFound(cString) & " Member <= " & xCode2.Text
End If

Set sourcetable = mydb.OpenRecordset(cString)
Set temptable = tempdb.OpenRecordset("SELECT  * FROM TEMP ", dbOpenDynaset)

With sourcetable
sourcetable.MoveFirst
Do
    temptable.AddNew
    temptable.str11 = "«·«⁄÷«¡ «·–Ì‰ ·„ Ì„— ⁄·ÌÂ„ ”‰…" & Format(xDate.Text, "yyyy/m/d")
    temptable.str2 = !Desca
    temptable.Str3 = !DESCE
    temptable!val1 = !Member
    temptable!val2 = 1
    temptable!Val3 = 0
    temptable!date1 = ![file1_10.datebegin]
    temptable!DATE2 = ![file1_11.datebegin]
    temptable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
End With
Report1.WindowShowExportBtn = False
Report1.ReportFileName = MainPath & "\RPT\rep1_4.rpt"
Report1.DataFiles(0) = tempFile
Report1.Action = 1: tempdb.Execute "Delete * from temp"
End Sub
Private Sub Report03_6()
cString = "Select Code,dESce,Desca,Title,dateBegin,DateLast From File1_10 "

If xCode1.Text <> "" Then
    cString = cString & turnFound(cString) & " code >= " & xCode1.Text
End If

If xCode2.Text <> "" Then
    cString = cString & turnFound(cString) & " code <= " & xCode2.Text
End If

Set sourcetable = mydb.OpenRecordset(cString & cWhere)
tempdb.Execute "delete * from temp"
Set temptable = tempdb.OpenRecordset("SELECT  * FROM TEMP ", dbOpenDynaset)
If sourcetable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
    Exit Sub
End If
With sourcetable
sourcetable.MoveFirst
Do
    temptable.AddNew
    temptable.str11 = "ﬂ· «·«⁄÷«¡" & Format(xDate.Text, "yyyy/m/d")
    temptable.str2 = !Desca
    temptable.Str3 = !DESCE
    temptable!val1 = !CODE
    temptable!val2 = 0
    temptable!Val3 = !CODE
    temptable!Str3 = "⁄÷Ê"
    temptable!Str20 = Chr(254)
    temptable!date1 = !datebegin
    temptable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
End With

cString = "Select Member,fILE1_11.dESCe,File1_11.Desca,File1_11.Title,File1_11.dateBegin,File1_11.datebirth,dateLast,isSave From File1_10 Inner Join File1_11 On File1_10.Code = File1_11.Member "
cString = cString & " where Relation = '1'"
       
If xCode1.Text <> "" Then
    cString = cString & turnFound(cString) & " Member >= " & xCode1.Text
End If

If xCode2.Text <> "" Then
    cString = cString & turnFound(cString) & " Member <= " & xCode2.Text
End If

Set sourcetable = mydb.OpenRecordset(cString)
Set temptable = tempdb.OpenRecordset("SELECT  * FROM TEMP ", dbOpenDynaset)

With sourcetable
sourcetable.MoveFirst
Do
    temptable.AddNew
    temptable.str11 = "ﬂ· «·«⁄÷«¡" & Format(xDate.Text, "yyyy/m/d")
    temptable.str2 = !Desca
    temptable.Str3 = !DESCE
    temptable!Str3 = " «»⁄"
    temptable!val1 = !Member
    temptable!val2 = 1
    temptable!Val3 = 0
    temptable!date1 = !datebegin
    temptable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
End With
Report1.WindowShowExportBtn = False
Report1.ReportFileName = MainPath & "\RPT\rep1_3.rpt"
Report1.DataFiles(0) = tempFile
End Sub
Private Sub Form_Unload(Cancel As Integer)
addSetting "header", xHeader.Text, TempSave(Me, , True)
addSetting "date", xDate.Text, TempSave(Me, , True)
addSetting "year", xYear.Text, TempSave(Me, , True)
On Error Resume Next
con.Close
Set con = Nothing
Err.Clear
End Sub

