VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form reportfrm31 
   Caption         =   "«⁄œ«œ «·«⁄÷«¡"
   ClientHeight    =   2685
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   10080
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
   ScaleHeight     =   2685
   ScaleWidth      =   10080
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
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   2070
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
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   2070
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
      TabIndex        =   11
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   2070
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   810
      Width           =   2760
      Begin VB.CheckBox xSafeOnly 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ… ›ﬁÿ"
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   810
         Width           =   2400
      End
      Begin VB.CheckBox xDiedOnly 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ «·„ Ê›ÌÌ‰ ›ﬁÿ"
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   135
         Width           =   1950
      End
      Begin VB.CheckBox xDropOnly 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ… ›ﬁÿ"
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   450
         Width           =   2400
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1230
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   810
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   135
         Value           =   1  'Checked
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   5355
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   135
      Width           =   4515
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
         TabIndex        =   0
         Tag             =   "D"
         Top             =   225
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
         TabIndex        =   1
         Tag             =   "D"
         Top             =   225
         Width           =   1680
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   0
         Left            =   135
         TabIndex        =   2
         Top             =   585
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
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   1
         Left            =   135
         TabIndex        =   3
         Top             =   990
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
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   1395
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
         Caption         =   "·„ Ì”œœ"
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
         Index           =   3
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "«Œ— ”œ«œ ›Ì"
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
         Index           =   2
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1035
         Width           =   1095
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
         TabIndex        =   16
         Top             =   270
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
         TabIndex        =   15
         Top             =   630
         Width           =   960
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3510
      Top             =   225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1215
      Top             =   315
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
   Begin MSDataListLib.DataCombo xType 
      Height          =   330
      Left            =   0
      TabIndex        =   21
      Top             =   270
      Visible         =   0   'False
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
      Left            =   3195
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   315
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "reportfrm31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty
Private Sub cmdApply_Click()
doPrint
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Function doPrint()
Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset
Dim nVal1 As Long, nVal2 As Long, nVal3 As Long, nVal4 As Long
Dim nVal5 As Long, nVal6 As Long, nVal7 As Long, nVal8 As Long

Dim aHeader(11)
cField = myiif("COALESCE(FILE1_10.GENDER,1) = 1", "1") & " AS MEM_MALE,"

cField = cField & _
        myiif("COALESCE(FILE1_10.GENDER,1) = 2", "1") & " AS MEM_FEMALE"

cString = "Select " & cField & _
          " From File1_10 "

If IsNumeric(cmdYear(0).Tag) Then
    aHeader(0) = "«·–Ì‰ ”œœÊ« „Ê”„ " & cmdYear(0).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) >= " & cmdYear(0).Tag
End If

If IsNumeric(cmdYear(1).Tag) Then
    aHeader(1) = "«Œ— ”œ«œ ·Â„ ›Ì „Ê”„ " & cmdYear(1).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) = " & cmdYear(1).Tag
End If

If IsNumeric(cmdYear(2).Tag) Then
    aHeader(2) = "·„ Ì”œœÊ« „Ê”„ " & cmdYear(2).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) < " & cmdYear(2).Tag
End If


If xType.MatchedWithList Then
    aHeader(4) = "›∆… «·⁄÷ÊÌ… : " & xType.text
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.TYPE = " & addvalue(xType.BoundText)
End If

If IsDate(xDate1.text) Then
    aHeader(5) = "„”œœ " & BetweenString(xDate1.text, xDate2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE WHERE DATE >= " & DateSq(xDate1.text) & " AND (NOT FORM_NO IS NULL) AND PAID_TYPES.IS_PAID = 1 )"
End If

If IsDate(xDate2.text) Then
    aHeader(5) = "„”œœ " & BetweenString(xDate1.text, xDate2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H WHERE DATE <= " & DateSq(xDate2.text) & " AND (NOT FORM_NO IS NULL))"
End If

If xSafeOnly.Value = 0 Then
    If xSafe.Value = 0 Then
        cWhere = cWhere & turn(cWhere, " and ") & " (dbo.f_save(FILE1_10.CODE) = 0)"
        'aHeader(5) = "»œÊ‰ Õ«›ŸÌ «·⁄÷ÊÌ…"
    Else
        aHeader(6) = "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ…"
    End If
End If

If xDropOnly.Value = 0 Then
    If xDrop.Value = 0 Then
        cWhere = cWhere & turn(cWhere, " and ") & " (FILE1_10.[DROP] = 0)"
        'aHeader(6) = "»œÊ‰ ”«ﬁÿÌ «·⁄÷ÊÌ…"
    Else
        aHeader(7) = "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ…"
    End If
End If

If xDropOnly.Value = 1 Then
    aHeader(8) = "«·”«ﬁÿÌ «·⁄÷ÊÌ… ›ﬁÿ"
    cWhere = cWhere & turn(cWhere, " and ") & "[DROP] = 1"
End If

If xSafeOnly.Value = 1 Then
    aHeader(9) = "Õ«›ŸÌ «·⁄÷ÊÌ… ›ﬁÿ"
    cWhere = cWhere & turn(cWhere, " and ") & " (dbo.f_save(FILE1_10.CODE) = 1)"
End If

If xDiedOnly.Value = 1 Then
    aHeader(10) = "«·„ Ê›ÌÌ‰ ›ﬁÿ"
    cWhere = cWhere & turn(cWhere, " and ") & "Died = 1"
ElseIf xDiedOnly.Value = 0 Then
    If xDied.Value = 0 Then
       cWhere2 = cWhere & turn(cWhere, " and ") & " (Died = 0)"
        aHeader(10) = "»œÊ‰ «·„ Ê›ÌÌ‰"
    Else
        aHeader(10) = "⁄—÷ «·„ Ê›ÌÌ‰"
    End If
End If

If cWhere2 <> "" Then
    cString = cString & " WHERE " & cWhere2
ElseIf cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

'cString = cString & " Group By File1_10.[type],type_codes.Desca"

Me.MousePointer = 11
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

' «·«⁄÷«¡ «·⁄«„·Ì‰
With sourcetable
If Not sourcetable.EOF Then
    nVal1 = mRound(sourcetable!MEM_MALE)
    nVal2 = mRound(sourcetable!MEM_FEMALE)
End If
    
    
cField = myiif("COALESCE(FILE1_11.GENDER,1) = 1", "1") & " AS MEM_REL_MALE,"
cField = cField & _
        myiif("COALESCE(FILE1_11.GENDER,1) = 2", "1") & " AS MEM_REL_FEMALE"

cString = "Select " & _
          cField & _
          " From File1_10 inner join file1_11 on file1_10.code = file1_11.Member"
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

sourcetable.Close
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext
'  «·«⁄÷«¡ «· «»⁄Ì‰ ··⁄÷Ê«·⁄«„·
If Not sourcetable.EOF Then
    nVal3 = mRound(sourcetable!MEM_REL_MALE)
    nVal4 = mRound(sourcetable!MEM_REL_FEMALE)
End If
 

cField = myiif("COALESCE(FILE2_10.GENDER,1) = 1 OR COALESCE(FILE2_10.GENDER,1) = 3", "1") & " AS MEM_MALE,"

cField = cField & _
        myiif("COALESCE(FILE2_10.GENDER,1) = 2", "1") & " AS MEM_FEMALE"

' «⁄÷«¡  ﬁ”Ìÿ
Dim aRet As Variant
aRet = GetFields("SELECT " & cField & " FROM FILE2_10 WHERE STATUS <= 2", con)
If Not IsEmpty(aRet) Then
    nVal5 = mRound(retFlag(aRet, "MEM_MALE"))
    nVal6 = mRound(retFlag(aRet, "MEM_FEMALE"))
End If


cField = myiif("COALESCE(FILE2_11.GENDER,1) = 1 OR COALESCE(FILE2_11.GENDER,1) = 3 ", "1") & " AS MEM_REL_MALE,"

cField = cField & _
        myiif("COALESCE(FILE2_11.GENDER,1) = 2", "1") & " AS MEM_REL_FEMALE"

cString = "Select " & _
          cField & _
          " From File2_10 inner join file2_11 on file2_10.code = file2_11.Member" & _
          " WHERE FILE2_10.STATUS <= 2"

' «⁄÷«¡  «»⁄Ì‰ ··⁄÷Ê«·„ﬁ”ÿ
aRet = GetFields(cString, con)
If Not IsEmpty(aRet) Then
    nVal7 = mRound(retFlag(aRet, "MEM_REL_MALE"))
    nVal8 = mRound(retFlag(aRet, "MEM_REL_FEMALE"))
End If
    
temptable.AddNew
temptable!val1 = nVal1
temptable!val2 = nVal2
temptable!VAL3 = nVal3
temptable!val4 = nVal4
temptable!VAL5 = nVal5
temptable!VAL6 = nVal6
temptable!VAL7 = nVal7
temptable!val8 = nVal8

temptable!str11 = TurnValue(retHeader(aHeader, 0, 3))
temptable!str12 = TurnValue(retHeader(aHeader, 3, 5))
temptable.Update

contemp.BeginTrans
contemp.CommitTrans

Report1.ReportFileName = sPath_App & "\REPORTS\REPORT31.rpt"
Report1.DataFiles(0) = tempFile
Report1.Action = 1
Me.MousePointer = 0

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

