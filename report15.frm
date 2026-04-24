VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form reportfrm15 
   Caption         =   "»Ì«‰«   «· Ê«»⁄"
   ClientHeight    =   4365
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   7425
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
   ScaleHeight     =   4365
   ScaleWidth      =   7425
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
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   3735
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
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   3735
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
      TabIndex        =   16
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   3735
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2250
      Width           =   2400
      Begin VB.CheckBox xDied 
         Appearance      =   0  'Flat
         Caption         =   " ⁄—÷ «·„ Ê›Ì‰ ›ﬁÿ"
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   945
         Width           =   2130
      End
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   540
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   2040
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3660
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   45
      Width           =   4785
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
      Begin VB.TextBox xNotes 
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
         TabIndex        =   12
         Tag             =   "D"
         Top             =   3150
         Width           =   3075
      End
      Begin VB.TextBox xAge2 
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
         TabIndex        =   10
         Tag             =   "D"
         Top             =   2385
         Width           =   1680
      End
      Begin VB.TextBox xAge1 
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
         TabIndex        =   9
         Tag             =   "D"
         Top             =   2385
         Width           =   1410
      End
      Begin VB.TextBox xdate_begin2 
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
         TabIndex        =   8
         Tag             =   "D"
         Top             =   2025
         Width           =   1680
      End
      Begin VB.TextBox xDate_begin1 
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
         TabIndex        =   7
         Tag             =   "D"
         Top             =   2025
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
         TabIndex        =   5
         Tag             =   "D"
         Top             =   1665
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
         TabIndex        =   6
         Tag             =   "D"
         Top             =   1665
         Width           =   1680
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   2745
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
      Begin MSDataListLib.DataCombo xRelation 
         Height          =   330
         Left            =   135
         TabIndex        =   3
         Top             =   900
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
      Begin MSDataListLib.DataCombo xGender 
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Top             =   1260
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
         Index           =   8
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   225
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "„·«ÕŸ« "
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
         TabIndex        =   27
         Top             =   3195
         Width           =   1005
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
         Index           =   6
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1350
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "œ—Ã… «·ﬁ—«»…"
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
         Index           =   3
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "«·”‰"
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
         TabIndex        =   24
         Top             =   2475
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ «· Õ«ﬁ"
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
         Index           =   0
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2115
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "«·›∆…"
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
         Index           =   5
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   630
         Width           =   1230
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
         TabIndex        =   21
         Top             =   1755
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
         TabIndex        =   20
         Top             =   2790
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
      Left            =   0
      Top             =   0
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
      Top             =   0
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
      Top             =   0
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
Attribute VB_Name = "reportfrm15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty, oSearchJob As New Search, oSearchComp As New Search
Private Sub cmdApply_Click()
doPrint
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Function doPrint()
Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset, cOr As String
Dim aHeader(11)
cString = "Select FILE1_10.CODE,file1_10.DESCA,TYPE_CODES.DescA as TYPE_Desca,FILE1_11.DESCA AS REL_DESCA,FILE1_11.DATE_BEGIN,FILE1_11.DATE_BIRTH,dbo.f_last_year_desca(file1_10.code) as year_desca,dbo.f_age(file1_11.date_birth," & DateSq(Date, True) & ") AS AGE,RELATION_CODES.DESCA AS RELATION_DESCA " & _
          " From File1_10 INNER JOIN FILE1_11 ON FILE1_10.CODE = FILE1_11.MEMBER INNER join TYPE_CODES on File1_10.TYPE = TYPE_CODES.CODE INNER JOIN RELATION_CODES ON FILE1_11.RELATION = RELATION_CODES.CODE"


If xRelation.MatchedWithList Then
    aHeader(0) = "œ—Ã… «·ﬁ—«»… : " & xRelation.text
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_11.relation = " & addvalue(xRelation.BoundText)
End If

If ValidNum(xcode1.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE1_10.code " & IIf(ValidNum(xCode2.text), " >= ", " = ") & addvalue(xcode1.text)
    aHeader(1) = IIf(ValidNum(xCode2.text), BetweenString(xcode1.text, xCode2.text, "„‰ —ﬁ„ ⁄÷ÊÌ… : ", "Õ Ì —ﬁ„ ⁄÷ÊÌ… : "), "—ﬁ„ ⁄÷ÊÌ… :" & xcode1.text)
End If

If ValidNum(xCode2.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE1_10.code <= " & addvalue(xCode2.text)
    aHeader(1) = BetweenString(xcode1.text, xCode2.text, "„‰ —ﬁ„ ⁄÷ÊÌ… : ", "Õ Ì —ﬁ„ ⁄÷ÊÌ… : ")
End If

If xType.MatchedWithList Then
    aHeader(2) = "«·›∆… : " & xType.text
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.TYPE = " & addvalue(xType.BoundText)
End If

If xGender.MatchedWithList Then
    aHeader(3) = "«·‰Ê⁄ : " & xGender.text
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.Gender = " & addvalue(xGender.BoundText)
End If

If IsNumeric(cmdYear(0).Tag) Then
     aHeader(4) = "«·–Ì‰ ”œœÊ« „Ê”„ " & cmdYear(0).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) >= " & cmdYear(0).Tag
End If


If IsDate(xDate_begin1.text) Then
    aHeader(5) = " «—ÌŒ «· Õ«ﬁ  " & BetweenString(xDate_begin1.text, xdate_begin2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_11.DATE_BEGIN >= " & DateSq(xDate_begin1.text)
End If

If IsDate(xdate_begin2.text) Then
    aHeader(5) = " «—ÌŒ «· Õ«ﬁ  " & BetweenString(xDate_begin1.text, xdate_begin2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_11.DATE_BEGIN <= " & DateSq(xdate_begin2.text)
End If


If IsDate(xDate1.text) Then
    aHeader(6) = "„”œœ " & BetweenString(xDate1.text, xDate2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE WHERE DATE >= " & DateSq(xDate1.text) & " AND (NOT FORM_NO IS NULL) AND PAID_TYPES.IS_PAID = 1 )"
End If

If IsDate(xDate2.text) Then
    aHeader(6) = "„”œœ " & BetweenString(xDate1.text, xDate2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H WHERE DATE <= " & DateSq(xDate2.text) & " AND (NOT FORM_NO IS NULL))"
End If

If xSafe.Value = 0 Then
    cWhere = cWhere & turn(cWhere, " and ") & " (dbo.f_save(FILE1_10.CODE) = 0)"
    aHeader(7) = "»œÊ‰ Õ«›ŸÌ «·⁄÷ÊÌ…"
Else
    'aHeader(5) = "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ…"
End If

If xDied.Value = 1 Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.DIED = 1"
    aHeader(7) = "«·„ Ê›ÌÌ‰"
Else
    'aHeader(5) = "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ…"
End If


If ValidNum(xAge1.text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " dbo.f_age(file1_10.date_birth," & DateSq(Date) & ") >= " & Val(xAge1.text)
    aHeader(8) = BetweenString(xAge1.text, xAge2.text, "„‰ ”‰", "Õ Ì ”‰")
End If

If ValidNum(xAge2.text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " dbo.f_age(file1_10.date_birth," & DateSq(Date) & ") <= " & Val(xAge2.text)
    aHeader(8) = BetweenString(xAge1.text, xAge2.text, "„‰ ”‰", "Õ Ì ”‰")
End If


If Trim(xNotes.text) <> "" Then
    aHeader(9) = "«·„·ÕÊŸ… : " & xNotes.text
    cWhere = cWhere & turn(cWhere, " and ") & MyParnAnd(xNotes, "file1_11.notes")
End If

If xDrop.Value = 0 Then
    cWhere2 = cWhere & turn(cWhere, " and ") & " (FILE1_10.[DROP] = 0)"
   ' aHeader(6) = "»œÊ‰ ”«ﬁÿÌ «·⁄÷ÊÌ…"
Else
    aHeader(10) = "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ…"
End If

'If xDied.Value = 0 Then
'   cWhere2 = cWhere & turn(cWhere, " and ") & " (Died = 0)"
'  '  aHeader(10) = "»œÊ‰ «·„ Ê›ÌÌ‰"
'Else
'    aHeader(5) = "⁄—÷ «·„ Ê›ÌÌ‰"
'End If

If cWhere2 <> "" Then
    cString = cString & " AND " & cWhere2
ElseIf cWhere <> "" Then
    cString = cString & " AND " & cWhere
End If

Me.MousePointer = 11
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!val1 = sourcetable!code
    temptable!str1 = ArbString(sourcetable!code)
    temptable!str2 = sourcetable!rel_Desca
    temptable!str3 = sourcetable!Desca
    temptable!str4 = TurnValue(ArbString(myFormat_p(sourcetable!date_begin)))
    temptable!str5 = TurnValue(sourcetable!TYPE_desca)
    temptable!str6 = TurnValue(ArbString(sourcetable!year_desca & ""))
    temptable!str7 = sourcetable!RELATION_DESCA
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
    
    Report1.ReportFileName = sPath_App & "\REPORTS\REPORT15.rpt"
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

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Form_Load()
openCon con

Set data1.Recordset = myRecordSet("SELECT CODE,DESCA FROM TYPE_CODES ORDER BY CODE", con)
Set xType.RowSource = data1
xType.ListField = "Desca"
xType.BoundColumn = "Code"

Set data2.Recordset = myRecordSet("SELECT CODE,DESCA FROM RELATION_CODES ORDER BY CODE", con)
Set xRelation.RowSource = data2
xRelation.ListField = "Desca"
xRelation.BoundColumn = "Code"

Set data3.Recordset = myRecordSet("select * from Gender_Codes", con)
Set xGender.RowSource = data3
xGender.ListField = "Desca"
xGender.BoundColumn = "Code"

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
Private Sub xGender_GotFocus()
myGotFocus xGender
End Sub
Private Sub xGender_LostFocus()
myLostFocus xGender
If Not xGender.MatchedWithList Then xGender.BoundText = ""
End Sub
Sub myProc()
If ActiveControl.Name = cmdYear(0).Name Then
    ActiveControl.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
    ActiveControl.Caption = IIf(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0) = "", "«Œ «— «·„Ê”„", oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
    oSearchYear.Hide
ElseIf ActiveControl.Name = xJob.Name Then
    xJob.BoundText = oSearchJob.grid1.TextMatrix(oSearchJob.grid1.Row, 0)
ElseIf ActiveControl.Name = xCompany.Name Then
    xCompany.BoundText = oSearchComp.grid1.TextMatrix(oSearchComp.grid1.Row, 0)
End If
End Sub

Private Sub xJob_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Job_Lookup Me, oSearchJob
End If
End Sub
Private Sub xAge2_GotFocus()
myGotFocus xAge2
End Sub
Private Sub xAge2_LostFocus()
myLostFocus xAge2
End Sub
Private Sub xAge1_GotFocus()
myGotFocus xAge1
End Sub
Private Sub xAge1_LostFocus()
myLostFocus xAge1
End Sub
Private Sub xdate_begin2_GotFocus()
myGotFocus xdate_begin2
End Sub
Private Sub xdate_begin2_LostFocus()
myLostFocus xdate_begin2
myValidDate xdate_begin2
End Sub
Private Sub xDate_begin1_GotFocus()
myGotFocus xDate_begin1
End Sub
Private Sub xDate_begin1_LostFocus()
myLostFocus xDate_begin1
myValidDate xDate_begin1
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
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

