VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpBank3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ›’Ì·Ì Õ—ﬂ… »‰ﬂÌ…"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   6435
   Begin VB.CommandButton cmdExit 
      Height          =   555
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   3060
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
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   3060
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
      TabIndex        =   15
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   3060
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   6270
      Begin VB.TextBox xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   3345
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1035
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xBank 
         Height          =   390
         Left            =   1035
         TabIndex        =   5
         Top             =   180
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XCODE 
         Height          =   390
         Left            =   1845
         TabIndex        =   9
         Top             =   1845
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "«·»‰œ"
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
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1935
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·»Ì«‰"
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
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1530
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·»‰ﬂ"
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
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   270
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·Ï  «—ÌŒ"
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
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1125
         Width           =   690
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   5805
      Top             =   450
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
      Width           =   2340
      _ExtentX        =   4128
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
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2430
      Width           =   6270
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„”ÕÊ»«  ›ﬁÿ"
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
         Height          =   195
         Index           =   2
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   1770
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«Ìœ«⁄«  ›ﬁÿ"
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
         Height          =   195
         Index           =   1
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·ﬂ·"
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
         Height          =   195
         Index           =   0
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   270
         Value           =   -1  'True
         Width           =   915
      End
   End
End
Attribute VB_Name = "rpBank3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim ChargeTable As ADODB.Recordset
Private Sub cmdApply_Click()
    doprint1
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
FixRpImage Me
openCon con
Set ChargeTable = New ADODB.Recordset
ChargeTable.Open "File5_00", con, adOpenStatic, adLockReadOnly, adCmdTable
xDate1.Text = ""
xdate2.Text = ""

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE5_10 ORDER BY DESCA"
Set xBank.RowSource = data1
xBank.ListField = "Desca"
xBank.BoundColumn = "code"

data2.ConnectionString = strCon
data2.RecordSource = "SELECT * FROM FILE5_00 ORDER BY DESCA"

Set XCODE.RowSource = data2
XCODE.ListField = "Desca"
XCODE.BoundColumn = "Code"

End Sub
Function MYVALID() As Boolean
If Not IsDate(xDate1.Text) And (xDate1.Text <> "") Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If Not IsDate(xdate2.Text) And (xdate2.Text <> "") Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(4)
If Not MYVALID Then Exit Sub
Dim i As Integer, nPrevious As Double

Dim sourcetable As ADODB.Recordset
Dim temptable As ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Set sourcetable = New ADODB.Recordset
cString = "Select FILE5_30H.BANK,FILE5_30.*,FILE5_00.DESCA AS FILE5_00DESCA,FILE5_10.DESCA AS FILE5_10DESCA FROM FILE5_30 INNER JOIN FILE5_30H ON FILE5_30.DOC_NO = FILE5_30H.DOC_NO INNER JOIN FILE5_00 ON FILE5_30.CODE = FILE5_00.CODE LEFT JOIN FILE5_10 ON FILE5_30H.BANK = FILE5_10.code "
If Option1(1).Value Then
    cString = cString & turnFound(cString) & " Value1  <> 0"
    aHeader(4) = "«·«Ìœ«⁄«  ›ﬁÿ"
End If

If Option1(2).Value Then
    cString = cString & turnFound(cString) & " Value2 <> 0"
    aHeader(4) = "«·„”ÕÊ»«  ›ﬁÿ"
End If


If Trim(xBank.BoundText) <> "" Then
    cString = cString & turnFound(cString) & " BANK = " & MyParn(xBank.BoundText)
    aHeader(0) = "[" & "«·»‰ﬂ : " & xBank.Text & "]"
End If

If Trim(XCODE.BoundText) <> "" Then
    cString = cString & turnFound(cString) & " file5_30.code = " & MyParn(XCODE.BoundText)
    aHeader(1) = "[" & "«·»‰œ : " & XCODE.Text & "]"
End If


If IsDate(xDate1.Text) Then
    cString = cString & turnFound(cString) & "Date >= " & DateSq(xDate1.Text)
    aHeader(2) = "[" & BetweenString(xDate1.Text, xdate2.Text) & "]"
End If
          
If IsDate(xdate2.Text) Then
    cString = cString & turnFound(cString) & "Date <= " & DateSq(xdate2.Text)
    aHeader(2) = "[" & BetweenString(xDate1.Text, xdate2.Text) & "]"
End If
                   
If Trim(xdesca.Text) <> "" Then
    cString = cString & turnFound(cString) & " file5_30.desca like " & MyParnAll(xdesca.Text)
    aHeader(3) = "[" & "«·»Ì«‰ : " & xdesca.Text & "]"
End If
                   
cString = cString & " Order by [Date],Doc_no,value2"

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
    Do Until sourcetable.EOF
        i = i + 1
        temptable.AddNew
        temptable!str1 = !doc_no
        temptable!str2 = !FILE5_00DESCA
        temptable!str3 = !Desca
        temptable!str4 = !FILE5_10desca
        temptable!Date1 = !Date
        temptable!val1 = Val(!value1 & "")
        temptable!val2 = Val(!Value2 & "")
        temptable!str21 = TurnValue(retHeader(aHeader, 0, 5))
        temptable!val10 = i
        temptable.Update
        sourcetable.MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.Report1.ReportFileName = App.Path & "\Reports\BANK3.rpt"
    main.Report1.DataFiles(0) = "c:\tempmrshd\Temp.MDB"
    main.Report1.Action = 1
End If
If temptable.State = adStateOpen Then temptable.Close
If sourcetable.State = adStateOpen Then sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

