VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form L_Sal 
   Caption         =   "ﬂ‘› «·„— »"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
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
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   495
      Width           =   1290
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1290
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1725
      Top             =   150
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
      Height          =   690
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   900
      Width           =   4845
      Begin VB.CommandButton Cmd_Print2 
         Caption         =   "⁄—÷ ≈Ã„«·Ï ··√ÃÊ—"
         Height          =   390
         Left            =   1417
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   233
         Width           =   1500
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷  ›’Ì·Ï ··√ÃÊ—"
         Height          =   390
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "≈·Ï  «—ÌŒ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   515
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "L_Sal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim DataTable As Recordset
Dim nOption As Integer
Dim cT_Time As Double
Function MYVALID()
If Not IsDate(xdate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
MYVALID = True
End Function
Private Sub CmdClear_Click()
xdate1.Text = ""
xDate2.Text = ""
End Sub

Private Sub Cmd_Print2_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TempTable = tempdb.OpenRecordset("TEMP")
With TempTable
cString = " SELECT FILE9_10.CODE , FILE9_10.NAMEA , FILE9_10.fix, Sum(FILE9_20.TIME) AS SumTIME, Sum(FILE9_20.DISC) AS SumDISC,  " & _
          " Sum(FILE9_20.Friday) AS SumFriday, Sum(FILE9_20.DEBIT) AS SumDEBIT , Sum(FILE9_20.ADDSAL) AS SumAddSAL, First(FILE9_10.SAL) AS EmpSal  " & _
          " FROM FILE9_10 RIGHT JOIN FILE9_20 ON FILE9_10.CODE = FILE9_20.CODE " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " GROUP BY FILE9_10.NAMEA , FILE9_10.CODE , FILE9_10.fix " & _
          " ORDER BY FILE9_10.CODE "
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
         cT_Time = TurnValue(DataTable.SumTIME, Null, 0) + TurnValue(DataTable.SumFRIDAY, Null, 0)
        .str1 = DataTable.code
        .str2 = DataTable.NAMEA
        If DataTable.Fix Then
            .val1 = 0
            .val2 = 0
            .VAL3 = 0
            .VAL4 = 0
            .VAL5 = 0
            .VAL6 = 0
            .VAL7 = 0
            .VAL8 = 0
            .VAL9 = DataTable.EMPSAL
            .VAL10 = 0
        Else
            .val1 = DataTable.SumTIME
            .val2 = DataTable.SumFRIDAY
            .VAL3 = cT_Time
            .VAL4 = DataTable.EMPSAL
            .VAL5 = (DataTable.EMPSAL / 30) * cT_Time
            .VAL6 = DataTable.SUMDEBIT
            .VAL7 = DataTable.SumDISC * (DataTable.EMPSAL / 30)
            .VAL8 = TurnValue(DataTable.SUMADDSAL, Null, 0)
            .VAL9 = ((DataTable.EMPSAL / 30) * cT_Time) - DataTable.SUMDEBIT - (DataTable.SumDISC * (DataTable.EMPSAL / 30)) + TurnValue(DataTable.SUMADDSAL, Null, 0)
            .VAL10 = 1
        End If
        .STR7 = "ﬂ‘› «·√ÃÊ— „‰ ÌÊ„" & xdate1.Text & " ·ÌÊ„ " & xDate2.Text
        .str9 = firsttitle
        .str10 = Secondtitle
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

If TempTable.RecordCount > 0 Then
    Report1.ReportFileName = PublicPath & "\Reports\L_SAL.rpt"
    Report1.DataFiles(0) = App.Path & "\Temp.mdb"
    Report1.Action = 1
Else
    MsgBox "·«  ÊÃœ √ÃÊ— „” Õﬁ… ·Â–… «·› —… "
End If
End With

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If TypeOf ActiveControl Is DBCombo Then ActiveControl.BoundText = ""
End If
End Sub
Private Sub Form_Load()
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "DELETE * FROM TEMP"
Set TempTable = tempdb.OpenRecordset("TEMP")
With TempTable
cString = " SELECT FILE9_10.CODE , FILE9_10.NAMEA , FILE9_10.fix, Sum(FILE9_20.TIME) AS SumTIME, Sum(FILE9_20.DISC) AS SumDISC,  " & _
          " Sum(FILE9_20.Friday) AS SumFriday, Sum(FILE9_20.DEBIT) AS SumDEBIT , Sum(FILE9_20.ADDSAL) AS SumAddSAL, First(FILE9_10.SAL) AS EmpSal  " & _
          " FROM FILE9_10 RIGHT JOIN FILE9_20 ON FILE9_10.CODE = FILE9_20.CODE " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          " GROUP BY FILE9_10.NAMEA , FILE9_10.CODE , FILE9_10.fix " & _
          " ORDER BY FILE9_10.CODE "
Set DataTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If DataTable.RecordCount > 0 Then
    DataTable.MoveFirst
    Do While True
        .AddNew
         cT_Time = TurnValue(DataTable.SumTIME, Null, 0) + TurnValue(DataTable.SumFRIDAY, Null, 0)
        .str1 = DataTable.code
        .str2 = DataTable.NAMEA
        If DataTable.Fix Then
            .val1 = 0
            .val2 = 0
            .VAL3 = 0
            .VAL4 = 0
            .VAL5 = 0
            .VAL6 = 0
            .VAL7 = 0
            .VAL8 = 0
            .VAL9 = DataTable.EMPSAL
            .VAL10 = 0
        Else
            .val1 = DataTable.SumTIME
            .val2 = DataTable.SumFRIDAY
            .VAL3 = cT_Time
            .VAL4 = DataTable.EMPSAL
            .VAL5 = (DataTable.EMPSAL / 30) * cT_Time
            .VAL6 = DataTable.SUMDEBIT
            .VAL7 = DataTable.SumDISC * (DataTable.EMPSAL / 30)
            .VAL8 = TurnValue(DataTable.SUMADDSAL, Null, 0)
            .VAL9 = ((DataTable.EMPSAL / 30) * cT_Time) - DataTable.SUMDEBIT - (DataTable.SumDISC * (DataTable.EMPSAL / 30)) + TurnValue(DataTable.SUMADDSAL, Null, 0)
            .VAL10 = 1
        End If
        .STR7 = "ﬂ‘› «·√ÃÊ— „‰ ÌÊ„" & xdate1.Text & " ·ÌÊ„ " & xDate2.Text
        .str9 = firsttitle
        .str10 = Secondtitle
        .Update
        DataTable.MoveNext
        If DataTable.EOF Then Exit Do
    Loop
End If

If TempTable.RecordCount > 0 Then
    Report1.ReportFileName = PublicPath & "\Reports\L2_SAL.rpt"
    Report1.DataFiles(0) = App.Path & "\Temp.mdb"
    Report1.Action = 1
Else
    MsgBox "·«  ÊÃœ √ÃÊ— „” Õﬁ… ·Â–… «·› —… "
End If
End With
End Sub
