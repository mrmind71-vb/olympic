VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form ChqRep1 
   Caption         =   " ﬁ«—Ì— «·‘Ìﬂ« "
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1440
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   330
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   75
      Width           =   1155
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2385
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   75
      Width           =   1230
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   3480
      Top             =   480
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
      Height          =   735
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   360
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï :"
      Height          =   225
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   390
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
      Height          =   195
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "ChqRep1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChqTable As Recordset
Dim TempTable As Recordset
Dim ClientTable As Recordset
Dim lExit As Boolean
Function MYVALID()
If Not IsDate(xdate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub CmdGroup_Click()
End Sub
Private Sub Form_Load()
Set TempTable = tempdb.OpenRecordset("TEMP")

End Sub
Private Sub CmdApply_Click()
lExit = False
If Not MYVALID Then Exit Sub
tempdb.Execute "Delete * From Temp"
Select Case publicFlag
    Case 1
        cClosed = "0"
        If lCust Then
            cHead = " √Ê—«ﬁ ﬁ»÷ „” Õﬁ… "
            cString = "select file5_20.* " & _
                    "FROM file5_20  " & _
                    " WHERE CLOSED = '0' and date_1 >= DateValue(" & MyParn(xdate1.Text) & ") And date_1 <= DateValue(" & MyParn(xDate2.Text) & ")"
                    cString = cString & " ORDER BY DATE_1"
            Set ChqTable = mydb.CreateSnapshot(cString)
        Else
            cHead = " √Ê—«ﬁ œ›⁄ „” Õﬁ… "
            cString = "select file5_21.* " & _
                    " FROM file5_21  " & _
                    " WHERE CLOSED = '0' and date_1 >= DateValue(" & MyParn(xdate1.Text) & ") And date_1 <= DateValue(" & MyParn(xDate2.Text) & ")"
            cString = cString & " ORDER BY DATE_1"
            Set ChqTable = mydb.CreateSnapshot(cString)
        End If
    Case 3
        cClosed = "1"
        If lCust Then
            cHead = " ·√Ê—«ﬁ ﬁ»÷ „— œ…  "
            cString = "select file5_20.* " & _
                    "from file5_20 " & _
                    " WHERE CLOSED = '1' and date_3 >= DateValue(" & MyParn(xdate1.Text) & ") And date_3 <= DateValue(" & MyParn(xDate2.Text) & ")"
            cString = cString & " ORDER BY DATE_3"
            Set ChqTable = mydb.CreateSnapshot(cString)
        Else
            cHead = " √Ê—«ﬁ œ›⁄ „— œ… "
            cString = "select file5_21.* " & _
                      "from file5_21 " & _
                    " WHERE CLOSED = '1' and date_3 >= DateValue(" & MyParn(xdate1.Text) & ") And date_3 <= DateValue(" & MyParn(xDate2.Text) & ")"
            cString = cString & " ORDER BY DATE_3"
            Set ChqTable = mydb.CreateSnapshot(cString)
        End If
    Case 2
        cClose = "2"
        If lCust Then
            cHead = " √Ê—«ﬁ ﬁ»÷ „Õ’·… "
            cString = "select file5_20.*  " & _
                    "from file5_20 " & _
                    " WHERE CLOSED = '2' and date_3 >= DateValue(" & MyParn(xdate1.Text) & ") And date_3 <= DateValue(" & MyParn(xDate2.Text) & ")"
            cString = cString & " ORDER BY DATE_3"
            Set ChqTable = mydb.CreateSnapshot(cString)
        Else
            cHead = " √Ê—«ﬁ œ›⁄ „”œœ… "
            cString = "select file5_21.* " & _
                      "from file5_21 " & _
                    " WHERE CLOSED = '2' and date_3 >= DateValue(" & MyParn(xdate1.Text) & ") And date_3 <= DateValue(" & MyParn(xDate2.Text) & ")"
            cString = cString & " ORDER BY DATE_3"
            Set ChqTable = mydb.CreateSnapshot(cString)
        End If
    
    Case 4
        If lCust Then
            cHead = " √Ê—«ﬁ ﬁ»÷ „Õ——…  "
            cString = "select file5_20.* " & _
                    "from file5_20 " & _
                    " WHERE date_R >= DateValue(" & MyParn(xdate1.Text) & ") And date_R <= DateValue(" & MyParn(xDate2.Text) & ")"
            cString = cString & " ORDER BY DATE_R"
        Else
            cHead = " √Ê—«ﬁ œ›⁄ „Õ——…  "
            cString = "select file5_21.*  " & _
                    "from file5_21 " & _
                    " WHERE date_R >= DateValue(" & MyParn(xdate1.Text) & ") And date_R <= DateValue(" & MyParn(xDate2.Text) & ")"
            cString = cString & " ORDER BY DATE_R"
        End If
        Set ChqTable = mydb.CreateSnapshot(cString)
End Select
             
With ChqTable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .CHK_ID
    TempTable.str2 = .code
    TempTable.str3 = .Name1
    TempTable.val1 = .Value
    If publicFlag = 1 Then
        TempTable.Date1 = .date_1
        TempTable.date2 = .date_R
    Else
        TempTable.Date1 = .date_3
        TempTable.date2 = .date_1
        TempTable.date3 = .date_R
    End If
    TempTable.STR6 = "  "
    TempTable.STR7 = cHead & " „‰  «—ÌŒ " & xdate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
    TempTable.str8 = .Memo
    TempTable.str19 = firsttitle
    TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
Else
    MsgBox "·«  ÊÃœ »Ì«‰«  ·· ﬁ—Ì—"
    lExit = True
End If
End With
If Not lExit Then
    Select Case publicFlag
        Case 1
            REPORT1.ReportFileName = App.Path & "\Reports\CHQREP1.rpt"
        Case 3
            REPORT1.ReportFileName = PublicPath & "\Reports\CHQREP3.rpt"
        Case 2
            REPORT1.ReportFileName = PublicPath & "\Reports\CHQREP2.rpt"
        Case 4
            REPORT1.ReportFileName = PublicPath & "\Reports\CHQREP6.rpt"
    End Select
    REPORT1.DataFiles(0) = App.Path & "\temp.mdb"
    REPORT1.Action = 1
End If
End Sub
Function myDateValue(pDate1)
myDateValue = "DateValue('" & pDate1 & "')"
End Function
