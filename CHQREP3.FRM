VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form ChqRep3 
   Caption         =   " ﬁ«—Ì— «·‘Ìﬂ« "
   ClientHeight    =   1740
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
   ScaleHeight     =   1740
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   150
      Width           =   1230
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1155
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1230
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   3450
      Top             =   1050
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
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   900
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ﬂÊœ"
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
      Left            =   4140
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   150
      Width           =   270
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ï :"
      Height          =   195
      Left            =   1590
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   360
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
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   765
   End
End
Attribute VB_Name = "ChqRep3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChqTable As Recordset
Dim TempTable As Recordset
Dim ClientTable As Recordset
Function MYVALID()
If Not IsDate(xdate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
If xCode.Text = "" Then Exit Function
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xGroup1.BoundText = ""
xGroup2.BoundText = ""
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
nOption = 1
cHead = cHead & "Ê –·ﬂ „‰  «—ÌŒ " & xdate1.Text & "≈·Ï  «—ÌŒ " & xDate2.Text
'Set ChqTable = mydb.CreateSnapshot(cString)

If publicFlag = 1 Then
    If lCust Then
        cString = "select file5_20.*  " & _
                  "from file5_20 WHERE CLOSED = '0' and date_1 >= " & myDateValue(xdate1.Text) & " And date_1 <= " & myDateValue(xDate2.Text) & _
                  " and File5_20.CODE = " & MyParn(xCode.Text) & " "
    Else
        cString = "select file5_21.*  " & _
                  "from file5_21 WHERE CLOSED = '0' and date_1 >= " & myDateValue(xdate1.Text) & " And date_1 <= " & myDateValue(xDate2.Text) & _
                  " and File5_21.CODE = " & MyParn(xCode.Text) & " "
    End If
Else
    If lCust Then
        cString = "select file5_20.*  " & _
                  "from file5_20 WHERE CLOSED = '1' and date_1 >= " & myDateValue(xdate1.Text) & " And date_1 <= " & myDateValue(xDate2.Text) & _
                  " and File5_20.CODE = " & MyParn(xCode.Text) & " "
    Else
        cString = "select file5_21.*  " & _
                  "from file5_21 WHERE CLOSED = '1' and date_1 >= " & myDateValue(xdate1.Text) & " And date_1 <= " & myDateValue(xDate2.Text) & _
                  " and File5_21.CODE = " & MyParn(xCode.Text) & " "
    End If
End If
Set ChqTable = mydb.CreateSnapshot(cString)
With ChqTable
If .RecordCount > 0 Then
    .MoveFirst
    If lCust Then
        If publicFlag = 1 Then
            cHead = " √Ê—«ﬁ ﬁ»÷ „” Õﬁ… ··⁄„Ì·   " & .Name1
        Else
            cHead = " √Ê—«ﬁ ﬁ»÷ „— œ… ⁄·Ï «·⁄„Ì·  " & .Name1
        End If
    Else
        If publicFlag = 1 Then
            cHead = " √Ê—«ﬁ œ›⁄ „” Õﬁ… ··„Ê—œ  " & .Name1
        Else
            cHead = " √Ê—«ﬁ œ›⁄ „— œ… ··„Ê—œ  " & .Name1
        End If
    End If
    Do While Not .EOF
        TempTable.AddNew
        TempTable.str1 = .CHK_ID
        TempTable.val1 = .Value
        If publicFlag = 1 Then
            TempTable.Date1 = .date_1
            TempTable.date2 = .date_R
        Else
            TempTable.Date1 = .date_3
            TempTable.date2 = .date_1
        End If
        TempTable.STR7 = cHead
        TempTable.str8 = " „‰  «—ÌŒ " & xdate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
        TempTable.str19 = firsttitle
        TempTable.str20 = Secondtitle

        TempTable.Update
        .MoveNext
    Loop
    Report1.ReportFileName = PublicPath & "\Reports\CHQREP4.rpt"
    Report1.DataFiles(0) = App.Path & "\temp.mdb"
    Report1.Action = 1
Else
    lExit = True
    MsgBox "·«  ÊÃœ »Ì«‰«  ·· ﬁ—Ì—"
End If
End With
End Sub
Function myDateValue(pDate1)
myDateValue = "DateValue('" & pDate1 & "')"
End Function
Private Sub xCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xCode.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(2)
    If lCust Then
        Set Generalarray(1) = Me
        Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As ≈”„ From File3_10"
        Generalarray(3) = "Where DescA Like '*cFilter*'"
            
        GrdArray(1) = 1000
        GrdArray(2) = 2600
               
        Lookupdata = Array(Generalarray, GrdArray)
        Load Search
        Search.Caption = "«” ⁄·«„ "
        Search.Show 1
    Else
        Set Generalarray(1) = Me
        Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As ≈”„ From File4_10"
        Generalarray(3) = "Where DescA Like '*cFilter*'"
            
        GrdArray(1) = 1000
        GrdArray(2) = 2600
               
        Lookupdata = Array(Generalarray, GrdArray)
        Load Search
        Search.Caption = "«” ⁄·«„ "
        Search.Show 1
    End If
End If
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub

