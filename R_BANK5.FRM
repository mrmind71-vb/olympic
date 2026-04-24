VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Bank5 
   Caption         =   " ﬁ«—Ì— «·≈⁄ „«œ« "
   ClientHeight    =   1995
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
   ScaleHeight     =   1995
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1155
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   1230
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3960
      Top             =   1320
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
      Height          =   645
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   2385
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   1005
      End
   End
   Begin MSDBCtls.DBCombo xBank 
      Bindings        =   "R_bank5.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   675
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·»‰ﬂ :"
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
      Left            =   3750
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   675
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï :"
      Height          =   225
      Left            =   1635
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   285
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
      Left            =   3750
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   195
      Width           =   765
   End
End
Attribute VB_Name = "R_Bank5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim ChargTable As Recordset
Dim CodeTable As Recordset
Dim nOption As Integer
Dim cType As String
Dim cHead As String
Function MYVALID()
If Not IsDate(xdate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
If IsEmpty(xBank.BoundText) Then Exit Function
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xBank.BoundText = ""
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub Form_Load()

Set TempTable = tempdb.OpenRecordset("TEMP")

Select Case publicFlag
    Case 1
        cHead = "«Ê—«ﬁ ﬁ»÷ €Ì— „Õ’·… "
    Case 2
        cHead = "«Ê—«ﬁ œ›⁄ €Ì— „”œœ… "
End Select
Data1.DatabaseName = MdbPath
Data1.RecordSource = "FILE5_10"
xBank.BoundColumn = "ID"
xBank.ListField = "DESCA"
Data1.Refresh
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "Delete * From Temp"
Select Case publicFlag
    Case 1
        cString = "SELECT FILE5_10.DESCA AS BankDesc  , FILE5_20.DATE_1, FILE5_20.DATE_R, FILE5_20.CHK_ID,  " & _
                  " FILE3_10.DESCA AS ClienName , FILE5_20.SER_NO, FILE5_20.VALUE " & _
                  " FROM FILE5_10 RIGHT JOIN (FILE5_20 LEFT JOIN FILE3_10 ON FILE5_20.CODE = FILE3_10.CODE) ON  " & _
                  " FILE5_10.ID = FILE5_20.ID_BANK " & _
                  " Where Date_3 >= DateValue(" & MyParn(xdate1.Text) & ")" & _
                  " and Date_3 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                  " And File5_20.CLOSEd = '0' "
        If xBank.BoundText <> "" Then
            cString = cString & " And File5_20.ID_BANK = " & MyParn(xBank.BoundText)
        End If
        cString = cString & " ORDER BY FILE5_20.DATE_1 "
    Case 2
        cString = "SELECT FILE5_10.DESCA AS BankDesc  , FILE5_21.DATE_1, FILE5_21.DATE_R, FILE5_21.CHK_ID,  " & _
                  " FILE4_10.DESCA AS ClienName , FILE5_21.SER_NO, FILE5_21.VALUE " & _
                  " FROM FILE5_10 RIGHT JOIN (FILE5_21 LEFT JOIN FILE4_10 ON FILE5_21.CODE = FILE4_10.CODE) ON  " & _
                  " FILE5_10.ID = FILE5_21.ID_BANK " & _
                  " Where Date_3 >= DateValue(" & MyParn(xdate1.Text) & ")" & _
                  " and Date_3 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                  " And File5_21.CLOSEd = '0' "
        If xBank.BoundText <> "" Then
            cString = cString & " And File5_21.ID_BANK = " & MyParn(xBank.BoundText)
        End If
        cString = cString & " ORDER BY FILE5_21.DATE_1 "
End Select
Set ChargTable = mydb.CreateSnapshot(cString)
With ChargTable
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            TempTable.AddNew
            TempTable.str1 = .BANKDESC
            TempTable.str2 = .CHK_ID
            TempTable.str3 = .CLIENNAME
            TempTable.str4 = .SER_NO
            If publicFlag = 1 Then
                TempTable.str6 = "√Ê—«ﬁ ﬁ»÷ „Õ’·… "
            Else
                TempTable.str6 = "√Ê—«ﬁ œ›⁄ „”œœ… Œ·«· › —… "
            End If
            TempTable.str7 = " „‰  «—ÌŒ " & xdate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
            TempTable.Date1 = .date_3
            TempTable.date2 = .date_R
            TempTable.val1 = !Value
            TempTable.str19 = firsttitle
            TempTable.str20 = Secondtitle
            TempTable.Update
            .MoveNext
        Loop
    End If
End With
Report1.ReportFileName = PublicPath & "\Reports\BANKREP4.RPT"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Function TurnValue(pSource, pOld, pNew)
   TurnValue = IIf(pSource = pOld Or (IsNull(pSource) And IsNull(pOld)), pNew, pSource)
End Function
Function myDateValue(pDate)
myDateValue = "#" & DateValue(pDate) & "#"
End Function
