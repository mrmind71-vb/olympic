VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form ChqRep5 
   Caption         =   " ﬁ«—Ì— «·‘Ìﬂ« "
   ClientHeight    =   2070
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
   ScaleHeight     =   2070
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   915
      Width           =   1155
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2385
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   915
      Width           =   1230
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   3360
      Top             =   1350
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
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -975
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   405
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -945
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   45
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1275
      Width           =   2415
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   360
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
   End
   Begin MSDBCtls.DBCombo xGroup1 
      Bindings        =   "Chqrep5.frx":0000
      Height          =   315
      Left            =   330
      TabIndex        =   2
      Top             =   105
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DBCombo1"
      RightToLeft     =   -1  'True
   End
   Begin MSDBCtls.DBCombo xGroup2 
      Bindings        =   "Chqrep5.frx":0014
      Height          =   315
      Left            =   315
      TabIndex        =   3
      Top             =   480
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DBCombo1"
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï :"
      Height          =   225
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   960
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
      TabIndex        =   6
      Top             =   960
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·Ï „Ã„Ê⁄… :"
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
      Left            =   3615
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   555
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰ „Ã„Ê⁄… :"
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
      Left            =   3660
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   195
      Width           =   1005
   End
End
Attribute VB_Name = "ChqRep5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChqTable As Recordset
Dim TempTable As Recordset
Dim GrClient As Recordset
Function MYVALID()
If Not IsDate(xDate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
If xGroup1.BoundText = "" Then Exit Function
If xGroup2.BoundText = "" Then Exit Function
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xGroup1.BoundText = ""
xGroup2.BoundText = ""
xDate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub Form_Load()

Set TempTable = tempdb.OpenRecordset("TEMP")

Data1.DatabaseName = MdbPath
If lCust Then
    Data1.RecordSource = "SELECT * FROM FILE1_70 WHERE FLAG = 3 "
Else
    Data1.RecordSource = "SELECT * FROM FILE1_70 WHERE FLAG = 4 "
End If

xGroup1.ListField = "Desca"
xGroup1.BoundColumn = "Code"
Data1.Refresh

Data2.DatabaseName = MdbPath
If lCust Then
    Data2.RecordSource = "SELECT * FROM FILE1_70 WHERE FLAG = 3 "
Else
    Data2.RecordSource = "SELECT * FROM FILE1_70 WHERE FLAG = 4 "
End If

xGroup2.ListField = "Desca"
xGroup2.BoundColumn = "Code"
Data2.Refresh
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
tempdb.Execute "Delete * From Temp"
nOption = 1

cHead = " ≈Ã„«·Ï „Êﬁ› «·‘Ìﬂ«  "
If lCust Then
    Set GrClient = mydb.CreateSnapshot("Select * From File1_70 Where Flag = 3")
Else
    Set GrClient = mydb.CreateSnapshot("Select * From File1_70 Where Flag = 4")
End If
             
With GrClient
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .DESCA
    TempTable.str2 = " ‘—ﬂ… ‰«„ﬂÊ "
    TempTable.str3 = cHead
    TempTable.str4 = " „‰  «—ÌŒ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
    TempTable.val1 = CalcChq(1, .code)
    TempTable.VAL2 = CalcChq(2, .code)
    TempTable.VAL3 = CalcChq(3, .code)
    TempTable.VAL4 = CalcChq(4, .code)
    TempTable.VAL5 = CalcChq(5, .code)
    TempTable.VAL6 = CalcChq(6, .code)
    TempTable.VAL7 = CalcChq(7, .code)
    TempTable.VAL8 = TempTable.VAL2 + TempTable.VAL4
    TempTable.str19 = firsttitle
    TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop

End If
End With
Report1.ReportFileName = PublicPath & "\Reports\CHQREP5.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Function CalcChq(cPass, cCode)
    NRETURN = 0
    CalcChq = NRETURN
    If lCust Then
        Select Case cPass
            Case 1
                cString = "select file5_20.* , file3_10.GROUP" & _
                          " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '0' and date_1 < DateValue(" & MyParn(xDate1.Text) & ")" & _
                          " and File3_10.GROUP = " & MyParn(cCode)
      
            Case 2
                cString = "select file5_20.* , file3_10.GROUP" & _
                          " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '2' and date_1 < DateValue(" & MyParn(xDate1.Text) & ") and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") and date_3 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                          " and File3_10.GROUP = " & MyParn(cCode)
            Case 3
                cString = "select file5_20.* , file3_10.GROUP" & _
                          " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE date_1 >= DateValue(" & MyParn(xDate1.Text) & ") And date_1 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                          " and File3_10.GROUP = " & MyParn(cCode)
            Case 4
                cString = "select file5_20.* , file3_10.GROUP" & _
                          " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '2' and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") And date_3 <= DateValue(" & MyParn(xDate2.Text) & ") And date_1 >= DateValue(" & MyParn(xDate1.Text) & ") and date_1 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                          " and File3_10.GROUP = " & MyParn(cCode)
            Case 5
                cString = "select file5_20.* , file3_10.GROUP" & _
                          " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '1' and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") And date_3 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                          " and File3_10.GROUP = " & MyParn(xGroup1.BoundText) & " and " & MyParn(cCode)
            Case 6
                cString = "select file5_20.* , file3_10.GROUP" & _
                          " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '0' and date_1 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                          " and File3_10.GROUP = " & MyParn(cCode)
            Case 7
                cString = "select file5_20.* , file3_10.GROUP" & _
                          " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '0' " & _
                          " and File3_10.GROUP = " & MyParn(cCode)
        End Select
    Else
        Select Case cPass
            Case 1
                cString = "select file5_21.* , file4_10.Group AS Cl_Group " & _
                          " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE WHERE CLOSED = '0' and date_1 < DateValue(" & MyParn(xDate1.Text) & ")" & _
                          " and File4_10.GROUP = " & MyParn(cCode)
      
            Case 2
                cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                          " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE WHERE CLOSED = '2' and date_1 < DateValue(" & MyParn(xDate1.Text) & ") and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") and date_3 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                          " and File4_10.GROUP = " & MyParn(cCode)
            Case 3
                cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                          " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE WHERE date_1 >= DateValue(" & MyParn(xDate1.Text) & ") And date_1 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                          " and File4_10.GROUP = " & MyParn(cCode)
            Case 4
                cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                          " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE WHERE CLOSED = '2' and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") And date_3 <= DateValue(" & MyParn(xDate2.Text) & ") And date_1 >= DateValue(" & MyParn(xDate1.Text) & ") and date_1 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                          " and File4_10.GROUP = " & MyParn(cCode)
            Case 5
                cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                          " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE WHERE CLOSED = '1' and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") And date_3 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                          " and File4_10.GROUP = " & MyParn(xGroup1.BoundText) & " and " & MyParn(cCode)
            Case 6
                cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                          " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE WHERE CLOSED = '0' and date_1 <= DateValue(" & MyParn(xDate2.Text) & ")" & _
                          " and File4_10.GROUP = " & MyParn(cCode)
            Case 7
                cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                          " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE WHERE CLOSED = '0' " & _
                          " and File4_10.GROUP = " & MyParn(cCode)
        End Select
    
    End If
    Set ChqTable = mydb.CreateSnapshot(cString)
    
    If ChqTable.RecordCount > 0 Then
        ChqTable.MoveFirst
        Do While Not ChqTable.EOF
            NRETURN = NRETURN + ChqTable.[Value]
            ChqTable.MoveNext
        Loop
        CalcChq = NRETURN
    End If
    CalcChq = NRETURN
End Function
Function myDateValue(pDate1)
myDateValue = "DateValue('" & pDate1 & "')"
End Function
