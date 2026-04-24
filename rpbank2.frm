VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpBank2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ«—Ì— «·»‰Êﬂ"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   5865
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
      Top             =   1575
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
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   1575
      Width           =   1500
   End
   Begin VB.CommandButton cmdExit 
      Height          =   555
      Left            =   90
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1575
      Width           =   1500
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   5250
      Top             =   1725
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
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   5655
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1035
         Width           =   1365
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   630
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   390
         Left            =   450
         TabIndex        =   2
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
         Left            =   4725
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   690
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
         Left            =   4725
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   675
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ã„Ê⁄…"
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
         Left            =   4725
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   735
      End
      Begin VB.Label iLabel 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3300
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1125
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "rpBank2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If TypeOf ActiveControl Is DataCombo And KeyCode = 46 Then ActiveControl.BoundText = ""
End Sub
Private Sub Form_Load()
FixRpImage Me

openCon con
data1.ConnectionString = strCon
data1.RecordSource = "Select * From File5_50 ORDER BY CODE"
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"
End Sub
Private Sub cmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate1.Text) And Trim(xDate1.Text) <> "" Then Exit Function
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(1)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cField1 = myiif("[TYPE] <= 4.5 ", _
        " [value1]  - [value2]") & _
        " As Balance "

cField2 = myiif("[TYPE] = 5", _
        " [value1]") & _
        " As SumOfChq1 "

cField3 = myiif("[TYPE] = 6", _
        " [value2]") & _
        " As SumOfChq2 "

cString = "SELECT File5_10.CODE, File5_10.DESCA, File5_10.[group], FILE5_50.DESCA AS GroupDesca," & _
          cField1 & "," & _
          cField2 & "," & _
          cField3 & _
          " FROM (BANKMOVE INNER JOIN File5_10 ON BANKMOVE.BANK = File5_10.CODE) INNER JOIN FILE5_50 ON FILE5_10.[GROUP] = FILE5_50.CODE "

If Trim(xGroup.BoundText) <> "" Then
    cString = cString & turnFound(cString) & " File5_10.[Group] = " & MyParn(xGroup.BoundText)
    aHeader(0) = "[" & "„Ã„Ê⁄… : " & xGroup.Text & "]"
End If
          
If IsDate(xDate1.Text) Then
    cString = cString & turn(cString) & " date >= " & DateSq(xDate1.Text)
    aHeader(0) = BetweenString(xDate1.Text, xdate2.Text)
End If

If IsDate(xdate2.Text) Then
    cString = cString & turn(cString) & " date <= " & DateSq(xdate2.Text)
    aHeader(0) = BetweenString(xDate1.Text, xdate2.Text)
End If

cString = cString & " GROUP BY FILE5_10.CODE, File5_10.DESCA, File5_10.[group],file5_50.DESCA"
                         
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    If Abs(sourcetable!BALANCE) >= 0 Then
        temptable.AddNew
        temptable!str1 = sourcetable!CODE
        temptable!str2 = sourcetable![Desca]
        temptable!str3 = sourcetable![Group]
        temptable!str4 = sourcetable!GroupDesca
        If Val(sourcetable!BALANCE & "") >= 0 Then
            temptable!val1 = sourcetable!BALANCE
        Else
            temptable!val2 = Abs(sourcetable!BALANCE)
        End If
        temptable!val3 = sourcetable!sumOfChq1
        temptable!val4 = sourcetable!sumOfChq2
        temptable!str21 = TurnValue(retHeader(aHeader, 0, 2))
        temptable.Update
    End If
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\bank2.rpt"
main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.REPORT1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
