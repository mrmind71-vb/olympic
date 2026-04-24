VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpitem2 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
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
   ScaleHeight     =   3075
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1830
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   90
      Width           =   6645
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1395
         Width           =   1815
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1000
         Width           =   1815
      End
      Begin VB.TextBox xitem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xstore 
         Height          =   315
         Left            =   1845
         TabIndex        =   13
         Top             =   620
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "„Œ“‰ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   675
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·’‰› :"
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
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "«·Ï :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1440
         Width           =   465
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
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1035
         Width           =   765
      End
      Begin VB.Label xitemDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   3705
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   5805
      Top             =   2025
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
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1845
      Width           =   3480
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   420
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   420
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   135
         Width           =   1140
      End
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·„Ê—œ :"
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
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1215
      Width           =   570
   End
End
Attribute VB_Name = "rpitem2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Sub myProc()
If ActiveControl.Name = XITEM.Name Then
    XITEM.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    xitemDesca.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
    Unload Search3
End If
End Sub
Private Sub cmdApply_Click()
If MYVALID Then doprint
End Sub
Private Sub CmdUndo_Click()
XITEM.Text = ""
xitemDesca.Caption = ""
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xitem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ItemsLookup
End Sub
Private Function MYVALID() As Boolean
If Not (IsDate(xdate1.Text) Or Trim(xdate1.Text) = "") Or Not (IsDate(xDate2.Text) Or Trim(xDate2.Text) = "") Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub Form_Load()
openCon con
data1.ConnectionString = strCon
data1.RecordSource = "Select Code,DescA From File0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
End Sub

Private Sub xitem_LostFocus()
xitemDesca.Caption = ""
If XITEM.Text = "" Then Exit Sub
xitemDesca.Caption = GetDesca("select desca from file1_10 where item = " & MyParn(XITEM.Text))
End Sub
Private Sub doprint()
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset, nBalance As Single
Dim aHeader(2)
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable


cString = "select file1_11.*,file1_12.desca from file1_11 LEFT join file1_12 on file1_11.type = file1_12.code " & _
          "where item = " & MyParn(XITEM.Text)

aHeader(0) = "[" & "··’‰› : " & xitemDesca.Caption & "]"
If IsDate(xdate1.Text) Then
    cString2 = "Select sum((file1_11.[in]) - (file1_11.out)) as balance " & _
                    " from file1_11 where date < " & DateSq(xdate1.Text) & _
                    " and ITEM = " & MyParn(XITEM.Text)

    If Trim(xStore.BoundText) <> "" Then
        cString2 = cString2 & turnFound(cString2) & " FILE1_11.STORE = " & MyParn(xStore.BoundText)
    End If
    nPrevious = Val(GetDesca(cString2) & "")
    
    cString = cString & turn(cString) & " FILE1_11.DATE >= " & DateSq(xdate1.Text)
    aHeader(1) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & " FILE1_11.DATE <= " & DateSq(xDate2.Text)
    aHeader(1) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If xStore.BoundText <> "" Then
    cString = cString & turn(cString) & " FILE1_11.STORE = " & MyParn(xStore.BoundText)
     aHeader(2) = "[·„Œ“‰ : " & xStore.Text & "]"
End If
cString = cString & " order by file1_11.[date],file1_11.[IN],file1_11.doc_id"

Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
If nPrevious <> 0 Then
        temptable.AddNew
        temptable!str1 = "—’Ìœ ”«»ﬁ"
        temptable!str2 = Null
        temptable!str7 = retHeader(aHeader, 0, 3)
        'temptable!Date1 = !Date
        If nPrevious < 0 Then temptable!val1 = Abs(nPrevious) Else temptable!val2 = nPrevious
        temptable!Val3 = nPrevious
        temptable.Update
End If


    Do Until .EOF
        temptable.AddNew
        If !Type = "6" Or !Type = "3" Then
            temptable!str1 = !desca & " " & GetDesca("SELECT DESCA FROM FILE3_10 WHERE CODE = " & MyParn(!CODECUST))
        ElseIf !Type = "T" Or !Type = "F" Then
            temptable!str1 = !desca & " " & GetDesca("SELECT DESCA FROM FILE0_40 WHERE CODE = " & MyParn(!code))
        Else
            temptable!str1 = !desca
        End If
        temptable!str2 = !doc_ID
        temptable!Date1 = !Date
        temptable!val1 = Val(!out & "")
        temptable!val2 = Val(!In & "")
        temptable!Val3 = nPrevious + Val(!In & "") - Val(!out & "")
        nPrevious = nPrevious + Val(!In & "") - Val(!out & "")
        temptable!str7 = retHeader(aHeader, 0, 3)
        temptable.Update
       .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    main.REPORT1.ReportFileName = App.Path & "\Reports\Item2.rpt"
    contemp.BeginTrans
    contemp.CommitTrans
    main.REPORT1.DataFiles(0) = tempFile
    main.REPORT1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select File1_10.item,File1_10.Desca From file1_10 "
Generalarray(2) = "Order by file1_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(FILE1_10.ITEM LIKE 'cFilter%' or  DESCA LIKE  'cFilter%') "


GrdArray(0, 0) = "ﬂÊœ «·’‰›"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·’‰›"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„ «·«’‰«›"
Search3.Show 1
End Sub

