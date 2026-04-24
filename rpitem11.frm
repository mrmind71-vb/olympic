VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpitem11 
   Caption         =   "ÿ»«⁄… "
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   2940
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfix 
      Caption         =   "÷»ÿ «· ﬂ·›…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2475
      Width           =   2445
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   0
      Width           =   6180
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   855
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1620
         Width           =   1725
      End
      Begin VB.TextBox xDoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1980
         Width           =   1680
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1620
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xstore 
         Height          =   315
         Left            =   855
         TabIndex        =   3
         Top             =   1260
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   855
         TabIndex        =   2
         Top             =   900
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroupMain 
         Height          =   315
         Left            =   855
         TabIndex        =   1
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   855
         TabIndex        =   0
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ :"
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
         Index           =   3
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì… :"
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·Ã—œ :"
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
         Index           =   2
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2025
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«· «—ÌŒ :"
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
         Index           =   1
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1710
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄… :"
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
         Index           =   0
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label4 
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1350
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   " ›—Ì€"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1530
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2475
      Width           =   1140
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "«” Ã«»…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2475
      Width           =   1185
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2475
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   5175
      Top             =   3015
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   180
      Top             =   630
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   2655
      Top             =   2520
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
   Begin VB.Label Label6 
      Height          =   255
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2175
      Width           =   1005
   End
End
Attribute VB_Name = "rpitem11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearch As New Search3
Private Sub cmdApply_Click()
doprint
End Sub
Private Sub CmdClear_Click()
xGroup.BoundText = ""
xStore.BoundText = ""
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "Select Code,DescA From File1_10SC ORDER BY DESCA"
Set xSection.RowSource = data1
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

data2.ConnectionString = strCon
data2.RecordSource = "Select Code,DescA From File1_50G order by Desca"
Set xGroupMain.RowSource = data2
xGroupMain.ListField = "Desca"
xGroupMain.BoundColumn = "Code"

data3.ConnectionString = strCon
data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
Set xGroup.RowSource = data3
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

data4.ConnectionString = strCon
data4.RecordSource = "Select Code,DescA From File0_40"
Set xStore.RowSource = data4
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xdate1.Text) And Trim(xdate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Function
End If
MYVALID = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xdate1_GotFocus()
myGotFocus xdate1
End Sub

Private Sub xdate1_Validate(Cancel As Boolean)
myValidDate xdate1
End Sub

Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xdate1
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
End Sub

Private Sub xdate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub

Private Sub xdoc_no_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then doclookup
End Sub

Private Sub xGroupMain_Validate(Cancel As Boolean)
If Not xGroupMain.MatchedWithList Then xGroupMain.BoundText = ""
data3.RecordSource = "Select Code,DescA From File1_50 " & IIf(xGroupMain.BoundText <> "", " where file1_50.[GROUP] = " & xGroupMain.BoundText, "")
data3.Refresh
End Sub
Private Sub cmdFix_Click()
Dim oCost As New cost_fixfrm
oCost.nFlag = 1
oCost.Show 1
End Sub
Private Sub FixCost()
    cCaption = Me.Caption
    openCon con
    Dim loctable As New ADODB.Recordset
    cString = "Select FILE0_10.ITEM,FILE0_10H.DATE,FILE0_10.ID FROM (FILE0_10 INNER JOIN FILE0_10H ON FILE0_10.DOC_NO = FILE0_10H.DOC_NO) INNER JOIN FILE1_10 ON FILE0_10.ITEM = FILE1_10.ITEM"
    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    
    nCount = loctable.RecordCount
    Dim I As Long
    con.BeginTrans
    Do Until loctable.EOF
        I = I + 1
        Me.Caption = cCaption & I & "  „‰ " & nCount
        nCost = LastCostDate(loctable!Item, Format(loctable!Date, "yyyy-mm-dd"), con)
        If nCost <> 0 Then
            con.Execute " UPDATE FILE0_10 SET FILE0_10.COST = " & nCost & _
                        " WHERE FILE0_10.ID = " & loctable!ID
        End If
        loctable.MoveNext
    Loop
    con.CommitTrans
    Me.Caption = cCaption
    MsgBox " „ ÷»ÿ «· ﬂ·›… »‰Ã«Ã"
lastsub:
    closeCon con
    Exit Sub
myerror:
Me.Caption = cCaption
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
    GoTo lastsub
End Sub






Private Sub xITEM_Change()

End Sub

Private Sub xDoc_No_LostFocus()
If Trim(xDoc_No.Text) <> "" Then xDoc_No = RetZero(xDoc_No.Text)
End Sub
Private Sub doclookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE,CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA " & _
                  " FROM FILE0_10H INNER JOIN FILE0_40 ON FILE0_10H.Store = FILE0_40.CODE "

Generalarray(2) = "Order by [Date]"
Generalarray(3) = 4200
Generalarray(5) = True


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ-«·„Œ“‰"
listarray(0, 1) = "@@Doc_No@@6 or  %%FILE0_40.DESCA%% OR " & _
                  "##[DATE]##"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·„Œ“‰"
GrdArray(3, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "«” ⁄·«„"
oSearch.Show 1
End Sub
Sub myProc()
xDoc_No.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
oSearch.Hide
End Sub
Private Sub doprint()
Dim aHeader(5)
If Not MYVALID Then Exit Sub
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

If Trim(xDoc_No.Text) <> "" Then
    cString = "SELECT FILE0_10.DIFFER,FILE1_10.ITEM,FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP] AS FILE1_50GROUP,FILE0_10.COST, FILE1_50.DESCA AS FILE1_50DESCA,FILE1_50G.DESCA AS FILE1_50GDESCA" & _
              " FROM (((FILE0_10H INNER JOIN FILE0_10 ON FILE0_10H.DOC_NO = FILE0_10.DOC_NO) INNER JOIN FILE1_10 ON FILE0_10.ITEM = FILE1_10.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_50G ON FILE1_50.[GROUP] = FILE1_50G.CODE  where differ <> 0"
Else
    cString = "SELECT Sum(FILE0_10.DIFFER) AS SUMOFDIFFER,FILE0_10.ITEM,FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP] AS FILE1_50GROUP,Sum(FILE0_10.DIFFER * FILE0_10.COST) AS TOTALCOST, FILE1_50.DESCA AS FILE1_50DESCA,FILE1_50G.DESCA AS FILE1_50GDESCA" & _
              " FROM (((FILE0_10H INNER JOIN FILE0_10 ON FILE0_10H.DOC_NO = FILE0_10.DOC_NO)INNER JOIN FILE1_10 ON FILE0_10.ITEM = FILE1_10.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_50G ON FILE1_50.[GROUP] = FILE1_50G.CODE "
End If

If IsNumeric(xSection.BoundText) Then
    cString = cString & turn(cString) & "File1_10.[SECTION] = " & xSection.BoundText
    aHeader(0) = "«·ﬁ”„ : " & xSection.Text
End If

If IsDate(xdate1.Text) Then
    cString = cString & turn(cString) & " date >= " & DateSq(xdate1.Text)
    aHeader(1) = BetweenString(xdate1.Text, xDate2.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & " date <= " & DateSq(xDate2.Text)
    aHeader(1) = BetweenString(xdate1.Text, xDate2.Text)
End If

If Trim(xGroup.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "File1_10.[GROUP] = " & xGroup.BoundText
    aHeader(2) = "„Ã„Ê⁄… " & xGroup.Text & "]"
End If

If Trim(xGroupMain.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "File1_50.[GROUP] = " & xGroupMain.BoundText
    aHeader(3) = "„Ã„Ê⁄… —∆Ì”Ì…" & xGroup.Text & "]"
End If


If Trim(xStore.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "File0_10H.store = " & MyParn(xStore.BoundText)
    aHeader(4) = "[" & "«·„Œ“‰ " & xStore.Text & "]"
End If

If Trim(xDoc_No.Text) <> "" Then
    cString = cString & turnFound(cString) & "File0_10H.DOC_NO = " & MyParn(xDoc_No.Text)
    aHeader(5) = "[" & "—ﬁ„ «·„” ‰œ : " & xDoc_No.Text & "]"
End If

If Trim(xDoc_No.Text) = "" Then
    cString = cString & " GROUP BY FILE0_10.ITEM, FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP], FILE1_50.DESCA,FILE1_50G.DESCA"
    cString = cString & turn(cString, " having ", " and ") & " sum(Differ) <> 0"
End If
          
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim cCondition As Boolean
With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!val11 = !FILE1_50GROUP
        temptable!str7 = ![file1_50GDESCA]
        temptable!val12 = ![Group]
        temptable!str8 = ![file1_50desca]
        temptable!str1 = !Item
        temptable!str2 = ![desca]
        
        If Trim(xDoc_No.Text) <> "" Then
            temptable!val1 = !Differ
            temptable!val2 = !cost
            temptable!Val3 = !Differ * !cost
       Else
            temptable!val1 = !SumofDiffer
            temptable!Val3 = !TotalCost
       End If
        temptable!str21 = TurnValue(retHeader(aHeader, 0, 6))
        temptable.Update
      .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    If xDoc_No.Text <> "" Then
        main.REPORT1.ReportFileName = App.Path & "\Reports\Item11.rpt"
    Else
        main.REPORT1.ReportFileName = App.Path & "\Reports\Item11T.rpt"
    End If
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
