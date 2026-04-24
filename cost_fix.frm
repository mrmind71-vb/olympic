VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cost_fixfrm 
   Caption         =   "÷»ÿ «· ﬂ·›…"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2700
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   45
      MaskColor       =   &H00FFFFFF&
      Picture         =   "cost_fix.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   " ﬂ·›… «·„»Ì⁄« "
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
      Left            =   1890
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1770
   End
   Begin VB.Frame frmProg1 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1620
      Width           =   3615
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   375
         Left            =   45
         TabIndex        =   9
         Top             =   180
         Visible         =   0   'False
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.StatusBar bar1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   7
      Top             =   2235
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
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
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   3570
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   2265
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   2265
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   780
      End
   End
End
Attribute VB_Name = "cost_fixfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nMode As Integer
Public nFlag As Integer
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
If nFlag = 0 Then
    fixSalesCost
ElseIf nFlag = 1 Then
    fixStockCost
End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub Form_Load()
xDate1.Text = Format(RetSetting("DATE1", TempSave(Me, nFlag & "")), "DD-MM-YYYY")
xdate2.Text = Format(RetSetting("DATE2", TempSave(Me, nFlag & "")), "DD-MM-YYYY")
openCon con
If nFlag = 0 Then
   cmdApply.Caption = " ﬂ·›… «·„»Ì⁄« "
ElseIf nFlag = 1 Then
   cmdApply.Caption = " ﬂ·›… «·Ã—œ"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
addSetting "DATE1", xDate1.Text, TempSave(Me, nFlag & "")
addSetting "DATE2", xdate2.Text, TempSave(Me, nFlag & "")
closeCon con
Set closefrm = Nothing
End Sub

Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_lostFocus()
myLostFocus xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_lostFocus()
myLostFocus xdate2
End Sub


Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xdate2_Validate(Cancel As Boolean)
myValidDate xdate2
End Sub
Private Sub fixSalesCost()
If MsgBox("÷»ÿ «· ﬂ·›…", vbOKCancel) <> vbOK Then Exit Sub
Dim nRecordCount As Long, i As Long, nAffect As Long, nTotal As Long, sCaption As String
sCaption = Me.Caption
Dim loctable As New ADODB.Recordset, cString As String, nCost As Double
cString = "SELECT FILE6_60.ITEM,FILE6_60.ID,FILE6_60H.DATE FROM FILE6_60 INNER JOIN FILE6_60H ON FILE6_60.DOC_NO = FILE6_60H.DOC_NO INNER JOIN SESSIONS ON FILE6_60H.SESSION = SESSIONS.CODE"
If IsDate(xDate1.Text) Then cString = cString & turn(cString) & "[SESSIONS].DATE_OPEN >= " & DateSq(xDate1.Text)
If IsDate(xdate2.Text) Then cString = cString & turn(cString) & "[SESSIONS].DATE_OPEN <= " & DateSq(xdate2.Text)
loctable.Open cString, con, adOpenStatic, adLockReadOnly
nRecordCount = loctable.RecordCount
prog1.Visible = True
prog1.Value = 0
con.BeginTrans
Do Until loctable.EOF
    i = i + 1
    Me.Caption = sCaption + " ”Ã· " & i & " „‰ " & nRecordCount
    prog1.Value = Round(i / (nRecordCount), 2) * 100
    cString = "UPDATE FILE6_60 SET FILE6_60.COST = " & LastCost_fin(loctable!Item, con, Format(loctable!Date, "YYYY-MM-DD"))
    cString = cString & turn(cString) & "FILE6_60.ID = " & loctable!ID
    con.Execute cString, nAffect
    nTotal = nTotal + nAffect
    loctable.MoveNext
Loop
con.CommitTrans
sCaption = Me.Caption
Inform " „ ÷»ÿ  ﬂ·›…" & nTotal & " ”Ã·"
bar1.Panels(1).Text = "⁄œœ «·”Ã·«  " & nTotal
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub fixStockCost()
If MsgBox("÷»ÿ «· ﬂ·›…", vbOKCancel) <> vbOK Then Exit Sub
Dim nRecordCount As Long, i As Long, nAffect As Long, nTotal As Long
Dim loctable As New ADODB.Recordset, cString As String, nCost As Double
cString = "SELECT FILE1_60.ITEM,FILE1_60.ID,FILE1_60H.DATE FROM FILE1_60 INNER JOIN FILE1_60H ON FILE1_60.DOC_NO = FILE1_60H.DOC_NO"
If IsDate(xDate1.Text) Then cString = cString & turn(cString) & "FILE1_60H.DATE >= " & DateSq(xDate1.Text)
If IsDate(xdate2.Text) Then cString = cString & turn(cString) & "FILE1_60H.DATE <= " & DateSq(xdate2.Text)
loctable.Open cString, con, adOpenStatic, adLockReadOnly
nRecordCount = loctable.RecordCount
prog1.Visible = True
prog1.Value = 0
con.BeginTrans
On Error GoTo myerror
Do Until loctable.EOF
    i = i + 1
    prog1.Value = Round(i / (nRecordCount), 2) * 100
    cString = "UPDATE FILE1_60 SET FILE1_60.COST = " & LastCost_fin(loctable!Item, con, Format(loctable!Date, "YYYY-MM-DD"))
    cString = cString & turn(cString) & "FILE1_60.ID = " & loctable!ID
    con.Execute cString, nAffect
    nTotal = nTotal + nAffect
    loctable.MoveNext
Loop
con.CommitTrans
Me.Caption = sCaption
Inform " „ ÷»ÿ  ﬂ·›…" & nTotal & " ”Ã·"
bar1.Panels(1).Text = "⁄œœ «·”Ã·«  " & nTotal
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

