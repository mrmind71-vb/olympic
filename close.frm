VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form closefrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "√€·«ﬁ › —…"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   3390
      Begin VB.CheckBox xClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«€·«ﬁ «·„” ‰œ"
         Enabled         =   0   'False
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
         Height          =   330
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   990
         Value           =   1  'Checked
         Width           =   1410
      End
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
         Width           =   1860
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
         Width           =   1860
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·Ï  «—ÌŒ"
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
         Left            =   2085
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„‰  «—ÌŒ"
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
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   660
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   555
      Left            =   45
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1485
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   979
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "close.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdApply 
      Height          =   555
      Left            =   1665
      TabIndex        =   7
      Top             =   1485
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«€·«ﬁ › —…"
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   8
      Top             =   2115
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   45
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   714
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "closefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sFile As String, sFieldClose As String, sFieldDate As String, pFilter As String, sCaption As String
Public bTrans As Boolean
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
Me.MousePointer = 11
CloseData
Me.MousePointer = 0
End Sub

Private Function CloseData()
Dim nRecord As Long, nRecords As Long, cWhere As String
cString = "UPDATE " & sFile & " SET " & sFieldClose & " = " & xClosed.Value & " FROM " & sFile
If IsDate(xDate1.Text) Then cWhere = sFieldDate & " >= " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cWhere = cWhere & turn(cWhere, " and ") & sFieldDate & " <= " & DateSq(xDate2.Text)
cWhere = cWhere & turn(cWhere, " and ") & sFieldClose & " = " & IIf(xClosed.Value = 1, 0, 1)
If pFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & pFilter

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

con.BeginTrans
On Error GoTo myerror
If cString2 <> "" Then con.Execute cString2
con.Execute cString, nRecords
con.CommitTrans
If nRecords Then
    Inform " „ " & IIf(xClosed.Value = 1, "≈€·«ﬁ ", "› Õ ") & nRecords & " „” ‰œ" & turn(sCaption, " ") & sCaption
    panel1(0).Caption = "⁄œœ «·”Ã·«  " & IIf(xClosed.Value = 1, "«·„€·ﬁ… ", "«·„› ÊÕ… ") & nRecords
Else
    panel1(0).Caption = "·„ Ì „ " & IIf(xClosed.Value = 1, "≈€·«ﬁ", "› Õ") & " «Ì ”Ã·« "
End If
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con
If sFieldClose = "" Then sFieldClose = "[CLOSED]"
If sFieldDate = "" Then sFieldDate = "[DATE]"
Dim cString As String, aRet As Variant
cString = "Select min(" & sFieldDate & ") as MinOfDate,Max(" & sFieldDate & ") as MaxOfDate  FROM " & sFile
cString = cString & turn(cString) & sFieldClose & " = " & IIf(xClosed.Value = 0, 1, 0)
If pFilter <> "" Then cString = cString & " AND " & pFilter
aRet = GetFields(cString, con)
xDate1.Text = myFormat_p(retFlag(aRet, "MinOfDate"))
xDate2.Text = myFormat_p(retFlag(aRet, "MaxOfDate"))
openCon con
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set closefrm = Nothing
End Sub

Private Sub xClosed_Click()
cmdApply.Caption = IIf(xClosed.Value = 0, "› Õ", "«€·«ﬁ")
End Sub

Private Sub xDate1_DblClick()
Set datefrm.oDate = xDate1
datefrm.Show 1
End Sub

Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
End Sub

Private Sub xdate2_DblClick()
Set datefrm.oDate = xDate2
datefrm.Show 1
End Sub

Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
End Sub
Private Sub xdate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xdate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub

