VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form trans_membefrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕÊÌ· «·Ì ⁄÷Ê ⁄«„·"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "trans_member.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2145
   ScaleWidth      =   7650
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1575
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   90
      TabIndex        =   5
      Top             =   1215
      Width           =   2490
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "trans_member.frx":000C
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdChange 
         Height          =   510
         Left            =   1260
         TabIndex        =   8
         Top             =   180
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "trans_member.frx":232F
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "trans_member.frx":4D24
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1230
      Left            =   90
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "r"
         Height          =   375
         Left            =   4320
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "Item_Code_Changed"
         Top             =   675
         Width           =   1860
      End
      Begin VB.Label xCode_I 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   1860
      End
      Begin VB.Label xdescA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   4140
      End
      Begin VB.Label Lab_Item_Code_Changed 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈·Ì —ﬁ„"
         Height          =   270
         Left            =   6255
         TabIndex        =   3
         Top             =   765
         Width           =   555
      End
      Begin VB.Label Lab_Item_Code_Org 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„‰ —ﬁ„ "
         Height          =   270
         Left            =   6255
         TabIndex        =   2
         Top             =   315
         Width           =   570
      End
   End
End
Attribute VB_Name = "trans_membefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public sCode As String, sDesca As String
Private Sub cmdChange_Click()
If Not MYVALID Then Exit Sub
If MsgBox(" ÕÊÌ· «·Ì ⁄÷ÊÌ… ⁄«„·… ", vbOKCancel) <> vbOK Then Exit Sub
If myreplace Then Inform " „  ÕÊÌ· «·⁄÷ÊÌ… »‰Ã«Õ"
Exit Sub
myerror:
MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "[CODE]", addvalue(xCode.text))
aInsert = AddFlag(aInsert, "[CodeInstall]", "CODE")
aInsert = AddFlag(aInsert, "Title", "Title")
aInsert = AddFlag(aInsert, "Desca", "Desca")
aInsert = AddFlag(aInsert, "Date_birth", "Date_birth")
'aInsert = AddFlag(aInsert, "Date_Begin", "Date_Begin")

aInsert = AddFlag(aInsert, "SES_NO", "SES_NO")
aInsert = AddFlag(aInsert, "CODE_MAIN", "SES_NO")

aInsert = AddFlag(aInsert, "Died", "Died")
aInsert = AddFlag(aInsert, "[Drop]", "[Drop]")
aInsert = AddFlag(aInsert, "Gender", "Gender")
aInsert = AddFlag(aInsert, "Social", "Social")
aInsert = AddFlag(aInsert, "Religion", "Religion")
aInsert = AddFlag(aInsert, "Id_no", "Id_no")
aInsert = AddFlag(aInsert, "Address", "Address")
aInsert = AddFlag(aInsert, "Phone", "Phone")
aInsert = AddFlag(aInsert, "Mobil", "Mobil")

aInsert = AddFlag(aInsert, "Job", "Job")
aInsert = AddFlag(aInsert, "Phone_work", "Phone_work")
aInsert = AddFlag(aInsert, "NOTES", "Phone_work")
aInsert = AddFlag(aInsert, "Degree", "Degree")
aInsert = AddFlag(aInsert, "Region", "Region")
aInsert = AddFlag(aInsert, "company", "company")
aInsert = AddFlag(aInsert, "Job_desca", "Job_desca")
aInsert = AddFlag(aInsert, "Type", "Type")
aInsert = AddFlag(aInsert, "[USERNAME]", "[USERNAME]")
aInsert = AddFlag(aInsert, "[TIME]", "[TIME]")
aInsert = AddFlag(aInsert, "[USERNAME2]", "[USERNAME2]")
aInsert = AddFlag(aInsert, "[TIME2]", "[TIME2]")



cInsert = addInsertTable(aInsert, "FILE1_10", "FILE2_10", "FILE2_10.CODE = " & sCode) & ";"


aInsert = AddFlag(Empty, "[MEMBER]", addvalue(xCode.text))
aInsert = AddFlag(aInsert, "CODE", "CODE")
aInsert = AddFlag(aInsert, "RELATION", "RELATION")
aInsert = AddFlag(aInsert, "GENDER", "GENDER")
aInsert = AddFlag(aInsert, "TITLE", "TITLE")
aInsert = AddFlag(aInsert, "DESCA", "DESCA")
aInsert = AddFlag(aInsert, "DATE_BIRTH", "DATE_BIRTH")
'aInsert = AddFlag(aInsert, "DATE_BEGIN", "DATE_BEGIN")
aInsert = AddFlag(aInsert, "NOTES", "NOTES")
aInsert = AddFlag(aInsert, "HANDI", "HANDI")
aInsert = AddFlag(aInsert, "PENDING", "PENDING")
cInsert = cInsert & ";" & _
        addInsertTable(aInsert, "FILE1_11", "FILE2_11", "FILE2_11.MEMBER = " & sCode) & ";"
con.BeginTrans
On Error GoTo myerror
con.Execute cInsert
con.CommitTrans
TransPhoto
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function TransPhoto()
Dim fs As New FileSystemObject
If validPhoto(RetPhoto_I(sCode)) Then
    fs.CopyFile RetPhoto_I(sCode), retPhoto(xCode.text)
End If
Dim loctable As New ADODB.Recordset
loctable.Open "select code from file2_11 where member = " & addvalue(sCode), con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    If validPhoto(RetAppendPhoto_i(sCode, loctable!code)) Then
        fs.CopyFile RetAppendPhoto_i(sCode, loctable!code), RetAppendPhoto(xCode.text, loctable!code)
    End If
    loctable.MoveNext
Loop
Set loctable = Nothing
End Function
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
'createAddInsertTable "SELECT * FROM FILE2_10", con
createAddInsertTable "SELECT * FROM FILE2_11", con
End Sub
Private Sub Form_Load()
openCon con
xCode_I.Caption = sCode
xDesca.Caption = sDesca
xCode.text = Newflag("FILE1_10", "CODE", con)
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set trans_memberfrm = Nothing
End Sub
Private Function MYVALID() As Boolean
Dim aRet As Variant
If Trim(xCode_I.Caption) = "" Then Exit Function

If Not ValidNum(xCode.text) Then
    MsgBox "—ﬁ„ €Ì— ’ÕÌÕ"
    Exit Function
End If

aRet = GetField("select code from file1_10 where codeinstall = " & addvalue(xCode_I.Caption), con)
If Not IsEmpty(aRet) Then
    If MsgBox("—ﬁ„ ⁄÷ÊÌ…  „ ‰ﬁ·Â „‰ ﬁ»· »—ﬁ„ " & aRet & " ‰ﬁ· ⁄·Ì «Ì Õ«·", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
End If


aRet = Member_Load(xCode.text, "CODE", con)
If Not IsEmpty(aRet) Then
    MsgBox "—ﬁ„ ⁄÷ÊÌ… „”Ã· „‰ ﬁ»· ··⁄÷Ê " & retFlag(aRet, "desca")
    Exit Function
End If
MYVALID = True
End Function
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
End Sub
