VERSION 5.00
Begin VB.Form FixCash 
   BackColor       =   &H00E0E0E0&
   Caption         =   " ”ÊÌ«  ‰ﬁœÌ…"
   ClientHeight    =   5025
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   5925
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
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
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox xDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3450
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1014
      Width           =   1740
   End
   Begin VB.TextBox xNameA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   150
      Locked          =   -1  'True
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1842
      Width           =   3240
   End
   Begin VB.PictureBox SSPanel2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   5925
      TabIndex        =   16
      Top             =   0
      Width           =   5925
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H80000018&
         Caption         =   "«÷«›…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3840
         TabIndex        =   22
         Top             =   90
         Width           =   870
      End
      Begin VB.CommandButton CmdInform 
         BackColor       =   &H80000018&
         Caption         =   "«” ⁄·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4800
         TabIndex        =   21
         Top             =   90
         Width           =   870
      End
      Begin VB.CommandButton CmdUndo 
         BackColor       =   &H00C0C0C0&
         Caption         =   " —«Ã⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2910
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   870
      End
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Õ–›"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1065
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   870
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Õ›Ÿ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1995
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   870
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0C0C0&
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
         Height          =   345
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   870
      End
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3450
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1842
      Width           =   1740
   End
   Begin VB.TextBox xDOC_NO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3450
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1740
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3450
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1428
      Width           =   1740
   End
   Begin VB.TextBox xValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3450
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2256
      Width           =   1740
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   5925
      TabIndex        =   6
      Top             =   4545
      Width           =   5925
      Begin VB.CommandButton CmdLast 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√ŒÌ—"
         Height          =   390
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   70
         Width           =   1065
      End
      Begin VB.CommandButton CmdFirst 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√Ê·"
         Height          =   390
         Left            =   2130
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   70
         Width           =   1140
      End
      Begin VB.CommandButton CmdNext 
         BackColor       =   &H00C0FFFF&
         Caption         =   "·«Õﬁ"
         Height          =   390
         Left            =   3255
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   70
         Width           =   1215
      End
      Begin VB.CommandButton CmdPrevious 
         BackColor       =   &H00C0FFFF&
         Caption         =   "”«»ﬁ"
         Height          =   390
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   70
         Width           =   1290
      End
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   150
      MaxLength       =   255
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2670
      Width           =   5040
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "„” ‰œ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5250
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1110
      Width           =   450
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1920
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "„”·”·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5250
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   675
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁÌ„…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5250
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2325
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5250
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1545
      Width           =   420
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5250
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3375
      Width           =   315
   End
End
Attribute VB_Name = "FixCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Dim CardTable As Recordset
Dim ClientTable As Recordset
Sub AddProc()
If Not formMode = addmode Then Handlecontrols addmode
formMode = addmode
myDefine
xDoc_No.Text = IncRec(myLastField(CardTable, "DOC_NO"))
End Sub
Sub EmptyProc()
formMode = EmptyMode
Handlecontrols EmptyMode
myDefine
xDoc_No.Text = "000001"
End Sub
Sub Handlecontrols(nMode)
Select Case nMode
Case Editmode
     cmdAdd.Enabled = True
     CmdDel.Enabled = True
     CmdInform.Enabled = True
     CmdExit.Enabled = True
     CmdSave.Enabled = True
     CmdUndo.Enabled = True
'    CmdPrevious.Enabled = True
'    CmdNext.Enabled = True
'    CmdLast.Enabled = True
'    CmdFirst.Enabled = True
     xCode.Enabled = False
Case addmode
    CmdInform.Enabled = False
    CmdDel.Enabled = False
    cmdAdd.Enabled = False
    CmdSave.Enabled = True
    CmdUndo.Enabled = True
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    CmdFirst.Enabled = False
    xCode.Enabled = True
Case EmptyMode
    CmdInform.Enabled = False
    CmdDel.Enabled = False
    cmdAdd.Enabled = False
    CmdSave.Enabled = True
    CmdUndo.Enabled = True
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    CmdFirst.Enabled = False
    xCode.Enabled = True
End Select
End Sub
Sub CardLookup()
Dim Generalarray(4)
Dim GrdArray(4)
Set Generalarray(1) = Me
If publicFlag = 1 Then
    Generalarray(2) = "Select DOC_NO as «·ﬂÊœ, [DATE] as [ «—ÌŒ] ,NAMEA as [≈”„ ],DescA as [»Ì«‰ ]From FILE8_20"
Else
    Generalarray(2) = "Select DOC_NO as «·ﬂÊœ, [DATE] as [ «—ÌŒ] ,NAMEA as [≈”„ ],DescA as [»Ì«‰ ]From FILE8_40"
End If
Generalarray(3) = " Where DescA Like '*cFilter*'"
Generalarray(4) = " ORDER BY DATE "

GrdArray(1) = 1200
GrdArray(2) = 1500
GrdArray(3) = 1500
GrdArray(4) = 2500
    
Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search.Caption = "«” ⁄·«„ "
Search.Show 1
End Sub
Sub editProc()
formMode = Editmode
Handlecontrols Editmode
End Sub
Sub myDefine()
xdesca.Text = ""
xDate.Text = ""
xCode.Text = ""
xNameA.Text = ""
xValue.Text = ""
xDoc.Text = ""
End Sub
Sub myProc()
If ActiveControl.Name = xCode.Name Then
    xCode.Text = GrdText(Search.Grid1, 0)
    xNameA.Text = GrdText(Search.Grid1, 1)
Else
    CardTable.FindFirst "DOC_NO = " & MyParn(GrdText(Search.Grid1, 0))
    MyLoad
End If
End Sub
Sub MyLoad()
xDoc_No.Text = CardTable.DOC_NO
xCode.Text = CardTable.CODE
xdesca.Text = TurnValue(CardTable.DESCA, Null, "")
ClientTable.FindFirst "Code = " & MyParn(xCode.Text)
xNameA.Text = IIf(ClientTable.NoMatch, "", TurnValue(ClientTable.DESCA, Null, ""))
xDoc.Text = TurnValue(CardTable.doc, Null, "")
xDate.Text = TurnValue(Format(CardTable!Date, "DD-MM-YYYY"), Null, "")
xValue.Text = TurnValue(CardTable!Value, Null, "")
End Sub
Sub MyReplace()
CardTable.FindFirst "DOC_NO = " & MyParn(xDoc_No.Text)
If CardTable.NoMatch Then
    CardTable.AddNew
Else
    CardTable.Edit
End If
CardTable.CODE = xCode.Text
CardTable.DESCA = TurnValue(xdesca.Text, "", Null)

CardTable.Value = Val(xValue.Text)
CardTable.Date = xDate.Text
CardTable.doc = TurnValue(xDoc.Text, "", Null)
CardTable.DOC_NO = TurnValue(xDoc_No.Text, "", Null)
CardTable.Update
    If publicFlag = "2" Then
        cString = " DELETE  FILE4_11.* FROM FILE4_11 WHERE FILE4_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '0' "
        mydb.Execute cString
    Else
        cString = " DELETE  FILE3_11.* FROM FILE3_11 WHERE FILE3_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '0' "
        mydb.Execute cString
    End If

    If publicFlag = "1" Then
        cString = "Insert Into File3_11(" & _
                  "[Type],Doc_Id,Code,[Date],Pay,DescA,SHOW )" & _
                  " Select '0',Doc_No,Code,[Date],[Value],' ”ÊÌ…' , '1' " & _
                  " From File8_20 WHERE FILE8_20.DOC_NO = " & MyParn(xDoc_No.Text)
        mydb.Execute cString
    Else
        cString = "Insert Into File4_11(" & _
                  "[Type],Doc_Id,Code,[Date],Pay,DescA,SHOW )" & _
                  " Select '0',Doc_No,Code,[Date],[Value],' ”ÊÌ…' , '1' " & _
                  " From File8_40 WHERE FILE8_40.DOC_NO = " & MyParn(xDoc_No.Text)
        mydb.Execute cString
    End If

End Sub
Function MYVALID()
If xDoc_No.Text = "" Then
    MsgBox " ”ÃÌ· „”·”· "
    Exit Function
End If

If xCode.Text = "" Then
    MsgBox " ”ÃÌ· ﬂÊœ"
    Exit Function
End If

If xDate.Text = "" Then
    MsgBox " ”ÃÌ·  «—ÌŒ "
    Exit Function
End If

If formMode <> Editmode Then
    CardTable.FindFirst "DOC_NO = " & MyParn(xDoc_No.Text)
    If Not CardTable.NoMatch Then Exit Function
End If
MYVALID = True
End Function
Private Sub CmdAdd_Click()
    AddProc
End Sub
Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    
    If publicFlag = "2" Then
        cString = " DELETE  FILE4_11.* FROM FILE4_11 WHERE FILE4_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '0' "
        mydb.Execute cString
    Else
        cString = " DELETE  FILE3_11.* FROM FILE3_11 WHERE FILE3_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '0' "
        mydb.Execute cString
    End If
    
    CardTable.Delete
    CardTable.Requery
    If CardTable.RecordCount > 0 Then
        CardTable.MoveLast
        MyLoad
    Else
        EmptyProc
    End If
End If
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    MyLoad
End If
End Sub
Private Sub cmdSave_Click()
msgBoxStr = IIf(addmove, "«÷«›… ”Ã· : Â· «‰  „Ê«›ﬁ ø", "Õ›Ÿ «· €ÌÌ—«  ! Â· √‰  „Ê«›ﬁ ø")
If Not MYVALID Then Exit Sub
If Not MsgBox(msgBoxStr, 4) = 6 Then
    CmdUndo_Click
    Exit Sub
End If
CardTable.FindFirst "Code = " & MyParn(xCode.Text)
If Not CardTable.NoMatch Then
    MyReplace
Else
    MyReplace
    AddProc
End If
End Sub
Private Sub CmdUndo_Click()
Select Case formMode
Case EmptyMode
    myDefine
Case addmode
    CardTable.MoveLast
    editProc
    MyLoad
Case Editmode
    MyLoad
End Select
End Sub
Private Sub Form_Load()
If publicFlag = 1 Then
    Set CardTable = mydb.OpenRecordset("SELECT * FROM file8_20 ORDER BY CODE ", dbOpenDynaset)
    Set ClientTable = mydb.OpenRecordset("file3_10", dbOpenSnapshot)
    Me.Caption = " ”ÊÌ«  ⁄„·«¡"
Else
    Set CardTable = mydb.OpenRecordset("SELECT * FROM file8_40 ORDER BY CODE ", dbOpenDynaset)
    Set ClientTable = mydb.OpenRecordset("file4_10", dbOpenSnapshot)
    Me.Caption = " ”ÊÌ«  „Ê—œÌ‰"
End If

If CardTable.RecordCount > 0 Then
     MyLoad
     editProc
Else
     EmptyProc
End If
End Sub
Private Sub xCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(3)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    If publicFlag = 1 Then
        Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê From File3_10"
    Else
        Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê From File4_10"
    End If
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 3000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub xCode_LostFocus()
    ClientTable.FindFirst "Code = " & MyParn(xCode.Text)
    xNameA.Text = IIf(ClientTable.NoMatch, "", TurnValue(ClientTable.DESCA, Null, ""))
End Sub

