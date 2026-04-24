VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PartFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁœÌ…"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13875
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
   ScaleHeight     =   8085
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   465
      ScaleWidth      =   13875
      TabIndex        =   16
      Top             =   7320
      Width           =   13875
      Begin VB.TextBox xtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   45
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CommandButton cmdfirst 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   45
         Width           =   435
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   945
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "”«»ﬁ"
         Top             =   45
         Width           =   435
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   " «·Ì"
         Top             =   45
         Width           =   435
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "«ŒÌ—"
         Top             =   45
         Width           =   435
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   15
      Top             =   7785
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:00 ’"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   2745
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   3630
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   3465
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   8370
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   5490
      Begin VB.CommandButton CmdInform 
         Caption         =   "≈” ⁄·«„"
         Height          =   390
         Left            =   4125
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   150
         Width           =   1290
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "„” ‰œ ÃœÌœ"
         Height          =   390
         Left            =   2775
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   150
         Width           =   1290
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         CausesValidation=   0   'False
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   150
         Width           =   1290
      End
      Begin VB.CommandButton CmdDelInv 
         Caption         =   "Õ–› «·„” ‰œ"
         Height          =   390
         Left            =   1425
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   150
         Width           =   1290
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   9090
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   585
      Width           =   4740
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1290
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   555
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   210
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   585
      Width           =   1500
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   585
         Width           =   1365
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1365
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   5595
      Left            =   45
      TabIndex        =   22
      Top             =   1665
      Width           =   13785
      _cx             =   24315
      _cy             =   9869
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "PartFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte
Dim CardTable As ADODB.Recordset, GRDTABLE As New ADODB.Recordset
Dim DocFileName As String, DocFileHeader As String, sName As String
Dim docMoveType As String
Dim DocTitle As String
Dim DocClient As String, CGROUP As String
Dim DocFileMove As String
Dim dLastdate As String, defBox As String
Dim DocField As String, dDateLast As String
Dim formMode
Dim lCellButton As Boolean
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim nTry As Integer
On Error Resume Next

CON.BeginTrans
For nTry = 1 To 10
    If xDoc_No.Enabled Then
        CON.Execute "insert into " & DocFileHeader & "( DOC_NO,[DATE],username)" & _
        "Values(" & _
        addstring(xDoc_No.Text) & "," & _
        DateSq(xDate.Text) & "," & _
        addstring(cUserName) & _
        ")"
    Else
        CON.Execute "update " & DocFileHeader & " Set " & _
        " [DATE] = " & DateSq(xDate.Text) & "," & _
        " USERNAME = " & addstring(cUserName) & _
        " WHERE DOC_NO = " & addstring(xDoc_No.Text)
    End If
    If Err.Number = 0 Then
        ' „—«Ã⁄… «·„” ‰œ
        CON.Execute "Delete * From " & DocFileName & " where Doc_No = " & MyParn(xDoc_No.Text) & " and  row > " & Grid1.Rows - 2
        
        With Grid1
            For i = 1 To .Rows - 2
                CON.Execute "Insert Into " & DocFileName & "(Doc_No,Charge,[Value],Box,Desca,row) " & _
                            " Values(" & _
                            addstring(xDoc_No.Text) & "," & _
                            addstring(.TextMatrix(i, 1)) & "," & _
                            addvalue(.TextMatrix(i, 4)) & "," & _
                            addstring(.TextMatrix(i, 0)) & "," & _
                            addstring(.TextMatrix(i, 3)) & "," & _
                            i & _
                            ")"
            
           If Err.Number = -2147467259 Then
                    Err.Clear
                    CON.Execute "update " & DocFileName & " set " & _
                        "[date] = " & DateSq(xDate.Text) & "," & _
                        " Charge = " & addstring(.TextMatrix(i, 1)) & "," & _
                        " [value] = " & Val(.TextMatrix(i, 4)) & "," & _
                        " box = " & addstring(.TextMatrix(i, 0)) & "," & _
                        " desca = " & addstring(.TextMatrix(i, 3)) & _
                        " where doc_no = " & MyParn(xDoc_No.Text) & _
                        " and [row] = " & i
                End If
                If Err.Number <> 0 Then GoTo MyError
            Next
        End With
    End If
    
    If Err.Number = 0 Then Exit For
    If Err.Number = -2147467259 And nTry < 10 Then
        Err.Clear
        xDoc_No.Text = RetZero(Val(xDoc_No.Text) + 1)
    End If
    If Err.Number <> 0 Then GoTo MyError
Next
CON.CommitTrans
myreplace = True
Exit Function
MyError:
CON.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
If ActiveControl.Name = Grid1.Name Then
    If Grid1.Col = 1 Then
        Grid1.TextMatrix(Grid1.row, 1) = Search3.Grid1.TextMatrix(Search3.Grid1.row, 0)
        GrdDesc Grid1.row
        If Grid1.row = Grid1.Rows - 1 Then
            Grid1.AddItem ""
            Grid1.TextMatrix(Grid1.Rows - 1, 0) = defBox
        End If
        Unload Search3
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_No = " & MyParn(Search3.Grid1.TextMatrix(Search3.Grid1.row, 0)), , adSearchForward, adBookmarkFirst
    MyLoad
    Unload Search3
End If
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo MyError
    CON.BeginTrans
    CON.Execute "Delete * From " & DocFileName & " where Doc_No = " & MyParn(xDoc_No.Text)
    CON.Execute "Delete * From " & DocFileHeader & " where Doc_No = " & MyParn(xDoc_No.Text)
    CON.CommitTrans
    CardTable.Requery
    If CardTable.EOF And CardTable.EOF Then
        myDefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        MyLoad
    End If
End If
Exit Sub
MyError:
CON.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
cString = "SELECT " & DocFileHeader & ".Doc_No, Format(" & DocFileHeader & ".Date,'dd-mm-yyyy'),First(" & DocClient & ".Desca)" & _
          " FROM (" & DocFileHeader & " inner join " & DocFileName & " on " & DocFileHeader & ".doc_no = " & DocFileName & ".Doc_NO) Inner Join " & DocClient & " on " & DocFileName & ".Charge = " & DocClient & ".Code"
          
Generalarray(1) = cString
Generalarray(2) = " group by " & DocFileHeader & ".Doc_No," & DocFileHeader & ".Date order by " & DocFileHeader & ".Doc_No," & DocFileHeader & ".Date"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·«”„- «—ÌŒ «·„” ‰œ"
listarray(0, 1) = "(%%" & DocClient & ".Desca%% or " & _
                  " ##" & DocFileHeader & ".Date##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·≈”„"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "Customers Query"
Search3.Show 1
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
Private Sub CmdNewInv_Click()
myDefine
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
MsgBox " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Requery
'If xDoc_No.Enabled Then
'    CmdNewInv_Click
'Else
    CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    MyLoad
'End If
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    myDefine
Else
    If xDoc_No.Enabled Then CardTable.MoveLast Else CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    MyLoad
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
bEdit = True
Select Case myPublic
    Case 1 '„’«—Ì›
        sName = "«·„’—Ê›"
        DocFileName = "File8_50"
        DocFileHeader = "FILE8_50H"
        DocClient = "FILE8_51"
    Case 2 '«·«Ì—«œ
        sName = "«·«Ì—«œ"
        DocFileName = "File8_60"
        DocFileHeader = "FILE8_60H"
        DocClient = "FILE8_61"
End Select
Me.Caption = DocTitle

Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM " & DocFileHeader & " ORDER BY DOC_NO", CON, adOpenStatic, adLockOptimistic, adCmdText
GRDTABLE.Open "Select " & DocFileName & ".*, " & DocClient & ".desca as ChargeDesca " & " From " & DocFileName & " left join " & DocClient & " on " & DocFileName & ".charge = " & DocClient & ".code", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
With Grid1
    .Cols = 5
    .Rows = 2
    .TextMatrix(1, 0) = True
    .Editable = flexEDKbdMouse
    If myPublic = 1 Or myPublic = 3 Then
        .FormatString = "Œ“‰…|" & "«·„’—Ê›|" & "Ê’› «·„’—Ê›|" & "«·»Ì«‰|" & "«·ﬁÌ„…"
    Else
        .FormatString = "Œ“‰…|" & "«·«Ì—«œ|" & "Ê’› «·«Ì—«œ|" & "«·»Ì«‰|" & "«·ﬁÌ„…"
    End If
    .ColWidth(0) = 1300
    .ColWidth(1) = 1000
    .ColWidth(2) = 2900
    .ColWidth(3) = 4000
    .ColWidth(4) = 1000
    For i = 1 To Grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    '.ColHidden(1) = Not (myPublic = 1)
    .ColComboList(0) = StrBox
    '.ColHidden(2) = Val(GetDesca("Select Sum(1) From file0_50")) <= 1
    '.ColHidden(2) = True
End With
defBox = RetDefBox
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    MyLoad
Else
    myDefine
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
GRDTABLE.Close
Set CardTable = Nothing
Set GRDTABLE = Nothing
End Sub

Private Sub Grid1_AfterEdit(ByVal row As Long, ByVal Col As Long)
If Grid1.TextMatrix(row, 1) <> "" Then Grid1.TextMatrix(row, 1) = RetZero(Grid1.TextMatrix(row, 1), 3)
If Col = 1 Then GrdDesc row
If Grid1.Col = 0 And row = Grid1.Rows - 2 Then Grid1.TextMatrix(Grid1.Rows - 1, 0) = Grid1.TextMatrix(Grid1.Rows - 2, 0)
CalcTotals
End Sub
Private Sub Grid1_EnterCell()
If Grid1.Col = 2 Or Grid1.Col = 7 Then Grid1.Editable = flexEDNone Else Grid1.Editable = flexEDKbdMouse
End Sub

Private Sub Grid1_GotFocus()
If Grid1.row = 0 Then
    Grid1.SetFocus
    Grid1.Select 1, 0
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And Grid1.Col = 1 Then grdLookup
If KeyCode = 46 And Grid1.row <> Grid1.Rows - 1 And Grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        Grid1.RemoveItem Grid1.row
    End If
End If
End Sub
Private Sub grid1_KeyDownEdit(ByVal row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And Col = 2 Then grdLookup
If KeyCode = 46 And row <> Grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        Grid1.RemoveItem row
    End If
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case Grid1.Col
    Case 0
        If KeyCode = 112 Then grdLookup
End Select
End Sub
Private Sub Grid1_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
If Grid1.row = Grid1.Rows - 1 Then
    Grid1.AddItem ""
    Grid1.TextMatrix(Grid1.Rows - 1, 0) = defBox
'    If Grid1.Rows > 2 Then Grid1.TextMatrix(Grid1.Rows - 2, 0) = Grid1.TextMatrix(Grid1.Rows - 3, 0)
End If
End Sub
Private Function MYVALID() As Boolean
If Trim(xDoc_No.Text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If

With Grid1
For i = 1 To .Rows - 2
    If .TextMatrix(i, 1) = "" Then
        .Select i, 0, i, Grid1.Cols - 1
        MsgBox "ﬂÊœ " & sName & "  €Ì— „ÊÃÊœ"
        Exit Function
    End If
    If Val(.TextMatrix(i, 4)) = 0 Then
        MsgBox "«·ﬁÌ„… €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub MyLoad()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xusername.Text = TurnValue(CardTable!UserName, Null, "")
GRDTABLE.Filter = "doc_no = " & MyParn(xDoc_No.Text)
    With Grid1
        .Rows = 1
        Do Until GRDTABLE.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = TurnValue(GRDTABLE!BOX, Null, "")
            .TextMatrix(.Rows - 1, 1) = GRDTABLE!CHARGE
            .TextMatrix(.Rows - 1, 2) = GRDTABLE!ChargeDesca & ""
            .TextMatrix(.Rows - 1, 3) = Trim(GRDTABLE!desca & "")
            .TextMatrix(.Rows - 1, 4) = Format(GRDTABLE!Value, "###0.00")
            
             GRDTABLE.MoveNext
        Loop
        .AddItem ""
        Grid1.TextMatrix(Grid1.Rows - 1, 0) = defBox
        'If Grid1.Rows > 2 Then Grid1.TextMatrix(Grid1.Rows - 1, 1) = Grid1.TextMatrix(Grid1.Rows - 2, 1)
    End With

CalcTotals
Handlecontrols LoadMode
End Sub
Private Sub myDefine()
If CardTable.EOF And CardTable.BOF Then
    xDoc_No.Text = RetZero("1", 6)
Else
    xDoc_No.Text = RetZero(Val(xDoc_No.Text) + 1, 6)
End If
xDate.Text = Format(Date, "dd-mm-yyyy")
Grid1.Rows = 1
Grid1.AddItem ""
Grid1.TextMatrix(Grid1.Rows - 1, 0) = defBox
Handlecontrols DefineMode
CalcTotals
End Sub
Private Sub Handlecontrols(nMode)
cmdNewinv.Enabled = (nMode = LoadMode And bEdit)
CmdFirst.Enabled = (nMode = LoadMode)
CmdLast.Enabled = (nMode = LoadMode)
CmdNext.Enabled = (nMode = LoadMode)
CmdDelInv.Enabled = (nMode = LoadMode) And bEdit
CmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
CmdSave.Enabled = bEdit
End Sub
Private Sub xDoc_No_LostFocus()
If Trim(xDoc_No.Text) = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad
End Sub
Private Function StrBox()
Dim BoxTable As ADODB.Recordset
Set BoxTable = New ADODB.Recordset
BoxTable.Open "SELECT * FROM file0_50 ORDER BY CODE ", CON, adOpenForwardOnly, adLockReadOnly, adCmdText
If Not (BoxTable.EOF And BoxTable.BOF) Then
    BoxTable.MoveFirst
    StrBox = "#  " & ";       "
    Do Until BoxTable.EOF
        StrBox = StrBox & "|#" & BoxTable!CODE & ";" & BoxTable!desca
        BoxTable.MoveNext
    Loop
End If
End Function
Private Sub grdLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From " & DocClient
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = True

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Private Function CalcTotals()
Dim nTotal As Double
With Grid1
For i = 1 To Grid1.Rows - 2
    nTotal = nTotal + Round(Val(Grid1.TextMatrix(i, 4)), 2)
Next
xtotal.Text = nTotal
StatusBar1.Panels(1).Text = "«·«Ã„«·Ì : " & Format(nTotal, "Fixed")
End With
End Function
Private Sub GrdDesc(nRow)
Grid1.TextMatrix(nRow, 2) = GetDesca("Select Desca From " & DocClient & " Where code = " & MyParn(Grid1.TextMatrix(nRow, 1))) & ""
'Grid1.TextMatrix(nRow, 8) = Grid1.TextMatrix(nRow, 7)
End Sub
Private Function RetDefBox() As String
Dim loctable As New ADODB.Recordset
loctable.Open "file0_50", CON, adOpenStatic, adLockReadOnly, adCmdTable
If loctable.EOF And loctable.BOF Then Exit Function
loctable.MoveLast
If loctable.RecordCount = 1 Then
    loctable.MoveFirst
    RetDefBox = Trim(loctable!CODE & "")
End If
End Function

Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.Text = "" Then Cancel = True
End Sub

