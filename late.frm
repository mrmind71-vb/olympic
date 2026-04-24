VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Latefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«ﬁ”«ÿ «·⁄÷ÊÌ…"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   11505
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint 
      Height          =   600
      Left            =   5580
      Picture         =   "late.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   90
      Width           =   1635
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   4095
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   675
      Width           =   7305
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "«ﬁ”«ÿ €Ì— „”Ã·…"
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
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   945
         Width           =   1335
      End
      Begin VB.Label xLate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   900
         Width           =   1680
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "≈Ã„«·Ì «·«ﬁ”«ÿ"
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
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   900
         Width           =   1680
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
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
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   585
         Width           =   510
      End
      Begin VB.Label xCode_Desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   4020
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄÷Ê"
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
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   480
      End
      Begin VB.Label xcode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   540
         Width           =   1680
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6630
      Left            =   90
      TabIndex        =   0
      Top             =   2025
      Width           =   11355
      _cx             =   20029
      _cy             =   11695
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   8685
      Width           =   11400
      Begin VB.Label xDiffer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label7 
         Caption         =   "«·»«ﬁÌ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label xTotal_check 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   8460
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "≈Ã„«·Ì «·«ﬁ”«ÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9945
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   900
      Width           =   1275
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "late.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "late.frx":49A3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   7245
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   4155
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "late.frx":6D06
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   510
         Left            =   1395
         MaskColor       =   &H00FFFFFF&
         Picture         =   "late.frx":9124
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   510
         Left            =   2745
         MaskColor       =   &H00FFFFFF&
         Picture         =   "late.frx":B9BE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   465
      Left            =   1800
      Top             =   -225
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
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
      Height          =   465
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
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
End
Attribute VB_Name = "Latefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
Dim con As New ADODB.Connection
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From FILE6_21 WHERE CODE = " & addvalue(xcode.Caption)
    con.CommitTrans
    myloadGrd
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdNewInv_Click()
Set addLatefrm.myForm = Me
addLatefrm.sDoc_no = xdoc_no.Caption
addLatefrm.sDate = xDate.Caption
addLatefrm.nLate = Val(xDiffer.Caption)
addLatefrm.Show 1
CalcTotals
End Sub
Private Sub CmdPrint_Click()
Dim cHeader1 As String
Dim aHeader As Variant
cHeader1 = "»Ì«‰  ›’Ì·Ì «ﬁ”«ÿ «·⁄÷Ê" & xCode_Desca.Caption
If IsDate(xDate.Caption) Then aHeader = AddFlag(aHeader, "» «—ÌŒ : " & xDate.Caption)
aHeader = AddFlag(aHeader, "«·«Ã„«·Ì : " & xTotal_check.Caption & " Ã‰ÌÂ")
PrintGrdNew.doprint grid1, 0.9, -4, cHeader1, retHeader(aHeader, 0, 1), retHeader(aHeader, 1, 2), , False, False, 9
PrintGrdNew.Show 1
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
myloadGrd
End Sub
Private Sub Form_Activate()
If Not bActivated Then
    bActivated = True
    On Error Resume Next
    If xdoc_no.Tag = LoadMode Then
        grid1.SetFocus
        Err.Clear
    End If
End If
End Sub
Private Sub Form_Load()
bEdit = True
openCon con
Set grid1.DataSource = DATA1
myload
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Set Latefrm = Nothing
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
CalcTotals
If mRound(xDiffer.Caption) <> 0 Then
    MsgBox "≈Ã„«·Ì «·«ﬁ”«ÿ ·«  ”«ÊÏ «Ã„«·Ì «·›« Ê—…"
    Exit Function
End If

With grid1
For I = 1 To .rows - 2
    If Not IsDate(.TextMatrix(I, 1)) Then
        .Select I, 0, I, grid1.Cols - 1
        MsgBox " «—ÌŒ «·«” Õﬁ«ﬁ €Ì— „”Ã·"
        Exit Function
    End If
    If mRound(.TextMatrix(I, 2)) = 0 Then
        MsgBox "ﬁÌ„… «·ﬁ”ÿ €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
Dim loctable As New ADODB.Recordset, cString As String
cString = "SELECT FILE2_10.CODE,FILE2_10.DESCA,dbo.mem_late(FILE2_10.CODE) AS LATE" & _
          " FROM FILE2_10"
cString = cString & turn(cString) & "FILE2_10.CODE = " & addvalue(xcode.Caption)
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (loctable.EOF And loctable.BOF) Then
    xcode.Caption = loctable!CODE & ""
    xDate.Caption = myFormat(loctable!Date)
    xCode_Desca.Caption = loctable!CODE_DESCA & ""
    xTotal.Caption = Myvalue(loctable!total)
    xLate.Caption = Myvalue(loctable!late)
End If
loctable.Close
Set loctable = Nothing
myloadGrd
CellPos 13, grid1.rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 46 And grid1.Row <> grid1.rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbDefaultButton2 + vbOKCancel) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete FROM FILE6_21 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        myloadGrd
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 2 Then
    If IsDate(.EditText) Then
        .EditText = Format(.EditText, "YYYY/M/D")
    Else
        Cancel = True
    End If
End If
End With
End Sub
Private Function CalcTotals()
Dim nTotal_inv As Double, nTotal_Check As Double
For I = 1 To grid1.rows - 2
    nTotal_Check = nTotal_Check + mRound(grid1.TextMatrix(I, 3))
Next
xTotal_check.Caption = mRound(nTotal_Check, 2)
xDiffer.Caption = mRound(xLate.Caption) - mRound(nTotal_Check)
End Function
Private Sub myloadGrd()
With grid1
    Dim cString As String
    cString = "SELECT CONVERT(VARCHAR(10),FILE6_21.DATE_DUE,111),FILE6_21.[VALUE],dbo.install_late(FILE6_21.ID)  ,'',FILE6_21.ID " & _
               " FROM FILE6_21"
    cString = cString & turn(cString) & "FILE6_21.DOC_NO = " & MyParn(xdoc_no.Caption)
    Set DATA1.Recordset = myRecordSet(cString, con)
    myAddItem
End With
CalcTotals
Fixgrd
End Sub
Private Sub Fixgrd()
Dim I As Long
With grid1
.FormatString = "«·„”·”·|" & "«· «—ÌŒ|" & " «·ﬁÌ„…|" & "«·„”œœ|" & " «—ÌŒ «·”œ«œ|"
.ColWidth(0) = 800
.ColWidth(1) = 1000
.ColWidth(2) = 1500
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.ColWidth(5) = 1500
.ColDataType(2) = flexDTDecimal
.ColDataType(3) = flexDTDecimal
.ColHidden(1) = True
.ColHidden(.Cols - 1) = True

FixSerial

For I = 1 To grid1.Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then
    CalcTotals
    Exit Sub
End If
With grid1
If Row = grid1.rows - 1 Then myAddItem
CalcTotals
'If myreplace(Row) Then
'    myloadgrd
'End If
End With
End Sub
Private Sub grid1_EnterCell()
If grid1.Col = 1 Or grid1.Col = 2 Or grid1.Col = 3 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Grid1_GotFocus()
grid1_EnterCell
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow(OldRow) Then myRemove OldRow
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then myRemove grid1.Row
End Sub
Private Function validRow(Row As Long) As Boolean
With grid1
If Not IsDate(.TextMatrix(Row, 2)) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 3)) Then Exit Function
End With
validRow = True
End Function
Private Sub myAddItem()
With grid1
.AddItem ""
MakeSerial
If grid1.rows > 2 Then
    If IsDate(.TextMatrix(.rows - 2, 2)) Then
        .TextMatrix(.rows - 1, 2) = Format(DateAdd("m", 1, .TextMatrix(.rows - 2, 2)), "yyyy/mm/dd")
    End If
End If
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 - IIf(Not IsDate(grid1.TextMatrix(Row, 5)), 1, 0) Then
    grid1.Col = Col + 1
ElseIf Row < grid1.rows - 1 Then
    Dim nCol As Long
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 2, 3)
    grid1.ShowCell grid1.Row, 3
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
MakeSerial Row
CalcTotals
End Sub
Private Sub MakeSerial(Optional nBeginRow As Long = 1)
For I = 1 To grid1.rows - 1
    grid1.TextMatrix(I, 0) = I
Next
End Sub
Private Function myreplaceGrd(Row As Long) As Boolean
Dim aInsert As Variant
With grid1
    For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.rows - 2, Row)
        aInsert = AddFlag(Empty, "CODE", addvalue(xcode.Caption))
        aInsert = AddFlag(aInsert, "DATE_DUE", addDate(grid1.TextMatrix(I, 2)))
        aInsert = AddFlag(aInsert, "[VALUE]", Val(grid1.TextMatrix(I, 3)))
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            aInsert = AddFlag(aInsert, "ID", grid1.TextMatrix(I, .Cols - 1))
            con.Execute addInsert(aInsert, "FILE6_21")
        Else
            con.Execute addUpdate(aInsert, "FILE6_21", "ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
    myreplaceGrd = True
End With
End Function
Private Function MyReplace(Optional Row As Long = -1, Optional Row2 As Long = -1) As Boolean
con.BeginTrans
myreplaceGrd Row
con.CommitTrans
MyReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub FixSerial()
Dim I As Long
For I = 1 To grid1.rows - 1
    grid1.TextMatrix(I, 0) = I
Next
End Sub
