VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CreatePaidFawry 
   BackColor       =   &H00FFFFFF&
   Caption         =   "„ÿ«·»«  ›Ê—Ì ··√⁄÷«¡ «·⁄«„·Ì‰"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar Prog2 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1185
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   3390
      Begin VB.TextBox xRecords 
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
         TabIndex        =   9
         Text            =   "100"
         Top             =   630
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
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄œœ ”Ã·« "
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·„” ‰œ"
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
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   990
      End
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   555
      Left            =   2475
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Excel File"
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdApply 
      Height          =   555
      Left            =   5220
      TabIndex        =   4
      Top             =   1260
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
      Caption         =   "⁄„· „ÿ«·»«  ›Ê—Ì"
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   5
      Top             =   2310
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   6
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   2115
      Visible         =   0   'False
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin Threed.SSCommand cmddel 
      Height          =   555
      Left            =   3960
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1230
      _ExtentX        =   2170
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
      Picture         =   "CreatePaidFawry.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "CreatePaidFawry.frx":279C
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   555
      Left            =   990
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1455
      _ExtentX        =   2566
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
      Picture         =   "CreatePaidFawry.frx":4C30
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7575
      Left            =   180
      TabIndex        =   13
      Top             =   2025
      Visible         =   0   'False
      Width           =   16935
      _cx             =   29871
      _cy             =   13361
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "CreatePaidFawry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sFile As String, sFieldClose As String, sFieldDate As String, pFilter As String, sCaption As String
Public bTrans As Boolean
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
CreateFawry
End Sub
Private Sub CmdDel_Click()
If MsgBox("Õ–›", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
con.BeginTrans
On Error GoTo myerror
con.Execute "delete from file6_60"
con.Execute "delete from file6_60h"
con.CommitTrans
Inform " „ «·Õ–›"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub
Private Sub cmdExcel_Click()
myloadgrd
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con
LoadText Me
xDate1.Text = myFormat_p(Date)
Set grid1.DataSource = data10
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set CreatePaidFawry = Nothing
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
Private Sub CreateFawry(Optional bBegin As Boolean = False)
Me.MousePointer = 11
Dim loctable As New ADODB.Recordset, nRecordcount As Long

Dim cString As String, cWhere As String
cString = "SELECT " & IIf(Val(xRecords.Text) > 0, " TOP " & xRecords.Text, "") & " FILE1_10.* FROM FILE1_10 LEFT JOIN FILE6_60H ON FILE6_60H.CODE = FILE1_10.CODE WHERE FILE6_60H.CODE IS NULL AND  dbo.f_last_year_code(FILE1_10.CODE) < " & sSeason

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

nRecordcount = loctable.RecordCount

con.BeginTrans
On Error GoTo myerror
Dim i As Long
If loctable.EOF And loctable.EOF Then
    MsgBox " „ ⁄„· ﬂ· «·„ÿ«·»« "
    Exit Sub
End If

prog1.Visible = True
prog1.Value = 0
Do Until loctable.EOF
    i = i + 1
    prog1.Value = Round(i / (nRecordcount), 2) * 100
    
    Dim aSql As Variant
    aRet = addPayment(loctable!CODE, myFormat(xDate1.Text), "1", con, "FILE6_60", "FILE6_60H", MAX_YEARS, True)
    If IsEmpty(retFlag(aRet, "error")) Then
        'If Not IsEmpty(retFlag(aRet, "msg")) Then
            'MsgBox retFlag(aRet, "msg")
        'End If
        aSql = retFlag(aRet, "sql")
        If Not IsEmpty(aSql) Then
            Prog2.Visible = True
            Prog2.Value = 0
            For i2 = 0 To UBound(aSql)
                Prog2.Value = IIf(Round(i2 / (UBound(aSql)), 2) > 1, 1, mRound(i2 / (UBound(aSql)), 2)) * 100
                con.Execute aSql(i2)
            Next
            Prog2.Visible = False
        End If
    Else
        con.Execute addError(loctable!CODE, retFlag(aRet, "error"))
    End If
    loctable.MoveNext
Loop
con.CommitTrans
Me.MousePointer = 0
prog1.Visible = False
panel1(0).Caption = "⁄„· " & GetField("SELECT COUNT(*) FROM FILE6_60H WHERE ERROR = 0", con) & " „ÿ«·»…"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
prog1.Visible = False
End Sub
Private Function addError(pCode As String, pDesca As String)
Dim sDoc_no As String
sDoc_no = Newflag("FILE6_60H", "DOC_NO", con)
aInsert = AddFlag(Empty, "DOC_NO", addVal(sDoc_no))
aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate1.Text))
aInsert = AddFlag(aInsert, "[DATE_ISSUE]", addDate(xDate1.Text))
aInsert = AddFlag(aInsert, "[CODE]", addvalue(pCode))
aInsert = AddFlag(aInsert, "[TYPE]", addvalue(1))
aInsert = AddFlag(aInsert, "[YEAR_CODE]", addvalue(sSeason))
aInsert = AddFlag(aInsert, "[YEARS]", "1")
aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cUserName))
aInsert = AddFlag(aInsert, "[TIME]", "getdate()")
aInsert = AddFlag(aInsert, "[ERROR]", "1")
aInsert = AddFlag(aInsert, "[ERROR_DESCA]", addstring(pDesca))
addError = addInsert(aInsert, "FILE6_60H")
End Function
Private Sub Fixgrd()
grid1.TextMatrix(0, 0) = "Billing Account"
grid1.TextMatrix(0, 1) = "Amount"
grid1.TextMatrix(0, 2) = "Issue date"
grid1.TextMatrix(0, 3) = "Expiration Date"
grid1.TextMatrix(0, 4) = "ExtraInfoEn"
grid1.TextMatrix(0, 5) = "Extra info Arabic"
grid1.TextMatrix(0, 6) = "Hidden Info"
grid1.TextMatrix(0, 7) = "BillRefNo"
grid1.TextMatrix(0, 8) = "Key1"
grid1.TextMatrix(0, 9) = "key2"
grid1.TextMatrix(0, 10) = "key3"
grid1.TextMatrix(0, 11) = "key4"
grid1.TextMatrix(0, 12) = "key5"
grid1.ColDataType(0) = flexDTDouble
grid1.ColDataType(1) = flexDTDouble
grid1.ColDataType(2) = flexDTDate
grid1.ColWidth(5) = 2500
grid1.ColWidth(8) = 2500
End Sub
Private Sub myloadgrd()
Dim cString As String
cString = "SELECT FILE6_60H.CODE, FILE6_60H.TOTAL, CONVERT(VARCHAR(10), GETDATE(), 111) AS Expr1, NULL AS Expr2, NULL AS Expr3, SUBSTRING(FILE1_10.DESCA, 1, 32)" & _
          " AS Expr4, NULL AS Expr5, NULL AS Expr6, SUBSTRING(FILE1_10.DESCA, 1, 32) AS Expr7,NULL,NULL,NULL,NULL" & _
          "  FROM FILE6_60H  INNER JOIN FILE1_10 ON FILE6_60H.CODE = FILE1_10.CODE WHERE FILE6_60H.ERROR = 0"
Set data10.Recordset = myRecordSet(cString, con)
Fixgrd
ToFileExel2 grid1, , , , , 1, , , , 12, , Me
End Sub
