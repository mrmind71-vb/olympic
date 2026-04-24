VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CreatePaidFawry_install 
   BackColor       =   &H00FFFFFF&
   Caption         =   "⁄„· „ÿ«·»«  ›Ê—Ì ··⁄÷ÊÌ… «·„ﬁ”ÿ…"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   45
      Width           =   4470
      Begin VB.TextBox xInstall_count 
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
         Height          =   345
         Left            =   495
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1455
      End
      Begin Threed.SSCommand cmdInstall_type 
         Height          =   330
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
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
         Caption         =   "..."
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄œœ «ﬁ”«ÿ „ √Œ—… ⁄·Ì «·⁄÷Ê"
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
         Index           =   2
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   2250
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   4470
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
         TabIndex        =   2
         Top             =   225
         Width           =   1860
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
         TabIndex        =   6
         Top             =   270
         Width           =   990
      End
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   555
      Left            =   5130
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1485
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   7
      Top             =   2355
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
         TabIndex        =   8
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
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   555
      Left            =   3645
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1485
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
      Picture         =   "CreatePaidFawry_install.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   1140
      Left            =   -270
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   12030
      _cx             =   21220
      _cy             =   2011
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
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
End
Attribute VB_Name = "CreatePaidFawry_install"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sFile As String, sFieldClose As String, sFieldDate As String, pFilter As String, sCaption As String
Public bTrans As Boolean
Dim con As New ADODB.Connection
Private Sub cmdApply_Click()
'CreateFawry
End Sub
Private Sub cmddel_Click()
If MsgBox("Õ–›", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
con.BeginTrans
On Error GoTo myerror
con.Execute "delete from file6_60"
con.Execute "delete from file6_60H"
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

Private Sub cmdInstall_type_Click()
install_typefrm.Show 1
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
Dim cString As String, cWhere As String
cString = "SELECT FILE2_10.CODE,SUM(INSTALL_BALANCE.TOTAL - INSTALL_BALANCE.TOTAL_PAID), CONVERT(VARCHAR(10), GETDATE(), 111) AS Expr1, NULL AS Expr2, NULL AS Expr3, SUBSTRING(FILE2_10.DESCA, 1, 32)" & _
          " AS Expr4, NULL AS Expr5, NULL AS Expr6, SUBSTRING(FILE2_10.DESCA, 1, 32) AS Expr7,NULL,NULL,NULL,NULL" & _
          " From File2_10 INNER JOIN INSTALL_BALANCE ON FILE2_10.CODE = INSTALL_BALANCE.CODE LEFT JOIN INSTALL_CODES ON FILE2_10.INSTALL_TYPE = INSTALL_CODES.CODE "
cWhere = "INSTALL_BALANCE.TOTAL - INSTALL_BALANCE.TOTAL_PAID  > 0 AND  FILE2_10.STATUS = 1"
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " GROUP BY FILE2_10.CODE,FILE2_10.DESCA,INSTALL_CODES.INSTALL_COUNT"
If Val(xInstall_count.Text) > 0 Then
    cHaving = "Sum(INSTALL_BALANCE.INS_COUNT) <= " & mRound(xInstall_count.Text)
ElseIf Val(xInstall_count.Text) = 0 Then
    cHaving = "(Sum(INSTALL_BALANCE.INS_COUNT) <= INSTALL_CODES.INSTALL_COUNT OR INSTALL_CODES.INSTALL_COUNT = 0) "
End If
If cHaving <> "" Then cString = cString & " Having " & cHaving
Set data10.Recordset = myRecordSet(cString, con)
Fixgrd
ToFileExel2 grid1, , , , , 1, , , , 12, , Me
End Sub

Private Sub xInstall_count_GotFocus()
myGotFocus xInstall_count
End Sub
Private Sub xInstall_count_LostFocus()
myLostFocus xInstall_count
End Sub
