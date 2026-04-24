VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FawryGetPamentInstallfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ãØÇáÈÇÊ ÝæÑí ááÇÚÖÇÁ ÇáãÞÓØíä"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1830
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar Prog2 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   975
      Visible         =   0   'False
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin Threed.SSCommand cmdApply 
      Height          =   555
      Left            =   5175
      TabIndex        =   0
      Top             =   225
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
      Caption         =   "ÓÍÈ ÇáÓÏÇÏ"
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   1365
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
         TabIndex        =   2
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
      TabIndex        =   3
      Top             =   1170
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
      Left            =   3690
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   225
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
      Picture         =   "GetFawaryPaymentInstall.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   870
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1725
      _cx             =   3043
      _cy             =   1535
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
Attribute VB_Name = "FawryGetPamentInstallfrm"
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
Private Sub cmdExcel_Click()
'myloadgrd
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con
Set grid1.DataSource = data10

End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set FawryGetPamentfrm = Nothing
End Sub
Private Sub CreateFawry(Optional bBegin As Boolean = False)
Me.MousePointer = 11
Dim loctable As New ADODB.Recordset, nRecordcount As Long
Dim sPath As String, i As Long, aInserts As Variant, nRecords As Long
sPath = sPath_App & "\fawry\csv_install"
aFiles = FolderFilesRet(sPath, "csv")

If IsEmpty(aFiles) Then Exit Sub
nRecordcount = UBound(aFiles) + 1

prog1.Visible = True
prog1.Value = 0
con.BeginTrans
For i = 0 To UBound(aFiles)
    prog1.Value = Round((i + 1) / (UBound(aFiles) + 1), 2) * 100
    aSql = InsertRows(sPath & "\" & aFiles(i))
    If Not IsEmpty(aSql) Then
        Prog2.Visible = True
        Prog2.Value = 0
        For i2 = 0 To UBound(aSql)
            Prog2.Value = mRound((i2 + 1) / (UBound(aSql) + 1), 2) * 100
            con.Execute aSql(i2), nRecord
            nRecords = nRecords + nRecord
        Next
        Prog2.Visible = False
    End If
Next i
con.CommitTrans
Me.MousePointer = 0
MsgBox "Êã ÓÏÇÏ " & nRecords & " ÈäÌÇÍ"
prog1.Visible = False
panel1(0).Caption = "Úãá " & nRercords & "ÓÏÇÏ"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
prog1.Visible = False
End Sub
Function FolderFilesRet(pFolder As String, sExt As String) As Variant
Dim fso As New FileSystemObject, File As File, FileCount As Long
Dim fNames As Variant
If Not fso.FolderExists(pFolder) Then
    Exit Function
End If
Set fold = fso.GetFolder(pFolder)
For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = LCase("." & sExt) And Len(File.Name) > 4 Then
        FolderFilesRet = AddFlag(FolderFilesRet, File.Name)
    End If
Next
End Function
Private Function InsertRows(cFile) As Variant
Dim FileName As Integer, TextLine As String
FileNumber = FreeFile
Open cFile For Input As #FileNumber    ' Open file.
Line Input #FileNumber, TextLine   ' Read line into variable.
Do While Not EOF(FileNumber)   ' Loop until end of file.
    Line Input #FileNumber, TextLine       ' Read line into variable.
    If TextLine <> "" Then
        cInsertRow = insertRow(TextLine)
        If cInsertRow <> "" Then
            InsertRows = AddFlag(InsertRows, cInsertRow)
        End If
    End If
Loop
Close #FileNumber    ' Close file.
Exit Function
End Function
Private Function insertRow(cString) As String
Dim aCol As Variant
If Trim(cString) = "" Then Exit Function
aCol = Split(cString, ";")
If UBound(aCol) <> 14 Then Exit Function
    
aInsert = AddFlag(Empty, "BILL_TYPE_CODE", addstring(aCol(0)))
aInsert = AddFlag(aInsert, "TRANS_NO", addstring(aCol(1)))
aInsert = AddFlag(aInsert, "TYPE_NAME", addstring(aCol(2)))
aInsert = AddFlag(aInsert, "TRANS_DATE", addDate(aCol(3) & " " & Format(aCol(4), "hh:nn")))
aInsert = AddFlag(aInsert, "BANK_CODE", addstring(aCol(5)))
aInsert = AddFlag(aInsert, "BANK_NAME", addstring(aCol(6)))
aInsert = AddFlag(aInsert, "BILL_AC_NO", addstring(aCol(7)))
aInsert = AddFlag(aInsert, "BILL_NO", addstring(aCol(8)))
aInsert = AddFlag(aInsert, "PAID_AMOUNT", mRound(aCol(9)))
aInsert = AddFlag(aInsert, "RECON_STATUS", addstring(aCol(10)))
aInsert = AddFlag(aInsert, "CHANNEL", addstring(aCol(11)))
aInsert = AddFlag(aInsert, "CHANNEL_ID", addstring(aCol(12)))
aInsert = AddFlag(aInsert, "RECEIPT_NO", addstring(aCol(13)))
aInsert = AddFlag(aInsert, "BILLER_TRANS_NO", addstring(aCol(14)))
insertRow = addInsert(aInsert, "FAWRY_TRANS_INSTALL", "dbo.[f_trans_no_found_install](" & addstring(aCol(1)) & ") = 0")
End Function

