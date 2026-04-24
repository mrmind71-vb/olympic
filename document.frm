VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form add_photos 
   Caption         =   "‘Ìﬂ«  «·«⁄÷«¡ «·„”Ã·Ì‰ »«·‘—ﬂ« "
   ClientHeight    =   10290
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ImgXCtrl6.ImgXCtrl ImgX1 
      Height          =   8565
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   15108
      BackColor       =   16777215
      BorderStyle     =   4
      AutoZoom        =   -1  'True
      SelectionLineType=   4
      Center          =   -1  'True
      ImageBorderThickness=   1
      DoubleBuffer    =   -1  'True
      LicenseUserName =   "mrmind"
      LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   555
      Left            =   1575
      Picture         =   "document.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8730
      Width           =   1365
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   135
      MaskColor       =   &H00FFFFFF&
      Picture         =   "document.frx":242A
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   8730
      UseMaskColor    =   -1  'True
      Width           =   1410
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   -2175
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   4140
      Top             =   10530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8610
      Left            =   9720
      TabIndex        =   0
      Top             =   45
      Width           =   10455
      _cx             =   18441
      _cy             =   15187
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
      BackColorFixed  =   12648384
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
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
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   6840
      Top             =   8595
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin ImgXCtrl6.ImgXCtrl imgx2 
      Height          =   8565
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   15108
      BackColor       =   16777215
      BorderStyle     =   4
      AutoZoom        =   -1  'True
      SelectionLineType=   4
      Center          =   -1  'True
      ImageBorderThickness=   1
      DoubleBuffer    =   -1  'True
      LicenseUserName =   "mrmind"
      LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
   End
   Begin VB.Label xPhoto 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   9675
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   8775
      Visible         =   0   'False
      Width           =   2040
   End
End
Attribute VB_Name = "add_photos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents twain As ImgXTwain
Attribute twain.VB_VarHelpID = -1
Public bedit As Boolean
Private WithEvents MyPrinter As ImgXPrint
Attribute MyPrinter.VB_VarHelpID = -1
Dim con As New ADODB.Connection
Public cCode As String
Dim fs As New FileSystemObject
Dim nCurrent As Integer, nOption As Integer
Dim aPhoto
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set ImgXTwain = Nothing
closeCon con
Set Document = Nothing
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Not validRow(Row) Then Exit Sub
If Row = .rows - 1 Then
   MyAddItem
End If
If Not myreplace(Row) Then
    myLoadGrd
ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
    myLoadGrd
End If
End With
End Sub
Private Sub Form_Load()
openCon con
Set grid1.DataSource = DATA1
myLoadGrd
CellPos 13, grid1.rows - 2, grid1.Cols - 1
End Sub
Private Sub cmdReplace_Click()
End Sub
Private Sub ScanImage()
On Error GoTo myerror
Set twain = New ImgXTwain
twain.OpenTwain Me.hwnd
If twain.QuerySupport(ixtcResolution) Then
     twain.Resolution = 150
End If
'twain.Acquire Check1.Value = 1, Me.hWnd
twain.Acquire False, Me.hwnd
Exit Sub
myerror:
MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub
Private Sub Twain_ImageAcquired(Image As ImgX_Image)
ReplaceFromImage Image
End Sub
Private Sub ReplaceFromImage(Image As ImgX_Image)
On Error GoTo myerror
cSource = Checks_Files(cCode, xPhoto.Caption)
MyCreateFolder Checks_Dir(cCode)
imgx1.Images.Replace Image, , False
imgx1.Refresh
imgx1.Export.ToFile cSource, ixfsJPG
LoadPhoto
Exit Sub
myerror:
imgx1.Images.Clear
Err.Clear
End Sub
Private Sub cmdPrint_Click()
Dim i As Long
Dim Index As Integer
Set MyPrinter = New ImgXPrint
MyPrinter.PageFrom = 1
MyPrinter.PageTo = grid1.rows - 2
MyPrinter.PageMax = grid1.rows - 2
MyPrinter.MarginLeft = 0
MyPrinter.MarginRight = 0
MyPrinter.MarginTop = 0
MyPrinter.MarginBottom = 15
MyPrinter.PageMin = 1
MyPrinter.Antialias = True
On Error GoTo myerror
If MyPrinter.ShowPrinter(Me.hwnd) Then
    If MyPrinter.Range = iprAllPages And grid1.rows - 1 > 0 Then
        For i = 1 To grid1.rows - 2
            If validPhoto(Checks_Files(cCode, cCode & "-" & grid1.TextMatrix(i, grid1.Cols - 1))) Then
                imgx2.Import.FromFile Checks_Files(cCode, cCode & "-" & grid1.TextMatrix(i, grid1.Cols - 1))
            End If
            MyPrinter.PrintImage "Print Document", imgx2.Images(0), False, True
        Next
    ElseIf MyPrinter.Range = ixprSelection Then
        MyPrinter.PrintImages "Print Document", imgx2.Images, False, True
    Else
        For i = MyPrinter.PageFrom To MyPrinter.PageTo
            If validPhoto(Checks_Files(cCode, cCode & "-" & grid1.TextMatrix(i, grid1.Cols - 1))) Then
                imgx2.Import.FromFile Checks_Files(cCode, cCode & "-" & grid1.TextMatrix(i, grid1.Cols - 1))
            End If
            MyPrinter.PrintImage "Print Document", imgx2.Images(0), False, True
        Next
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub LoadPhoto()
On Error GoTo myerror
imgx1.Images.Clear
If Not MyCreateFolder(Checks_Dir(cCode)) Then Exit Sub
If xPhoto.Caption <> "" Then
    imgx1.Import.FromFile Checks_Files(cCode, xPhoto.Caption)
End If
Exit Sub
myerror:
Err.Clear
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "„”·”·|«·»Ì«‰|„”Õ|’Ê—…|Õ–›"
.ColWidth(0) = 700
.ColWidth(1) = 7000
.ColWidth(2) = 800
.ColWidth(3) = 800
.ColWidth(4) = 800
.ColComboList(2) = "..."
.ColComboList(3) = "..."
.ColComboList(4) = "..."
.ColHidden(.Cols - 1) = True
For i = 0 To grid1.Cols - 1
    grid1.ColAlignment(i) = flexAlignRightCenter
Next
MakeSerial
End With
End Sub
Private Sub Handlecontrols()
'Me.cmdDelAll.Enabled = bedit
'Me.cmdAddReplace.Enabled = bedit
'Me.cmdDelCur.Enabled = bedit
'Me.cmdAddPhoto.Enabled = bedit
End Sub
Private Function myCut(pString) As String
Dim aLocal As Variant
aLocal = Split(pString, "\")
For i = 0 To UBound(aLocal) - 2
    myCut = myCut & turn(myCut, "\") & aLocal(i)
Next
End Function
Private Function myreplace(Optional Row As Long = -1)
Dim aInsert As Variant
With grid1
    con.BeginTrans
    On Error GoTo myerror
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.rows - 2, Row)
        aInsert = AddFlag(Empty, "DESCA", addstring(.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "CODE", addstring(cCode))
        aInsert = AddFlag(aInsert, "row", i)
        If grid1.TextMatrix(i, .Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE6_90_D")
        Else
            con.Execute addUpdate(aInsert, "FILE6_90_D", "ID = " & .TextMatrix(i, .Cols - 1))
        End If
    Next
    con.CommitTrans
End With
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myLoadGrd()
Dim cString As String
cString = "SELECT DESCA as [«·«”„],NULL ,NULL,NULL,FILE6_90_D.ID" & _
          " FROM FILE6_90_D"
cString = cString & turn(cString) & "FILE6_90_D.CODE = " & MyParn(cCode)
cString = cString & " order by FILE6_90_D.ROW"
Set DATA1.Recordset = myRecordSet(cString, con)
grid1_EnterCell
MyAddItem
Fixgrd
End Sub
Private Function MyAddItem()
grid1.AddItem ""
End Function

Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .rows - 1 And OldRow <> 0 Then
    If Not validRow(OldRow) Then .RemoveItem OldRow
End If
End With
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
Set fs = CreateObject("Scripting.FileSystemObject")
If Col = 2 And xPhoto.Caption <> "" Then
    nOption = 2
    ScanImage
    
'    camerafrm.sPicture = Checks_Files(cCode, xPhoto.Caption)
'    camerafrm.Show 1
    LoadPhoto
ElseIf Col = 3 And xPhoto.Caption <> "" Then
    MyCreateFolder Checks_Dir(cCode)
    Dim cFile As String, cNewFile As String
    Common1.FileName = ""
    Common1.InitDir = Checks_Dir(cCode)
    Common1.Filter = "Pictures (*.Jpg)|*.Jpg"
    Common1.ShowOpen
    If Common1.FileTitle <> "" Then
        cFile = Common1.FileName
        If cFile <> "" Then
            cNewFile = Checks_Files(cCode, xPhoto.Caption)
            fs.CopyFile cFile, cNewFile
        End If
        LoadPhoto
    End If
ElseIf Col = 4 Then
    If MsgBox("Õ–› «·’Ê—…", vbOKCancel + vbDefaultButton2) Then
        fs.DeleteFile Checks_Files(cCode, xPhoto.Caption)
        LoadPhoto
    End If
End If
Exit Sub
myerror:
        MsgBox Err.Description
        Err.Clear
End Sub
Private Sub grid1_EnterCell()
If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    xPhoto.Caption = ""
Else
    xPhoto.Caption = cCode & "-" & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
End If
End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
If keyscii = 13 And grid1.Col <> 2 And grid1.Col <> 3 Then KeyAscii = 0
End Sub

Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If (Not validRow(.Row)) And .Row <> .rows - 1 And .Row <> 0 Then
    .RemoveItem .Row
End If
End With
End Sub
Private Function validRow(Row As Long) As Boolean
With grid1
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Trim(grid1.EditText) = "" Then
    Cancel = True
End If
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub xPhoto_Change()
LoadPhoto
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 46 Then
    If Trim(grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) <> "" Then
        If MsgBox("«·€«¡ «·”Ã· ?? Â· «‰  „ √ﬂœ", vbYesNo + vbDefaultButton2) = vbYes Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "Delete  from FILE6_90_D  where id = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            If fs.FileExists(Checks_Files(cCode, cCode & "-" & grid1.TextMatrix(grid1.Row, grid1.Cols - 1))) Then
                fs.DeleteFile Checks_Files(cCode, cCode & "-" & grid1.TextMatrix(grid1.Row, grid1.Cols - 1))
            End If
            con.CommitTrans
            myreplace
            myLoadGrd
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
myLoadGrd
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 Then
    grid1.Select Row, Col + 1
ElseIf Row < grid1.rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 1)
    grid1.ShowCell Row + 1, 1
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With grid1
If .rows = 0 Then Exit Sub
If Trim(grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) = "" Then Exit Sub
R = .Row
R = .DragRow(R)
If R <> grid1.rows - 1 Then
    myreplace
End If
myLoadGrd
End With
End Sub
Private Sub Twain_CanCloseTwain()
    ' This event is called after you call Acquire.
    ' It let's you know when it's safe to call CloseTwain.
    twain.CloseTwain
    ' Steps menu
End Sub

