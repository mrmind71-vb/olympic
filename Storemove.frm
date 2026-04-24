VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form StoreMove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… „Œ“‰"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   2295
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   720
      Width           =   2355
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ"
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
         Height          =   270
         Left            =   1635
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   270
         Width           =   555
      End
      Begin VB.Label xBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   2760
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   1410
         Picture         =   "Storemove.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "Storemove.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   1365
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   4185
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
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   7470
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   45
      Width           =   7710
      Begin VB.TextBox xDate2 
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
         Height          =   330
         Left            =   1845
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   2310
      End
      Begin VB.TextBox xdate1 
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
         Height          =   330
         Left            =   4185
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   2310
      End
      Begin VB.TextBox xItem 
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
         Height          =   330
         Left            =   4185
         MaxLength       =   15
         TabIndex        =   0
         Top             =   180
         Width           =   2310
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   4185
         TabIndex        =   1
         Top             =   540
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   12
         Top             =   180
         Width           =   4065
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ"
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
         Height          =   270
         Left            =   6630
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   945
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰"
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
         Height          =   270
         Left            =   6630
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ «·’‰›"
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
         Height          =   270
         Left            =   6630
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -555
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   405
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6945
      Left            =   180
      TabIndex        =   11
      Top             =   1440
      Width           =   15000
      _cx             =   26458
      _cy             =   12250
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "StoreMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public aData As Variant
Dim con As New ADODB.Connection
Dim oSearch As New Search3
Sub myload()
Dim loctable As New ADODB.Recordset
cString = "select file1_11.*,file1_12.desca ,file3_10.desca as custDesca,file4_10.desca as supDesca,file0_40.desca as storeDesca" & _
          " From ((((file1_11 inner join file1_10 on file1_11.item = file1_10.item) left join file3_10 on file1_11.codecust = file3_10.code) left join file4_10 on file1_11.codesup = file4_10.code) left join file1_12 on file1_11.type = file1_12.code) left join file0_40 on file1_11.store = file0_40.code"
cString = cString & turn(cString) & " file1_11.item = " & MyParn(XITEM.Text)

If IsDate(xdate1.Text) Then
    cString = cString & turn(cString) & " file1_11.date >= " & DateSq(xdate1.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & " file1_11.date <= " & DateSq(xDate2.Text)
End If

If xStore.MatchedWithList Then
    cString = cString & turn(cString) & " file1_11.store = " & MyParn(xStore.BoundText)
End If

cString = cString & " Order by Date,[out],FILE1_12.[ORDER]"

With grid1
    .Rows = 1
    If IsDate(xdate1.Text) Then
       cString2 = "Select sum([IN] - OUT ) as Balance from file1_11 where file1_11.item = " & MyParn(XITEM.Text) & _
                  " and file1_11.date < " & DateSq(xdate1.Text) & cwhere
       nPrevious = Val(GetDesca(cString2))
       If nPrevious <> 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate1.Text
            .TextMatrix(.Rows - 1, 3) = nPrevious
       End If
    End If

    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext

    Do Until loctable.EOF
         grid1.AddItem ""
         nPrevious = nPrevious + Val(loctable!In & "") - Val(loctable!out & "")
        If loctable!Type = "F" Or loctable!Type = "T" Then
            .TextMatrix(.Rows - 1, 0) = loctable!desca & " " & loctable!StoreDesca
        Else
            .TextMatrix(.Rows - 1, 0) = loctable!desca
        End If
        .TextMatrix(.Rows - 1, 1) = Format(Val(loctable!out & ""), "#0.00")
        .TextMatrix(.Rows - 1, 2) = Format(Val(loctable!In & ""), "#0.00")
        .TextMatrix(.Rows - 1, 3) = Format(nPrevious, "#0.00")
        .TextMatrix(.Rows - 1, 4) = Format(loctable!Date, "yyyy/mm/dd")
        If loctable!Type = 3 Or loctable!Type = 6 Then
            .TextMatrix(.Rows - 1, 5) = Format(Val(loctable!price & ""), "fixed")
        End If
        .TextMatrix(.Rows - 1, 6) = loctable!doc_ID & ""
        .TextMatrix(.Rows - 1, 7) = IIf(Trim(loctable!CustDesca) & "" = "", loctable!SUPDESCA & "", loctable!CustDesca)
        .TextMatrix(.Rows - 1, 8) = loctable!Type & ""
'        If .TextMatrix(.Rows - 1, 8) = "8" Then
'            .TextMatrix(.Rows - 1, 7) = GetDesca("Select Remark from file1_82H WHERE FILE1_82h.doc_no = " & MyParn(.TextMatrix(.Rows - 1, 6)))
'        End If
        loctable.MoveNext
    Loop
End With
On Error Resume Next
XITEM.SetFocus
End Sub
Sub myProc()
If ActiveControl.Name = XITEM.Name Then
    ActiveControl.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
End If
End Sub
Function MYVALID() As Boolean
If XITEM.Text = "" Then
    MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
    Exit Function
End If
MYVALID = True
End Function
Private Sub cmdcorect_Click()

End Sub
Private Sub CmdGo_Click()
If Not MYVALID Then Exit Sub
grid1.ColHidden(3) = False
If Trim(XITEM.Text) = "" Then Exit Sub
myload
xBal.Caption = Format(Val(grid1.TextMatrix(grid1.Rows - 1, 3)), "#0.00")
On Error Resume Next
XITEM.SetFocus
End Sub
Private Sub cmdExit_Click()
Unload StoreMove
End Sub
Private Sub Form_Activate()
If XITEM.Text <> "" Then
    On Error Resume Next
    xStore.SetFocus
    Err.Clear
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    xitem_LostFocus
    If cmdGo.Enabled Then CmdGo_Click
End If
End Sub
Private Sub Form_Load()
openCon con
data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

With grid1
.TextMatrix(0, 0) = "»Ì«‰"
.TextMatrix(0, 1) = "’«œ—"
.TextMatrix(0, 2) = "Ê«—œ"
.TextMatrix(0, 3) = "—’Ìœ"
.TextMatrix(0, 4) = " «—ÌŒ"
.TextMatrix(0, 5) = "”⁄—"
.TextMatrix(0, 6) = "„” ‰œ"
.TextMatrix(0, 7) = "≈”„"


grid1.ColWidth(0) = 3000
grid1.ColWidth(1) = 1000
grid1.ColWidth(2) = 1000
grid1.ColWidth(3) = 1000
grid1.ColWidth(4) = 1500
grid1.ColWidth(5) = 1000
grid1.ColWidth(6) = 1700
grid1.ColWidth(7) = 3500
grid1.ColHidden(8) = True
End With
For I = 0 To grid1.Cols - 1
    grid1.ColAlignment(I) = flexAlignRightCenter
Next
If Not IsEmpty(aData) Then
    If IsDate(retFlag(aData, "DATE1")) Then xdate1.Text = retFlag(aData, "DATE1")
    If IsDate(retFlag(aData, "DATE2")) Then xdate1.Text = retFlag(aData, "DATE2")
    If validItem(retFlag(aData, "ITEM"), con) Then
        XITEM.Text = retFlag(aData, "ITEM")
        myload
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Unload oSearch
Err.Clear
End Sub
Private Sub grid1_dblClick()
Select Case LCase(grid1.TextMatrix(grid1.Row, 8))
Dim cDoc_no As String
Case "2", "7"
    Purchasefrm.myPublic = IIf(grid1.TextMatrix(grid1.Row, 8) = "2", 0, 1)
    Purchasefrm.sDoc_no = grid1.TextMatrix(grid1.Row, 6)
    Purchasefrm.Show
Case "3", "6"
    salesfrm.myPublic = IIf(grid1.TextMatrix(grid1.Row, 8) = "6", 0, 1)
    salesfrm.sDoc_no = grid1.TextMatrix(grid1.Row, 6)
    salesfrm.Show
End Select
End Sub
Private Sub xcustname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xCustName.Text, 7
End If
End Sub
Private Sub xdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub
Private Sub xITEM_Change()
grid1.Rows = 1
cmdGo.Enabled = Trim(XITEM.Text) <> ""
End Sub
Private Sub xITEM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub
Private Sub xItem_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ItemsLookupAll Me, oSearch
End Sub
Private Sub xitem_LostFocus()
myLostFocus XITEM
If Not cmdGo.Enabled And xStore.BoundText <> "" Then cmdGo.Enabled = True
xdesca.Caption = ""
If Trim(XITEM.Text) = "" Then Exit Sub
xdesca.Caption = GetDesca("Select Desca from file1_10 where ITEM = " & MyParn(XITEM.Text))
End Sub
Private Sub xStore_Click(Area As Integer)
If Not cmdGo.Enabled Then cmdGo.Enabled = True
End Sub
Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select File1_10.item,File1_10.Desca From file1_10 "
Generalarray(2) = "Order by file1_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(FILE1_10.ITEM LIKE 'cFilter%' or  %%DESCA%%) "


GrdArray(0, 0) = "ﬂÊœ «·’‰›"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·’‰›"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "«” ⁄·«„ «·«’‰«›"
oSearch.Show 1
End Sub
Private Sub xStore_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
On Error Resume Next
XITEM.SetFocus
End Sub
Private Sub xStore_LostFocus()
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xitem_GotFocus()
myGotFocus XITEM
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xdate1_LostFocus()
myValidDate xdate1
myLostFocus xdate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myValidDate xDate2
myLostFocus xDate2
End Sub
Private Sub LastOne_GotFocus()
myGotFocus LastOne
End Sub
Private Sub LastOne_LostFocus()
myLostFocus LastOne
End Sub
