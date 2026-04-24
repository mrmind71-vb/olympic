VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form itemsGrdfrm 
   Caption         =   "»Ì«‰«  «·«’‰«›"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10650
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   510
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   9450
      Width           =   2310
   End
   Begin VB.Frame Frame7 
      Height          =   915
      Left            =   15300
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   4410
      Width           =   3660
      Begin VB.CommandButton cmdTrans 
         Caption         =   "‰ﬁ· «·Ì ›« Ê—… „‘ —Ì« "
         Height          =   645
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   180
         Width           =   3435
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1680
      Left            =   15255
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2700
      Width           =   3750
      Begin VB.CommandButton cmdExel 
         Height          =   645
         Left            =   90
         Picture         =   "itemsgrdfrm.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "⁄—÷"
         Top             =   900
         Width           =   3570
      End
      Begin Threed.SSCommand cmdBarCode 
         Height          =   690
         Left            =   90
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   180
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   1217
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "itemsgrdfrm.frx":27EB
         Caption         =   " ÕÊÌ· ··»«—ﬂÊœ"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2670
      Left            =   15255
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   0
      Width           =   3750
      Begin VB.Image Image1 
         Height          =   2400
         Left            =   90
         Stretch         =   -1  'True
         Top             =   180
         Width           =   3570
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   375
      Left            =   -720
      TabIndex        =   19
      Top             =   6210
      Visible         =   0   'False
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame4 
      Height          =   780
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   8685
      Width           =   2130
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "itemsgrdfrm.frx":5184
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1950
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   8685
      Width           =   2490
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1530
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1500
      Left            =   12555
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   8640
      Width           =   2625
      Begin VB.CommandButton Command4 
         Caption         =   " ⁄œÌ· «·„Ã„Ê⁄… «·—∆Ì”Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1035
         Width           =   2445
      End
      Begin VB.CommandButton Command3 
         Caption         =   " ⁄œÌ· «·„Ã„Ê⁄… "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   630
         Width           =   2445
      End
      Begin VB.CommandButton Command1 
         Caption         =   " ⁄œÌ· «·√ﬁ”«„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   2445
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   4725
      TabIndex        =   5
      Top             =   8640
      Width           =   7800
      Begin VB.TextBox XITEM 
         Appearance      =   0  'Flat
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   2445
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   3825
         TabIndex        =   2
         Top             =   180
         Width           =   2445
         _ExtentX        =   4313
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
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   3825
         TabIndex        =   1
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… «·’‰› :"
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
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂÊœ :"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁ”„ :"
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
         Left            =   6330
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "≈”„ :"
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
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   555
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2295
      Top             =   315
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   4455
      Top             =   1575
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   2475
      Top             =   1125
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   270
      Top             =   1350
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   405
      TabIndex        =   18
      Top             =   1710
      Visible         =   0   'False
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8565
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   15135
      _cx             =   26696
      _cy             =   15108
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
      BackColorSel    =   8454143
      ForeColorSel    =   128
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
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
Attribute VB_Name = "itemsgrdfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public aPublic, bedit As Boolean
Dim clist1 As String, cList2 As String, cList3 As String, cList4 As String
Dim CardTable As New adodb.Recordset
Dim con As New adodb.Connection
Private Sub myload()
Dim cFilter As String
On Error GoTo myerror
Dim cField1 As String
'                   0                   1                           2                       3                       4                           5                   6                       7
cString = "SELECT ITEM as [«·ﬂÊœ],FILE1_10.DESCA as [«·»Ì«‰],FILE1_10.[GROUP] as [«·„Ã„Ê⁄…],[SECTION] as [«·ﬁ”„],FILE1_10.SUPLER  as [«·„Ê—œ] ,COST AS [”⁄— «· ﬂ·›…],PRICE AS [”⁄— „” Â·ﬂ] , BALANCE AS [«·ﬂ„Ì…],FILE1_10.[SHOW] AS [≈ŸÂ«—] ,ID_ITEM " & _
         " FROM FILE1_10 LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE"

If IsNumeric(xgroup.BoundText) Then
    cString = cString & turn(cString) & " FILE1_10.[GROUP] = " & xgroup.BoundText
End If

If IsNumeric(xSection.BoundText) Then
    cString = cString & turn(cString) & " FILE1_10.[SECTION] = " & xSection.BoundText
End If

If Trim(xDescA.Text) <> "" Then
    cString = cString & turn(cString) & MyParnAnd(xDescA.Text, "FILE1_10.desca")
End If

If Trim(xItem.Text) <> "" Then
    cString = cString & turn(cString) & MyParnAnd(xItem.Text, "ITEM")
End If

cString = cString & " order by FILE1_10.ID_ITEM "
data1.RecordSource = cString
data1.Refresh
myaddItem
lblTotal.Caption = IIf(grid1.Rows < 3, "", "≈Ã„«·Ì ⁄œœ «·«’‰«› : " & grid1.Rows - 2)
Fixgrd
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdBarCode_Click()
If Val(grid1.TextMatrix(grid1.row, 7)) = 0 Then
    MsgBox "·«  ÊÃœ ﬂ„Ì… ··’‰›"
    Exit Sub
End If

Dim cString As String
    con.BeginTrans
    On Error GoTo myerror:
    cString = "INSERT INTO ADDPRINT(ITEM,QUANT,ISPRINT) " & _
               "VALUES(" & grid1.TextMatrix(grid1.row, 0) & "," & _
                            grid1.TextMatrix(grid1.row, 7) & "," & _
                            "1" & _
               ")"
    con.Execute cString
    con.CommitTrans
barcodefrm.Show 1
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub cmdExel_Click()
grid1.ColHidden(2) = True
grid1.ColHidden(3) = True
grid1.ColHidden(4) = True
ToFileExel grid1
grid1.ColHidden(2) = False
grid1.ColHidden(3) = False
grid1.ColHidden(4) = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CMDTRANS_Click()
If MsgBox("‰ﬁ· «·Ì ›« Ê—… „‘ —Ì« ", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
If GetDesca("Select doc_no from FILE7_20 WHERE DOC_NO = '000001'") <> "" Then
    MsgBox "«·›« Ê—… „ÊÃÊœ…"
    Exit Sub
End If

Dim loctable As New adodb.Recordset, cString As String
cString = "Select * from file1_10"
cString = cString & turn(cString) & " file1_10.BALANCE <> 0"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert()
ReDim aInsert(7, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring("000001")

aInsert(1, 0) = "code"
aInsert(1, 1) = addstring("000001")

aInsert(2, 0) = "[Date]"
aInsert(2, 1) = addDate(Format(Date, "dd-mm-yyyy"))

aInsert(3, 0) = "store"
aInsert(3, 1) = addstring("01")

aInsert(4, 0) = "Discount"
aInsert(4, 1) = 0

aInsert(5, 0) = "Tax"
aInsert(5, 1) = 0
con.Execute CreateInsert(aInsert, "File7_20h")
    
    

Dim i As Long
ReDim aInsert(4, 1)
Do Until loctable.EOF
    i = i + 1
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring("000001")
    
    aInsert(1, 0) = "item"
    aInsert(1, 1) = addstring(loctable!Item)
    
    aInsert(2, 0) = "quant"
    aInsert(2, 1) = Val(loctable!balance)

    aInsert(3, 0) = "Price"
    aInsert(3, 1) = Val(loctable!cost)

    aInsert(4, 0) = "row"
    aInsert(4, 1) = i
    con.Execute CreateInsert(aInsert, "File7_20")
    loctable.MoveNext
Loop
Inform " „ «÷«›… ›« Ê—… «·„‘ —Ì«  »‰Ã«Õ"
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub Command1_Click()
ReDim aPublic(5)
aPublic(0) = "FILE1_10SC"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "«·ﬂÊœ"
aPublic(4) = "«·»Ì«‰"
aPublic(5) = "√ﬁ”«„ «·«’‰«›"
FlagFrm.nMin = 0
FlagFrm.bedit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
cList2 = StrList2("select * from file1_10SC order by desca")
grid1.ColComboList(3) = cList2
DATA3.Refresh
End Sub

Private Sub Command2_Click()
'If grid1.Rows = 1 Then
'    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»⁄Â«"
'    Exit Sub
'End If
'
'If Not doprint Then
'    MsgBox "·«  ÊÃœ ”Ã·«  ··ÿ»«⁄…"
'    Exit Sub
'End If
'CardPrintNew1.PrintArray
'CardPrintNew1.Show 1
getItems
End Sub

Private Sub Command3_Click()
ReDim aPublic(5)
itemsGroupFrm.bedit = True
itemsGroupFrm.Show 1
clist1 = StrList2("select * from file1_50 order by desca")
grid1.ColComboList(2) = clist1
data2.Refresh
End Sub

Private Sub Command4_Click()
ReDim aPublic(5)
aPublic(0) = "FILE1_50G"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "«·ﬂÊœ"
aPublic(4) = "«·»Ì«‰"
aPublic(5) = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
FlagFrm.bedit = True
FlagFrm.nMin = 0
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set itemsgrdfrm = Nothing
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal row As Long, ByVal col As Long)
Dim nCode As Integer
If Not validRow(row) Then Exit Sub
If grid1.row = grid1.Rows - 1 Then
    myaddItem
End If
 
If Not myreplace(row) Then
    If grid1.TextMatrix(row, grid1.Cols - 1) = "" Then myload
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
myload
End Sub
Private Sub grid1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 Then
    If Not validRow(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub
Private Sub grid1_EnterCell()
With grid1
    LoadPhoto grid1.TextMatrix(grid1.row, 0)
    If (grid1.col = 0 And grid1.TextMatrix(grid1.row, 0) <> "") Then
        grid1.Editable = flexEDNone
    Else
        grid1.Editable = flexEDKbdMouse
    End If
End With
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cCode As String
On Error GoTo myerror
If KeyCode = 116 And grid1.row = grid1.Rows - 1 And grid1.col = 0 Then
    nCode = Val(GetDesca("Select Max(item2) from file1_10 ")) + 1
    If nCode < 1001 Then nCode = 1001
    grid1.TextMatrix(grid1.row, 0) = nCode
    grid1_AfterEdit grid1.row, grid1.col
End If
If KeyCode = 46 And grid1.row <> 0 And grid1.row <> grid1.Rows - 1 Then
    If Trim(grid1.TextMatrix(grid1.row, 0)) <> "" Then
        If MsgBox("Õ–› «·’‰› !! „Ê«›ﬁ", vbOKCancel + vbDefaultButton2) = vbOK Then
            con.BeginTrans
            con.Execute "delete from file1_10 where item = " & MyParn(grid1.TextMatrix(grid1.row, 0))
            con.CommitTrans
            grid1.RemoveItem grid1.row
        End If
    End If
End If
Exit Sub
myerror:
If Err.Number <> 0 Then MsgBox Err.Description
con.RollbackTrans
myload
End Sub
Private Sub Form_Load()
bedit = True
openCon con

data2.ConnectionString = strCon
data2.RecordSource = "FILE1_50"
Set xgroup.RowSource = data2
xgroup.ListField = "Desca"
xgroup.BoundColumn = "Code"

DATA3.ConnectionString = strCon
DATA3.RecordSource = "FILE1_10SC"
Set xSection.RowSource = DATA3
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

data4.ConnectionString = strCon
data4.RecordSource = "FILE1_50G"

Set grid1.DataSource = data1
data1.ConnectionString = strCon
With grid1
clist1 = StrList2("Select code,desca from file1_50 order by desca")
cList2 = StrList2("Select code,desca from file1_10sc order by desca")
cList3 = StrList2("Select code,desca from FILE4_10 order by desca")
myload
grid1.row = grid1.Rows - 1
grid1.ShowCell grid1.Rows - 1, 0
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
If Not validRow(grid1.row) And grid1.row <> grid1.Rows - 1 And grid1.row <> 0 Then grid1.RemoveItem OldRow
End Sub
Private Sub Grid1_ValidateEdit(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
If col = 0 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "ﬂÊœ «·’‰› „ÿ·Ê»"
        Cancel = True
        Exit Sub
    End If
End If
If col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "Ê’› «·’‰› „ÿ·Ê»"
        Cancel = True
    End If
End If
If col = 2 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "„Ã„Ê⁄… «·’‰› „ÿ·Ê»…"
        Cancel = True
    End If
End If

Exit Sub
myerror:
On Error Resume Next
If Err.Number <> 0 Then MsgBox Err.Description
CardTable.CancelUpdate
con.RollbackTrans
myload
Err.Clear
End Sub
Private Sub Fixgrd()
With grid1
.ColComboList(2) = clist1
.ColComboList(3) = cList2
.ColComboList(4) = cList3
.ColWidth(0) = 1800
.ColWidth(1) = 4000
.ColWidth(2) = 1300
.ColWidth(3) = 1300
.ColWidth(4) = 1300
.ColWidth(5) = 900
.ColWidth(6) = 900
.ColWidth(7) = 900
.ColHidden(.Cols - 1) = True
.RowHeight(0) = 800
.WordWrap = True
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub

Private Sub xDesca_Change()
    myload
End Sub
Private Sub xgroup_LostFocus()
With xgroup
If Not .MatchedWithList Then .BoundText = ""
If grid1.Rows = 2 Then
    grid1.TextMatrix(grid1.Rows - 1, 2) = .BoundText
    grid1.TextMatrix(grid1.Rows - 1, 3) = xSection.BoundText
End If
End With
End Sub
Private Sub xGroupMain_Validate(Cancel As Boolean)
    myload
End Sub
Private Sub xITEM_Change()
    myload
End Sub
Private Sub xGroup_Click(Area As Integer)
    If Area = 2 Then myload
End Sub
Private Sub xgroup_Validate(Cancel As Boolean)
    myload
End Sub
Private Sub xSection_Click(Area As Integer)
    If Area = 2 Then myload
End Sub
Private Sub xSection_LostFocus()
With xSection
If Not .MatchedWithList Then .BoundText = ""
If grid1.Rows = 2 Then
    grid1.TextMatrix(grid1.Rows - 1, 3) = .BoundText
    grid1.TextMatrix(grid1.Rows - 1, 2) = xgroup.BoundText
End If
End With
End Sub
Private Sub xSection_Validate(Cancel As Boolean)
    myload
End Sub
Private Function validRow(nRow) As Boolean
    If Trim(grid1.TextMatrix(nRow, 0)) = "" Then Exit Function
    If Trim(grid1.TextMatrix(nRow, 1)) = "" Then Exit Function
    validRow = True
End Function
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
     CellPos KeyCode, grid1.row, grid1.col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If Not (col = 2 Or col = 3 Or col = 4) Then CellPos KeyCode, row, col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal row As Long, ByVal col As Long)
KeyCode = 0
If col < grid1.Cols - 3 Then
    If col = 0 Then
        grid1.Select row, 1
    ElseIf col = 1 Then
        If grid1.TextMatrix(row, 2) <> "" Then
            If grid1.TextMatrix(row, 3) <> "" Then
                If grid1.TextMatrix(row, 4) <> "" Then
                    grid1.Select row, 5
                Else
                    grid1.Select row, 4
                End If
            Else
                grid1.Select row, 3
            End If
        End If
    ElseIf col = 2 Then
        If grid1.TextMatrix(row, 3) <> "" Then
            If grid1.TextMatrix(row, 4) <> "" Then
                grid1.Select row, 5
            Else
                grid1.Select row, 4
            End If
        Else
            grid1.Select row, 3
        End If
    Else
        grid1.Select row, col + 1
    End If
ElseIf row < grid1.Rows - 1 Then
    grid1.row = row + 1
    grid1.Select row + 1, 0
    grid1.ShowCell row + 1, 0
End If
End Sub
Private Sub myaddItem()
With grid1
    .AddItem ""
    If .Rows > 2 Then
        grid1.TextMatrix(.Rows - 1, 2) = grid1.TextMatrix(.Rows - 2, 2)
        grid1.TextMatrix(.Rows - 1, 3) = grid1.TextMatrix(.Rows - 2, 3)
    End If
    If xgroup.MatchedWithList And grid1.TextMatrix(.Rows - 1, 2) = "" Then
        grid1.TextMatrix(.Rows - 1, 2) = xgroup.BoundText
    End If
    If xSection.MatchedWithList And grid1.TextMatrix(.Rows - 1, 3) = "" Then
        grid1.TextMatrix(.Rows - 1, 3) = xSection.BoundText
    End If
End With
End Sub
Sub LoadPhoto(ByVal sItem As String)
'On Error GoTo myerror
Image1.Picture = LoadPicture("")
sItem = Trim(App.Path & "\PICT" & turn(sItem, "\") & sItem & ".jpg")
If Dir(sItem) = "" Then Exit Sub
Image1.Picture = LoadPicture(sItem)
Exit Sub
myerror:
Err.Clear
End Sub
Private Function doprint() As Boolean
SettingArray(cUpMargin) = MyMeasure(0)
SettingArray(cRightMargin) = MyMeasure(1)
SettingArray(cCardWidth) = MyMeasure(19)
SettingArray(cCardHeight) = MyMeasure(4)
SettingArray(cRows) = 5
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(21)

contemp.Execute "delete * From Card"

Dim tCard As New adodb.Recordset
tCard.Open "card", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)
nUP = 0

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = 1
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡

Dim nLine As Integer, cString As String
nSpace = MyMeasure(0.4)
For i = 1 To grid1.Rows - 1
    If Dir(App.Path & "\PICT" & turn(grid1.TextMatrix(i, 0), "\") & grid1.TextMatrix(i, 0) & ".jpg") <> "" Then
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        nLine = 0
                                    
        tCard.AddNew
        tCard!Right = MyMeasure(0.2)
        tCard!Top = MyMeasure(0)
        tCard!Width = MyMeasure(5)
        tCard!Height = MyMeasure(3)
        tCard!isPhoto = True
        tCard!Text = App.Path & "\PICT" & turn(grid1.TextMatrix(i, 0), "\") & grid1.TextMatrix(i, 0) & ".jpg"
        tCard!CardNo = nCard
        tCard.Update
                                          
                                          
        ' «·«”„
        tCard.AddNew
        tCard!Right = MyMeasure(10)
        tCard!Top = MyMeasure(nUP) + (nSpace * nLine)
        tCard!Width = MyMeasure(19)
        tCard!Height = 0
        tCard!FontName = "simplified arabic"
        tCard!FontBold = False
        tCard!ForeColor = vbBlack
        tCard!fontsize = 12
        tCard!Text = "Description  : " & grid1.TextMatrix(i, 1)
        tCard!TextAlign = taRightTop
        tCard!CardNo = nCard
        tCard.Update
        nLine = nLine + 1
                
        ' «·—ﬁ„ «·ﬁÊ„Ì
        tCard.AddNew
        tCard!Right = MyMeasure(10)
        tCard!Top = MyMeasure(nUP) + (nSpace * nLine)
        tCard!Width = MyMeasure(19)
        tCard!Height = 0
        tCard!FontName = "simplified arabic"
        tCard!FontBold = False
        tCard!ForeColor = vbBlack
        tCard!fontsize = 12
        tCard!Text = "Price : " & grid1.TextMatrix(i, 6)
        tCard!TextAlign = taRightTop
        tCard!CardNo = nCard
        tCard.Update
        nLine = nLine + 1
    End If
Next
tCard.Requery
doprint = Not (tCard.EOF And tCard.BOF)
Set CardTable = Nothing
End With
End Function
Private Function myreplace(row As Long) As Boolean
Dim aInsert(9, 1), cItem2 As String
aInsert(0, 0) = "item"
aInsert(0, 1) = addstring(grid1.TextMatrix(row, 0))

aInsert(1, 0) = "desca"
aInsert(1, 1) = addstring(grid1.TextMatrix(row, 1))

aInsert(2, 0) = "[Group]"
aInsert(2, 1) = addvalue(grid1.TextMatrix(row, 2))

aInsert(3, 0) = "[SECTION]"
aInsert(3, 1) = addvalue(grid1.TextMatrix(row, 3))

aInsert(4, 0) = "[SUPLER]"
aInsert(4, 1) = addstring(grid1.TextMatrix(row, 4))

aInsert(5, 0) = "[COST]"
aInsert(5, 1) = Val(grid1.TextMatrix(row, 5))

aInsert(6, 0) = "[PRICE]"
aInsert(6, 1) = Val(grid1.TextMatrix(row, 6))

cItem2 = grid1.TextMatrix(row, 0)
If (Not IsNumeric(cItem2)) Or Len(cItem2) > 6 Then cItem2 = ""

aInsert(7, 0) = "item2"
aInsert(7, 1) = Val(cItem2)

aInsert(8, 0) = "BALANCE"
aInsert(8, 1) = Val(grid1.TextMatrix(row, 7))

aInsert(9, 0) = "SHOW"
aInsert(9, 1) = IIf(Val(grid1.TextMatrix(row, 8)) = 0, "0", "1")

con.BeginTrans
On Error GoTo myerror
If grid1.TextMatrix(row, grid1.Cols - 1) = "" Then
    con.Execute CreateInsert(aInsert, "FILE1_10")
    grid1.TextMatrix(row, grid1.Cols - 1) = grid1.TextMatrix(grid1.row, 0)
Else
    con.Execute CreateUpdate(aInsert, "FILE1_10", " WHERE FILE1_10.id_item = " & grid1.TextMatrix(row, grid1.Cols - 1))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function
Private Function getItems() As Long
Dim conmdb As New adodb.Connection, loctable As New adodb.Recordset
'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & App.Path & "\MDB\DATA.mdb"
Dim cFile As String

'con.Execute "Delete from file1_10"
cFile = "FILE1_10"
cString = "SELECT * FROM FILE1_10 WHERE ITEM IS NULL"
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert(25, 1)
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If

nRecord = 100
Do Until loctable.EOF
    nRecord = nRecord + 1
    
    aInsert(0, 0) = "item"
    aInsert(0, 1) = addstring(nRecord)

    aInsert(1, 0) = "desca"
    aInsert(1, 1) = addstring(loctable!Desca & "")

    aInsert(2, 0) = "[Group]"
    aInsert(2, 1) = addvalue(2)

    aInsert(3, 0) = "[SECTION]"
    aInsert(3, 1) = addvalue(2)

    aInsert(4, 0) = "[SUPLER]"
    aInsert(4, 1) = addstring("000001")

    
    con.Execute CreateInsert(aInsert, "FILE1_10"), nAffect
    loctable.MoveNext
    getItems = getItems + nAffect
Loop
lastsub:
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getItems = -1
GoTo lastsub
End Function

