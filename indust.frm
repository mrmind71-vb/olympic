VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form industFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ’‰Ì⁄"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
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
   ScaleHeight     =   8265
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame6 
      Height          =   555
      Left            =   -270
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   3660
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   135
         Width           =   3540
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   -45
      Width           =   5535
      Begin VB.CommandButton cmdPrint 
         Caption         =   "ÿ»«⁄… "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5490
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Caption         =   "Õ–› «·„” ‰œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
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
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "„” ‰œ ÃœÌœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2790
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Caption         =   "≈” ⁄·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4140
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   900
      Width           =   1455
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   540
         Width           =   1275
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -3555
      Top             =   -300
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
      Caption         =   "data1"
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
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   7425
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   7515
      Width           =   1995
      Begin VB.CommandButton cmdFirst 
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   150
         Width           =   435
      End
      Begin VB.CommandButton cmdPrevious 
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
         Left            =   570
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   150
         Width           =   435
      End
      Begin VB.CommandButton cmdNext 
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
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   150
         Width           =   435
      End
      Begin VB.CommandButton cmdLast 
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
         Left            =   1455
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Move Last"
         Top             =   150
         Width           =   435
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   -3555
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   1530
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   585
      Width           =   7845
      Begin VB.TextBox xCharge 
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
         Left            =   5265
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   900
         Width           =   1290
      End
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
         Left            =   5265
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
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo xStore1 
         Height          =   315
         Left            =   3870
         TabIndex        =   2
         Top             =   540
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xStore2 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "„’«—Ì› :"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   930
         Width           =   930
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ì „Œ“‰ :"
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Œ“‰ :"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   585
         Width           =   825
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   210
         Width           =   930
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " „Ê«œ Œ«„"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3750
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1890
      Width           =   9375
      Begin VSFlex7LCtl.VSFlexGrid GRID1 
         Height          =   3435
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   225
         Width           =   9195
         _cx             =   16219
         _cy             =   6059
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
   Begin VB.Frame Frame7 
      Caption         =   "«·„‰ Ã« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   5625
      Width           =   9375
      Begin VSFlex7LCtl.VSFlexGrid GRID1 
         Height          =   1500
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   225
         Width           =   9240
         _cx             =   16298
         _cy             =   2646
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
End
Attribute VB_Name = "industFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SEARCH31 As New Search3, search32 As New Search3
Dim CardTable As ADODB.Recordset, grdTable As New ADODB.Recordset, GRDTABLE2 As New ADODB.Recordset
Dim tBalStore  As ADODB.Recordset
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup(Index)
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select File1_10.item,File1_10.Desca From file1_10 where val(type & '') = " & IIf(Index = 0, 1, 0)

Generalarray(2) = "Order by file1_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(FILE1_10.ITEM LIKE 'cFilter%' or  %%DESCA%%) "


GrdArray(0, 0) = "ﬂÊœ «·’‰›"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·’‰›"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„ «·«’‰«›"
Search3.Show 1
End Sub
Private Function MyReplace() As Boolean
Dim nTry As Integer
On Error Resume Next
CON.BeginTrans
For nTry = 1 To 10
    If xDoc_No.Enabled Then
        CON.Execute "insert into FILE1_90h(doc_no,[date],store1,store2,charge)" & _
                    " values(" & _
                    addstring(xDoc_No.Text) & "," & _
                    DateSq(xDate.Text) & "," & _
                    addstring(xstore1.BoundText) & "," & _
                    addstring(xstore2.BoundText) & "," & _
                    Val(xCharge.Text) & _
                    ")"
    Else
        CON.Execute "update FILE1_90h " & _
                    "set [date] = " & DateSq(xDate.Text) & "," & _
                    "store1 = " & MyParn(xstore1.BoundText) & "," & _
                    "store2 = " & MyParn(xstore2.BoundText) & "," & _
                    "charge = " & Val(xCharge.Text) & _
                    " where doc_no = " & MyParn(xDoc_No.Text)
    End If
    
    If Err.Number = 0 Then
        With grid1(0)
       ' Õ–› Õ—ﬂ… √’‰«› «·„” ‰œ
            CON.Execute " Delete * From FILE1_90 where Doc_No = " & MyParn(xDoc_No.Text) & " and row > " & .Rows - 2
            For i = 1 To .Rows - 2
                nCost = itemCost(.TextMatrix(i, 0), xDate.Text)
                CON.Execute "Insert Into FILE1_90 (Doc_No,[Date],Store1,Store2,[Item],Quant,cost,Row) " & _
                " Values(" & _
                addstring(xDoc_No.Text) & "," & _
                DateSq(xDate.Text) & "," & _
                addstring(xstore1.BoundText) & "," & _
                addstring(xstore2.BoundText) & "," & _
                addstring(.TextMatrix(i, 0)) & "," & _
                addvalue(.TextMatrix(i, 2)) & "," & _
                nCost & "," & _
                  i & _
                ")"
                If Err.Number = -2147467259 Then
                    Err.Clear
                    CON.Execute "update FILE1_90 set " & _
                        "[date] = " & DateSq(xDate.Text) & "," & _
                        "Store1 = " & addstring(xstore1.BoundText) & "," & _
                        "Store2 = " & addstring(xstore2.BoundText) & "," & _
                        "item = " & addstring(.TextMatrix(i, 0)) & "," & _
                        "quant = " & Val(.TextMatrix(i, 2)) & "," & _
                        "cost = " & nCost & _
                        " where doc_no = " & MyParn(xDoc_No.Text) & _
                        " and [row] = " & i
                End If
                If Err.Number <> 0 Then Exit For
            Next
        End With
    End If
    
    If Err.Number = 0 Then
       ' Õ–› Õ—ﬂ… √’‰«› «·„” ‰œ
        With grid1(1)
        CON.Execute " Delete * From FILE1_91 where Doc_No = " & MyParn(xDoc_No.Text) & " and row > " & .Rows - 2
        For i = 1 To .Rows - 2
            CON.Execute "Insert Into FILE1_91 (Doc_No,[Date],Store1,Store2,[Item],Quant,Row) " & _
            " Values(" & _
            addstring(xDoc_No.Text) & "," & _
            DateSq(xDate.Text) & "," & _
            addstring(xstore1.BoundText) & "," & _
            addstring(xstore2.BoundText) & "," & _
            addstring(.TextMatrix(i, 0)) & "," & _
            addvalue(.TextMatrix(i, 2)) & "," & _
              i & _
            ")"
            If Err.Number = -2147467259 Then
                Err.Clear
                CON.Execute "update FILE1_91 set " & _
                    "[date] = " & DateSq(xDate.Text) & "," & _
                    "Store1 = " & addstring(xstore1.BoundText) & "," & _
                    "Store2 = " & addstring(xstore2.BoundText) & "," & _
                    "item = " & addstring(.TextMatrix(i, 0)) & "," & _
                    "quant = " & Val(.TextMatrix(i, 2)) & _
                    " where doc_no = " & MyParn(xDoc_No.Text) & _
                    " and [row] = " & i
            End If
            If Err.Number <> 0 Then Exit For
        Next
        End With
    End If
        
    If Err.Number = 0 Then Exit For
    If Err.Number = -2147467259 And nTry < 10 Then
        Err.Clear
        xDoc_No.Text = RetZero(Val(xDoc_No.Text) + 1)
    End If
    If Err.Number <> 0 Then GoTo myerror
Next
CON.CommitTrans
MyReplace = True
Exit Function
myerror:
CON.RollbackTrans
If Err.Number <> 0 Then MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1(0).Name Then
    Index = ActiveControl.Index
    If Index = 1 And grid1(Index).Row > 1 Then Exit Sub
    grid1(Index).TextMatrix(grid1(Index).Row, 0) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    grid1(Index).TextMatrix(grid1(Index).Row, 2) = "1"
    GrdDesc grid1(Index).Row, Index
    If grid1(Index).Row = grid1(Index).Rows - 1 Then
        grid1(Index).TextMatrix(grid1(Index).Rows - 1, 2) = 1
        grid1(Index).AddItem ""
        grid1(Index).Select grid1(Index).Rows - 1, 0
    ElseIf grid1(Index).Row = grid1(Index).Rows - 2 Then
        grid1(Index).TextMatrix(grid1(Index).Rows - 2, 2) = 1
        grid1(Index).Select grid1(Index).Rows - 1, 0
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_no = " & MyParn(SEARCH31.grid1.TextMatrix(SEARCH31.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    MyLoad
    SEARCH31.Hide
ElseIf ActiveControl.Name = xDoc_No.Name Then
    xDoc_No.Text = SEARCH31.grid1.TextMatrix(SEARCH31.grid1.Row, 0)
    SEARCH31.Hide
Else
    ActiveControl.Text = search32.grid1.TextMatrix(search32.grid1.Row, 0)
    Unload search32
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Unload Search
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    CON.BeginTrans
    CON.Execute " Delete * From FILE1_90 where Doc_No = " & MyParn(xDoc_No.Text)
    CON.Execute " Delete * From FILE1_90H where Doc_No = " & MyParn(xDoc_No.Text)
    CON.CommitTrans
    CardTable.Requery
    If CardTable.BOF And CardTable.BOF Then
        myDefine
    Else
        CardTable.Find "doc_no < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.EOF Then CardTable.MoveFirst
        MyLoad
    End If
    MsgBox " „ Õ–› «·„” ‰œ »‰Ã«Õ"
End If
Exit Sub
myerror:
CON.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
If MsgBox("Œ—ÊÃ !! ” ›ﬁœ ﬂ· «·»Ì«‰«  «·€Ì— „Õ›ÊŸ… ! „Ê«›ﬁ ø", vbYesNo + vbDefaultButton2) = vbYes Then Unload Me
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE, Format([DATE],'yyyy/mm/dd'),FILE0_40.DESCA,FILE0_40_1.DESCA " & _
                  " FROM (FILE1_90H INNER JOIN FILE0_40 ON FILE1_90H.Store1 = FILE0_40.CODE) INNER JOIN FILE0_40 AS FILE0_40_1 ON FILE1_90H.STORE2 = FILE0_40_1.CODE "

Generalarray(2) = "Order by Date"
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«·«”„-«· «—ÌŒ"
listarray(0, 1) = "(Doc_No Like '%cFilter%' OR " & _
                  " ##[DATE]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "„‰ „Œ“‰"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "≈·Ì „Œ“‰"
GrdArray(4, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
Load SEARCH31
SEARCH31.Caption = "«” ⁄·«„"
SEARCH31.Show 1
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
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
On Error Resume Next
myDefine
xDoc_No.SetFocus
End Sub

Private Sub cmdPrint_Click()
doprint
End Sub
Private Sub cmdSave_Click()
foundOther 0
foundOther 1
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
MsgBox " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Requery
'CardTable.FindFirst "Doc_No = " & MyParn(xDoc_No.Text)
'If xDoc_No.Enabled Then
    'CmdNewInv_Click
'Else
    CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    MyLoad
'End If
End Sub
Private Sub CmdUndo_Click()
If CardTable.BOF And CardTable.EOF Then
    myDefine
    Exit Sub
End If
'CardTable.FindFirst "Doc_No = " & MyParn(xDoc_No.Text)
CardTable.Find "DOC_NO = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
MyLoad
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 83 Then cmdSave_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
'dLastdate = lastDate("FILE1_90")
bEdit = True
Set CardTable = New ADODB.Recordset
CardTable.Open "FILE1_90H", CON, adOpenKeyset, adLockOptimistic, adCmdTable

'tBalStore.Open "file1_11", CON, adOpenDynamic, adLockOptimistic, adCmdTableDirect
'tBalStore.Index = "ndxStore"
grdTable.Open "select * from FILE1_90 ORDER BY ROW", CON, adOpenStatic, adLockReadOnly, adCmdText
GRDTABLE2.Open "select * from FILE1_91 ORDER BY ROW", CON, adOpenStatic, adLockReadOnly, adCmdText

data1.ConnectionString = CON.ConnectionString
data1.RecordSource = "FILE0_40"
Set xstore1.RowSource = data1
xstore1.ListField = "Desca"
xstore1.BoundColumn = "Code"

Set xstore2.RowSource = data1
xstore2.ListField = "Desca"
xstore2.BoundColumn = "Code"
For i = 0 To 1
    With grid1(i)
        .Cols = 7
        .Rows = 2
        .Editable = flexEDKbdMouse
        .FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹‹‹›|" & "«·ﬂ„Ì…|" & "«·—’Ìœ|" & "«· ﬂ·›…|" & "«·«Ã„«·Ì|" & ""
        .ColWidth(0) = 1500
        .ColWidth(1) = 5500
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColHidden(6) = True
        .ColHidden(3) = True
        .ColHidden(4) = True
        .ColHidden(5) = True
        
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColComboList(5) = "..."
        .ColComboList(0) = "..."
    End With
Next
If Not (CardTable.BOF And CardTable.EOF) Then
'    CmdNewInv_Click
    CardTable.MoveLast
    MyLoad
Else
    myDefine
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
grdTable.Close
GRDTABLE2.Close
Set CardTable = Nothing
Set tBalStore = Nothing
Set grdTable = Nothing
Set GRDTABLE2 = Nothing
Unload Search3
Unload SEARCH31
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
If grid1(Index).Col = 0 Then GrdDesc grid1(Index).Row, Index
'calcTotals
End Sub
Private Sub grid1_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then ItemsLookup Index
End Sub

Private Sub Grid1_EnterCell(Index As Integer)
If grid1(Index).Col = 1 Or grid1(Index).Col = 3 Or grid1(Index).Col = 4 Then
    grid1(Index).Editable = flexEDNone
Else
    grid1(Index).Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_GotFocus(Index As Integer)
With grid1(Index)
'    If .Row = 0 Then
'    .Select 1, 0, 1, 0
'    End If
End With
End Sub
Private Sub Grid1_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Index = 1 And Row > 1 Then Exit Sub
If grid1(Index).Row = grid1(Index).Rows - 1 Then grid1(Index).AddItem ""
End Sub

Private Sub xDate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CmdInform_Click
End Sub
Private Function MYVALID() As Boolean
CardTable.Find "DOC_NO = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF And xDoc_No.Enabled Then
    MsgBox "„” ‰œ »‰›” «·—ﬁ„ „‰ ﬁ»·"
    Exit Function
End If

If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
'If IsDate(dLastdate) Then
'    If DateValue(xDate.Text) <= DateValue(dLastdate) Then
'        MsgBox "«· «—ÌŒ «ﬁ· „‰ «Œ—  «—ÌŒ «€·«ﬁ"
'        Exit Function
'    End If
'End If
If xstore1.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·«Ê·"
    Exit Function
End If

If xstore2.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·À«‰Ì"
    Exit Function
End If

If grid1(1).Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If


With grid1(Index)
For i = 1 To .Rows - 2
    If .TextMatrix(i, 0) = "" Then
        .Select i, 0, i, .Cols - 1
        MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
        Exit Function
    Else
        cItem = GetDesca("select item from file1_10 where item = " & MyParn(.TextMatrix(i, 0))) & ""
        If cItem = "" Then
            MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
            Exit Function
        End If
    End If
    If Val(.TextMatrix(i, 2)) = 0 Then
        .Select i, 0, i, .Cols - 1
        MsgBox "ﬂ„Ì… «·’‰› €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub MyLoad()
xDoc_No.Text = CardTable!doc_no
xusername.Text = TurnValue(CardTable!UserName, Null, "")
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xstore1.BoundText = CardTable!store1 & ""
xstore2.BoundText = CardTable!Store2 & ""
xCharge.Text = Format(CardTable!CHARGE, "fixed")
grdTable.Filter = "DOC_NO = " & MyParn(xDoc_No.Text)
With grid1(0)
    .Rows = 1
    Do Until grdTable.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = TurnValue(grdTable!Item, Null, "")
        .TextMatrix(.Rows - 1, 1) = GetDesca("select desca from file1_10 where item = " & MyParn(grdTable!Item & ""))
        .TextMatrix(.Rows - 1, 2) = grdTable!Quant & ""
         grdTable.MoveNext
    Loop
    .AddItem ""
End With

GRDTABLE2.Filter = "DOC_NO = " & MyParn(xDoc_No.Text)
With grid1(1)
    .Rows = 1
    Do Until GRDTABLE2.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = TurnValue(GRDTABLE2!Item, Null, "")
        .TextMatrix(.Rows - 1, 1) = GetDesca("select desca from file1_10 where item = " & MyParn(GRDTABLE2!Item & ""))
        .TextMatrix(.Rows - 1, 2) = GRDTABLE2!Quant & ""
         GRDTABLE2.MoveNext
    Loop
    .AddItem ""
End With

Handlecontrols LoadMode
'calcTotals
End Sub
Private Sub myDefine()
If CardTable.EOF And CardTable.BOF Then
    xDoc_No.Text = RetZero("1")
Else
    CardTable.MoveLast
    xDoc_No.Text = RetZero(Val(CardTable!doc_no & "") + 1)
End If
xusername.Text = ""
xDate.Text = Format(Date, "dd-mm-yyyy")
xstore1.BoundText = ""
xstore2.BoundText = ""
'xTotal.Caption = ""
'xTotalCost.Caption = ""
grid1(0).Rows = 1
grid1(0).AddItem ""
grid1(1).Rows = 1
grid1(1).AddItem ""
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewinv.Enabled = nMode = LoadMode And bEdit
cmdSave.Enabled = (bEdit)
CmdDelInv.Enabled = nMode = LoadMode And bEdit
cmdfirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
End Sub
Private Sub xDoc_No_LostFocus()
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.BOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad
End Sub
Private Sub Grid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1(Index).Col = 0 Then
    ItemsLookup Index
End If

If KeyCode = 46 And grid1(Index).Row <> grid1(Index).Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        grid1(Index).RemoveItem grid1(Index).Row
    End If
End If
End Sub
Private Sub grid1_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case Col
    Case 0
        If KeyCode = 27 Then Exit Sub
        If KeyCode = 112 Then ItemsLookup Index
End Select
End Sub
Private Sub GrdDesc(Row, Index)
With grid1(Index)
.TextMatrix(Row, 1) = ""
.TextMatrix(Row, 2) = ""
.TextMatrix(Row, 3) = ""
.TextMatrix(Row, 4) = ""
If .TextMatrix(Row, 0) = "" Then Exit Sub
.TextMatrix(Row, 1) = GetDesca("Select desca from file1_10 where item = " & MyParn(.TextMatrix(Row, 0))) & ""
End With
   ' If Trim(xStore1.BoundText) <> "" And IsDate(xDate.Text) Then Grid1.TextMatrix(Row, 3) = Format(RetItemBalance(Grid1.TextMatrix(Row, 0), xStore1.Text, xDate.Text), "#0.0000")
'End If
End Sub
Private Function CalcTotals()
Dim nTotalQuant As Double, nTotalCost As Double
With grid1(Index)
For i = 1 To .Rows - 2
    .TextMatrix(i, 5) = Val(.TextMatrix(i, 2)) * Val(.TextMatrix(i, 4))
    nTotalQuant = nTotalQuant + Val(.TextMatrix(i, 2))
    nTotalCost = nTotalCost + Val(.TextMatrix(i, 5))
Next
xTotal.Caption = Format(nTotalQuant, "Fixed")
xTotalCost.Caption = Format(nTotalCost, "Fixed")
End With
End Function
Private Function FoundOtherRow(nRow, nCol, Index) As Integer
FoundOtherRow = -1
With grid1(Index)
For i = 1 To .Rows - 2
    If i <> nRow Then
        If Trim(.TextMatrix(i, nCol)) = Trim(.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End With
End Function
Private Sub foundOther(Index)
With grid1(Index)
For i = 1 To .Rows - 2
    nRow = FoundOtherRow(i, 0, Index)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & .TextMatrix(nRow, 1) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & nRow & IIf(Index = 0, " «·„Ê«œ «·Œ«„", " «·„‰ Ã« ")
        Exit Sub
    End If
Next
End With
End Sub
Private Sub doprint()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With grid1(Index)
For i = 1 To .Rows - 2
    temptable.AddNew
    temptable!str21 = "„” ‰œ  ÕÊÌ· —ﬁ„ : " & Format(xDoc_No.Text)
    temptable!date3 = DateFix(xDate.Text)
    temptable!str2 = TurnValue(xstore1.Text)
    temptable!str3 = TurnValue(xstore2.Text)
    temptable!str4 = TurnValue(.TextMatrix(i, 0))
    temptable!str5 = TurnValue(.TextMatrix(i, 1))
    temptable!val2 = TurnValue(Val(.TextMatrix(i, 2)))
    temptable!val1 = TurnValue(Val(.TextMatrix(i, 4)))
    temptable!val3 = TurnValue(Val(.TextMatrix(i, 5)))
    temptable!val4 = TurnValue(Val(xTotalCost.Caption))
    temptable!Val10 = i
    If Val(xTotal.Caption) <> 0 Then
        temptable!STR6 = MyOnly(Val(xTotalCost.Caption))
    End If
    temptable.Update
Next
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\TRANS.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub






