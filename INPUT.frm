VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form inputfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ê«—œ"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
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
   ScaleHeight     =   8700
   ScaleWidth      =   13050
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
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
      Left            =   5775
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   120
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
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
      Left            =   225
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   825
      Width           =   1320
   End
   Begin VB.CommandButton CMD_BAR 
      Caption         =   " —ÕÌ· «” Ìﬂ—"
      Height          =   420
      Left            =   3660
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   90
      Width           =   1995
   End
   Begin VB.Frame Frame6 
      Height          =   555
      Left            =   -1035
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   -90
      Visible         =   0   'False
      Width           =   3660
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   135
         Width           =   3540
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   7425
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   -45
      Width           =   5535
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1380
      Left            =   5085
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   540
      Width           =   7845
      Begin VB.TextBox xRemark 
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
         Left            =   150
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   975
         Width           =   6390
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
         Width           =   1425
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   3825
         TabIndex        =   21
         Top             =   540
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "„·«ÕŸ«  "
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
         Left            =   6645
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1035
         Width           =   930
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
         TabIndex        =   22
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
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
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
         TabIndex        =   7
         Top             =   210
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   3555
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   540
      Width           =   1500
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   540
         Width           =   1320
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
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
      Height          =   570
      Left            =   11025
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   8010
      Width           =   1980
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         ToolTipText     =   "Move Last"
         Top             =   150
         Width           =   435
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   5955
      Left            =   45
      TabIndex        =   23
      Top             =   2025
      Width           =   12930
      _cx             =   22807
      _cy             =   10504
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
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
      RightToLeft     =   0   'False
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
   Begin VB.Frame Frame5 
      Height          =   600
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   8010
      Width           =   7890
      Begin VB.Label xTotalCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   1440
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «· ﬂ·›… :"
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
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·≈Ã„«·Ì:"
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
         Left            =   6975
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   225
         Width           =   780
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   180
         Width           =   1440
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
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
End
Attribute VB_Name = "inputfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim dLastdate As String, SEARCH31 As New Search3, search32 As New Search3
Dim CardTable As ADODB.Recordset, grdTable As New ADODB.Recordset
Dim tBalStore  As New ADODB.Recordset
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select File1_10.item,File1_10.Desca From file1_10 "
Generalarray(2) = "Order by file1_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(FILE1_10.ITEM LIKE 'cFilter%' or  DESCA LIKE  'cFilter%') "


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
        CON.Execute "insert into FILE1_80h(doc_no,[Remark],[date],store)" & _
                    " values(" & _
                    addstring(xDoc_No.Text) & "," & _
                    addstring(xRemark.Text) & "," & _
                    DateSq(xDate.Text) & "," & _
                    addstring(xstore.BoundText) & _
                    ")"
    Else
        CON.Execute "update FILE1_80h " & _
                    "set [date] = " & DateSq(xDate.Text) & "," & _
                    "store = " & MyParn(xstore.BoundText) & "," & _
                    " Remark= " & MyParn(xRemark.Text) & _
                    " where doc_no = " & MyParn(xDoc_No.Text)
    End If
    If Err.Number = 0 Then
       ' Õ–› Õ—ﬂ… √’‰«› «·„” ‰œ
        CON.Execute " Delete * From FILE1_80 where Doc_No = " & MyParn(xDoc_No.Text) & " and row > " & grid1.Rows - 2
        With grid1
            For i = 1 To .Rows - 2
                nCost = Val(GetDesca("select File1_10.cost from file1_10 where file1_10.item = " & MyParn(grid1.TextMatrix(i, 0))))
                CON.Execute "Insert Into FILE1_80 (Doc_No,[Date],STORE,[Item],Quant,cost,Row) " & _
                " Values(" & _
                addstring(xDoc_No.Text) & "," & _
                DateSq(xDate.Text) & "," & _
                addstring(xstore.BoundText) & "," & _
                addstring(.TextMatrix(i, 0)) & "," & _
                addvalue(.TextMatrix(i, 2)) & "," & _
                nCost & "," & _
                  i & _
                ")"
                If Err.Number = -2147467259 Then
                    Err.Clear
                    CON.Execute "update FILE1_80 set " & _
                        "[date] = " & DateSq(xDate.Text) & "," & _
                        "STORE = " & addstring(xstore.BoundText) & "," & _
                        "item = " & addstring(.TextMatrix(i, 0)) & "," & _
                        "quant = " & Val(.TextMatrix(i, 2)) & _
                        " where doc_no = " & MyParn(xDoc_No.Text) & _
                        " and [row] = " & i
                End If
                If Err.Number <> 0 Then GoTo myerror
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
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(Search3.grid1.TextMatrix(Search3.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    grid1.TextMatrix(grid1.Row, 0) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = "1"
    GrdDesc grid1.Row
    If grid1.Row = grid1.Rows - 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 2) = 1
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 0
    ElseIf grid1.Row = grid1.Rows - 2 Then
        grid1.TextMatrix(grid1.Rows - 2, 2) = 1
        grid1.Select grid1.Rows - 1, 0
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
End Sub

Private Sub CMD_BAR_Click()
    With grid1
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 2)) > 0 Then
                CON.Execute "Insert Into ADDPRINT(Item,Quant,isPrint) " & _
                    " Values(" & _
                    addstring(.TextMatrix(i, 0)) & "," & _
                    addvalue(.TextMatrix(i, 2)) & "," & _
                    "TRUE" & _
                    ")"
            End If
        Next i
    End With

End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    CON.BeginTrans
    CON.Execute " Delete * From FILE1_80 where Doc_No = " & MyParn(xDoc_No.Text)
    CON.Execute " Delete * From FILE1_80H where Doc_No = " & MyParn(xDoc_No.Text)
    CON.CommitTrans
    CardTable.Requery
    
    CmdNewInv_Click
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
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE, Format([DATE],'yyyy/mm/dd'),FILE0_40.DESCA " & _
                  " FROM FILE1_80H INNER JOIN FILE0_40 ON FILE1_80H.STORE = FILE0_40.CODE"

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
foundOther
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
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
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
bEdit = True
Set CardTable = New ADODB.Recordset
CardTable.Open "FILE1_80H", CON, adOpenKeyset, adLockOptimistic, adCmdTable
grdTable.Open "select * from FILE1_80 ORDER BY ROW", CON, adOpenStatic, adLockReadOnly, adCmdText

data1.ConnectionString = CON.ConnectionString
data1.RecordSource = "FILE0_40"

Set xstore.RowSource = data1
xstore.ListField = "Desca"
xstore.BoundColumn = "Code"
tBalStore.Open "ITEM_BAL_S", CON, adOpenKeyset, adLockOptimistic, adCmdTable
With grid1
    .Cols = 7
    .Rows = 2
    .Editable = flexEDKbdMouse
    .FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹‹‹›|" & "«·ﬂ„Ì…|" & "«·—’Ìœ|" & "«· ﬂ·›…|" & "«·«Ã„«·Ì|" & ""
    .ColWidth(0) = 2500
    .ColWidth(1) = 5000
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColHidden(6) = True
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(5) = flexAlignRightCenter
    '.ColComboList(5) = "..."
    '.ColComboList(0) = "..."
End With
CmdNewInv_Click
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload Search3
Unload SEARCH31
If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
tBalStore.Close
Set CardTable = Nothing
Set tBalStore = Nothing
Set grdTable = Nothing
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If grid1.Col = 0 Then GrdDesc grid1.Row
CalcTotals
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then
    ItemsLookupAll Me, Search3
Else
    If Val(grid1.TextMatrix(Row, 2)) <= 0 Then
        MsgBox "ﬂ„Ì… «·«’‰«› «ﬁ· „‰ «Ê  ”«ÊÌ ’›—"
        Exit Sub
    End If
    
    If Val(grid1.TextMatrix(Row, grid1.Cols - 1)) = 0 Then
        MsgBox "«·’‰› ·„ ÌÕ›Ÿ »⁄œ"
        Exit Sub
    End If
End If
End Sub

Private Sub Grid1_EnterCell()
If grid1.Col = 1 Or grid1.Col = 3 Or grid1.Col = 4 Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
SetKbLayout IIf(grid1.Col = 0, Lang_EN, Lang_AR)

End Sub
Private Sub Grid1_GotFocus()
With grid1
    If grid1.Row <= 1 Then
    .Select 1, 0, 1, 0
    .ShowCell 1, 0
    End If
End With
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 Then grid1.AddItem "", grid1.Row
End Sub

Private Sub Grid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Col = 2 Then
        grid1.Row = Row + 1
        grid1.Col = IIf(Row = grid1.Rows - 2, 0, 3)
    ElseIf Col = 0 Then
        grid1.Col = 2
     End If
End If

End Sub

Private Sub Grid1_LostFocus()
SetKbLayout Lang_AR
End Sub

Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then grid1.AddItem ""
If grid1.Col = 2 Then
    tBalStore.Filter = " ITEM = " & MyParn(grid1.TextMatrix(Row, 0)) & " AND STORE = " & MyParn(xstore.BoundText)
    If Not tBalStore.EOF Then
        nBalance = Format(Val(tBalStore!BAL & ""), "#0.00")
    Else
        nBalance = 0
    End If
    
    grid1.TextMatrix(Row, 3) = nBalance
End If
End Sub

Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Then
    nFound = FoundOtheritem(Row, Col, Trim(grid1.EditText))
    If nFound <> -1 Then
        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
        Cancel = True
    End If
    
    cItem = GetDesca("select item from file1_10 where item = " & MyParn(.EditText)) & ""
    If cItem = "" Then
        MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
        Exit Sub
    End If
End If
End With
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
If xstore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·«Ê·"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If


With grid1
For i = 1 To .Rows - 2
    If .TextMatrix(i, 0) = "" Then
        .Select i, 0, i, grid1.Cols - 1
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
        .Select i, 0, i, grid1.Cols - 1
        MsgBox "ﬂ„Ì… «·’‰› €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub MyLoad()
xDoc_No.Text = CardTable!doc_no
'xusername.Text = TurnValue(CardTable!UserName, Null, "")
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xstore.BoundText = CardTable!store
xRemark.Text = CardTable!remark & ""
grdTable.Filter = "DOC_NO = " & MyParn(xDoc_No.Text)
With grid1
    .Rows = 1
    Do Until grdTable.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = TurnValue(grdTable!Item, Null, "")
        .TextMatrix(.Rows - 1, 1) = GetDesca("select desca from file1_10 where item = " & MyParn(grdTable!Item & ""))
        .TextMatrix(.Rows - 1, 2) = grdTable!Quant & ""
        .TextMatrix(.Rows - 1, 4) = grdTable!cost & ""
         grdTable.MoveNext
    Loop
    .AddItem ""
End With
Handlecontrols LoadMode
CalcTotals
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
xstore.BoundText = ""
xTotal.Caption = ""
xTotalCost.Caption = ""
xRemark.Text = ""
grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewinv.Enabled = nMode = LoadMode And bEdit
CmdSave.Enabled = (bEdit)
CmdDelInv.Enabled = nMode = LoadMode And bEdit
cmdFirst.Enabled = (nMode = LoadMode)
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
Private Sub Grid1_ChangeEdit()
'If Grid1.Col = 1 Then GrdDesc Grid1.Row
'CalcTotals
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, Search3
End If

If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        grid1.RemoveItem grid1.Row
    End If
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case Col
    Case 0
        If KeyCode = 27 Then Exit Sub
        If KeyCode = 112 Then ItemsLookupAll Me, Search3
End Select
End Sub
Private Sub GrdDesc(Row)
grid1.TextMatrix(Row, 1) = ""
grid1.TextMatrix(Row, 2) = ""
grid1.TextMatrix(Row, 3) = ""
grid1.TextMatrix(Row, 4) = ""

tBalStore.Filter = " ITEM = " & MyParn(grid1.TextMatrix(Row, 0)) & " AND STORE = " & MyParn(xstore.BoundText)
If Not tBalStore.EOF Then
    nBalance = Format(Val(tBalStore!BAL & ""), "#0.00")
Else
    nBalance = 0
End If
grid1.TextMatrix(Row, 3) = nBalance
If grid1.TextMatrix(Row, 0) = "" Then Exit Sub
    grid1.TextMatrix(Row, 1) = GetDesca("Select desca from file1_10 where item = " & MyParn(grid1.TextMatrix(Row, 0))) & ""
   ' If Trim(xSTORE.BoundText) <> "" And IsDate(xDate.Text) Then Grid1.TextMatrix(Row, 3) = Format(RetItemBalance(Grid1.TextMatrix(Row, 0), xSTORE.Text, xDate.Text), "#0.0000")
'End If
End Sub
Private Function CalcTotals()
Dim nTotalQuant As Double, nTotalCost As Double
With grid1
For i = 1 To grid1.Rows - 2
    grid1.TextMatrix(i, 5) = Val(grid1.TextMatrix(i, 2)) * Val(grid1.TextMatrix(i, 4))
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(i, 2))
    nTotalCost = nTotalCost + Val(grid1.TextMatrix(i, 5))
Next
xTotal.Caption = Format(nTotalQuant, "Fixed")
xTotalCost.Caption = Format(nTotalCost, "Fixed")
End With
End Function
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub foundOther()
For i = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(i, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 1) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & nRow
        Exit Sub
    End If
Next
End Sub
Private Function RetItemBalance(cItem, cStore, dDate) As Double
If cItem = "" Then Exit Function
movetable.Seek Array(cItem, cStore), adSeekFirstEQ
Do Until movetable.EOF
    If IsNull(movetable!Date) Then Exit Do
    If Trim(movetable!Item) <> cItem Or cStore <> movetable!store Or DateValue(movetable!Date) > DateValue(Format(dDate, "dd-mm-yyyy")) Then Exit Do
    If Not (movetable!Type = cItemmove And movetable!doc_ID = xDoc_No.Text) Then
        RetItemBalance = RetItemBalance + TurnValue(movetable!In, Null, 0) - TurnValue(movetable!out, Null, 0)
    End If
    movetable.MoveNext
Loop
End Function
Private Sub doprint()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For i = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!str21 = "„” ‰œ Ê«—œ —ﬁ„ : " & Format(xDoc_No.Text)
    temptable!date3 = DateFix(xDate.Text)
    temptable!str2 = TurnValue(xstore.Text)
    temptable!str3 = TurnValue(xRemark.Text, "", Null)
    temptable!str4 = TurnValue(grid1.TextMatrix(i, 0))
    temptable!str5 = TurnValue(grid1.TextMatrix(i, 1))
    temptable!val2 = TurnValue(Val(grid1.TextMatrix(i, 2)))
    temptable!val1 = TurnValue(Val(grid1.TextMatrix(i, 4)))
    temptable!val3 = TurnValue(Val(grid1.TextMatrix(i, 5)))
    temptable!val4 = TurnValue(Val(xTotalCost.Caption))
    temptable!Val10 = i
    If Val(xTotal.Caption) <> 0 Then
        temptable!STR6 = MyOnly(Val(xTotalCost.Caption))
    End If
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\R_INPUT.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub





Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function

