VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form impcostpricefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÊÓÚíÑ ÇáÑÓÇáÉ"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10695
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      Height          =   975
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   9645
      Width           =   4530
      Begin VB.TextBox xfilterItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "ÈÍË"
         Top             =   525
         Width           =   2535
      End
      Begin VB.TextBox xfilter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "ÈÍË"
         Top             =   180
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "ÈÍË ÈßæÏ ÇáÕäÝ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   525
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "ÈÍË ÈÇáÕäÝ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ÈíÇäÇÊ ÇáÑÓÇáÉ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   90
      Width           =   9510
      Begin VB.Label xStore 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   1995
      End
      Begin VB.Label xcodedesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   585
         Width           =   3480
      End
      Begin VB.Label xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label xDoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   1995
      End
      Begin VB.Label Label4 
         Caption         =   "ÇáãÎÒä :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "ÊÇÑíÎ ÇáÑÓÇáÉ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "ÇÓã ÇáãæÑÏ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8055
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ÑÞã ÇáÑÓÇáÉ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8055
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   330
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   9645
      Width           =   3615
      Begin VB.CommandButton cmdExit 
         Caption         =   "ÎÑæÌ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1680
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "ÊÓÚíÑ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1680
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   14910
      _cx             =   26300
      _cy             =   14843
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      BackColorBkg    =   -2147483636
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   4230
      Top             =   225
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
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   375
      Left            =   9390
      TabIndex        =   15
      Top             =   9720
      Visible         =   0   'False
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "impcostpricefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
    If Not myreplace Then Exit Sub
    Inform "Êã ÍÝÙ ÇáãÓÊäÏ ÈäÌÇÍ"
    myload
End Sub
Private Sub Form_Load()
openCon con
xDoc_No.Caption = impcostfrm.xDoc_No.Text
xDate.Caption = impcostfrm.xDate.Text
xCodeDesca.Caption = impcostfrm.xCodeDesca.Caption
xStore.Caption = impcostfrm.xStore.Text
Set Grid1.DataSource = DATA3
DATA3.ConnectionString = strCon
myload
End Sub
Private Sub myload()
    cString = "SELECT FILE7_60.ROW,FILE1_10.ITEM,FILE1_10.DESCA,FILE7_60.Quant,FILE7_60.Price,FILE1_10.DISCOUNT,0 AS TOTALFRGN,FILE7_60.TOTAL,FILE7_60.COST,FILE7_60.RATE1,FILE7_60.PRICE1,FILE7_60.RATE2,FILE7_60.PRICE2,FILE1_10.MAXDISC,FILE7_60.ID" & _
          " FROM FILE7_60 INNER JOIN FILE1_10 ON FILE7_60.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.Caption) & _
          " ORDER BY FILE7_60.ROW"
    DATA3.RecordSource = cString
    DATA3.Refresh
    'CalcTotals
    Fixgrd
End Sub
Private Sub Fixgrd()
With Grid1
    .Editable = flexEDKbd
'                     0     1           2           3           4           5           6               7           8           9           10              11              12              13
    .FormatString = "ã|" & "ßæÏ|" & "ÇáÕäÜÝ|" & "ÇáßãíÉ|" & "ÇáÓÚÑ|" & "ÇáÎÕã|" & "ÇáÅÌãÇáí ÈÇáÚãáÉ|" & "ÇáÇÌãÇáí|" & "Ê ÇáæÍÏÉ|" & "äÓÈÉ ÌãáÉ|" & "ÓÚÑ ÌãáÉ|" & "äÓÈÉ ãÓÊåáß|" & "ÓÚÑ ãÓÊåáß|" & "ÍÏ ÇáÎÕã"
    .WordWrap = True
    .RowHeight(0) = 700
    .ColWidth(0) = 500
    .ColWidth(1) = 1500
    .ColWidth(2) = 3000
    .ColWidth(3) = 0
    .ColWidth(4) = 0
    .ColWidth(5) = 1000
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .ColWidth(8) = 1100
    .ColWidth(9) = 1100
    .ColWidth(10) = 1100
    .ColWidth(11) = 1100
    .ColWidth(12) = 1100
    .ColWidth(13) = 800
    .ColHidden(14) = True
    .ExplorerBar = flexExSortShow
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub Grid1_EnterCell()
If Grid1.Col > 8 Or Grid1.Col = 5 Then Grid1.Editable = flexEDKbdMouse Else Grid1.Editable = flexEDNone
If Grid1.Col = 9 And Grid1.Row > 1 Then
    If Val(Grid1.TextMatrix(Grid1.Row, 9)) = 0 Then
        Grid1.TextMatrix(Grid1.Row, 9) = Grid1.TextMatrix(Grid1.Row - 1, 9)
        NN = Val(Grid1.TextMatrix(Grid1.Row, 9))
    With Grid1
    If (.Col = 9 Or .Col = 11) And NN <> 0 Then
        nRate = 1 + (NN / 100)
        Grid1.TextMatrix(.Row, .Col + 1) = myNear(Round(Val(Grid1.TextMatrix(.Row, .Col - 1)) * nRate, 2), 0.5)
    ElseIf (.Col = 10 Or .Col = 12) And Val(Grid1.TextMatrix(.Row, .Col - 2)) <> 0 Then
        Grid1.TextMatrix(.Row, .Col - 1) = Round((NN - Val(Grid1.TextMatrix(.Row, .Col - 2))) / Val(Grid1.TextMatrix(.Row, .Col - 2)) * 100, 2)
    End If
    End With
    
    
    End If
End If

If Grid1.Col = 11 And Grid1.Row > 1 Then
    If Val(Grid1.TextMatrix(Grid1.Row, 11)) = 0 Then
        Grid1.TextMatrix(Grid1.Row, 11) = Grid1.TextMatrix(Grid1.Row - 1, 11)
        NN = Val(Grid1.TextMatrix(Grid1.Row, 11))
    
        With Grid1
        If (.Col = 9 Or .Col = 11) And NN <> 0 Then
            nRate = 1 + (NN / 100)
            Grid1.TextMatrix(.Row, .Col + 1) = myNear(Round(Val(Grid1.TextMatrix(.Row, .Col - 1)) * nRate, 2), 0.5)
        ElseIf (.Col = 10 Or .Col = 12) And Val(Grid1.TextMatrix(.Row, .Col - 2)) <> 0 Then
            Grid1.TextMatrix(.Row, .Col - 1) = Round((NN - Val(Grid1.TextMatrix(.Row, .Col - 2))) / Val(Grid1.TextMatrix(.Row, .Col - 2)) * 100, 2)
        End If
        End With
    
    
    End If
End If

If Grid1.Col = 5 And Grid1.Row > 1 Then
    If Val(Grid1.TextMatrix(Grid1.Row, 5)) = 0 Then
        Grid1.TextMatrix(Grid1.Row, 5) = Grid1.TextMatrix(Grid1.Row - 1, 5)
    End If
End If

If Grid1.Col = 13 And Grid1.Row > 1 Then
    If Val(Grid1.TextMatrix(Grid1.Row, 13)) = 0 Then
        Grid1.TextMatrix(Grid1.Row, 13) = Grid1.TextMatrix(Grid1.Row - 1, 13)
    End If
End If

With Grid1
    If .Col = 12 Or .Col = 10 Then
        If IsNumeric(.TextMatrix(.Row, .Col)) Then
            If Val(Grid1.TextMatrix(.Row, .Col - 2)) <> 0 Then Grid1.TextMatrix(.Row, .Col - 1) = myNear(Round((Val(.TextMatrix(.Row, .Col)) - Val(Grid1.TextMatrix(.Row, .Col - 2))) / Val(Grid1.TextMatrix(.Row, .Col - 2)) * 100, 2), 1)
        Else
            Grid1.TextMatrix(.Row, .Col - 1) = 0
        End If
    End If
End With

End Sub
Private Sub Grid1_LeaveCell()
With Grid1
    If .Col = 11 Then
        If Val(.TextMatrix(.Row, 11)) = 0 Then
            .TextMatrix(.Row, 12) = .TextMatrix(.Row, 10)
        End If
    End If
End With
End Sub

Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid1
If (Col = 9 Or Col = 11) And Val(.EditText) <> 0 Then
'    If Val(grid1.TextMatrix(Row, Col + 1)) = 0 Then
        nRate = 1 + (Val(.EditText) / 100)
        Grid1.TextMatrix(Row, Col + 1) = myNear(Round(Val(Grid1.TextMatrix(Row, Col - 1)) * nRate, 2), 0.5)
'    End If
ElseIf (Col = 10 Or Col = 12) And Val(Grid1.TextMatrix(Row, Col - 2)) <> 0 Then
    If Val(.EditText) <> 0 Then
        Grid1.TextMatrix(Row, Col - 1) = myNear(Round((Val(.EditText) - Val(Grid1.TextMatrix(Row, Col - 2))) / Val(Grid1.TextMatrix(Row, Col - 2)) * 100, 2), 1)
    Else
        Grid1.TextMatrix(Row, Col - 1) = 0
    End If
End If
End With
End Sub
Private Function myreplace() As Boolean
Dim aGrid(3, 1)
On Error GoTo myerror
con.BeginTrans
prog1.Value = 0
prog1.Visible = True
With Grid1
     For i = 1 To .Rows - 1
'         prog1.Value = Round(I / (grid1.Rows - 1), 2) * 100
         aGrid(0, 0) = "Rate1": aGrid(0, 1) = Val(Grid1.TextMatrix(i, 9))
         'aGrid(0, 0) = "Rate1": aGrid(0, 1) = Round(Val(grid1.TextMatrix(I, 10) - Val(grid1.TextMatrix(I, 8))) / Val(grid1.TextMatrix(I, 8)) * 100, 2)
         aGrid(1, 0) = "Price1": aGrid(1, 1) = Val(Grid1.TextMatrix(i, 10))
         aGrid(2, 0) = "RATE2": aGrid(2, 1) = Val(Grid1.TextMatrix(i, 11))
         'aGrid(2, 0) = "RATE2": aGrid(2, 1) = Round(Val(grid1.TextMatrix(I, 12) - Val(grid1.TextMatrix(I, 8))) / Val(grid1.TextMatrix(I, 8)) * 100, 2)
         aGrid(3, 0) = "PRICE2": aGrid(3, 1) = Val(Grid1.TextMatrix(i, 12))
'         aGrid(4, 0) = "DISCOUNT": aGrid(4, 1) = Val(grid1.TextMatrix(I, 5))
         
         con.Execute CreateUpdate(aGrid, "FILE7_60", " where ID = " & Grid1.TextMatrix(i, Grid1.Cols - 1), Array(-1))
         
         If Val(Grid1.TextMatrix(i, 10)) > 0 Then con.Execute "update file1_10 set file1_10.price = " & Val(Grid1.TextMatrix(i, 10)) & " where file1_10.item = " & MyParn(Grid1.TextMatrix(i, 1))
         If Val(Grid1.TextMatrix(i, 12)) > 0 Then con.Execute "update file1_10 set file1_10.price2 = " & Val(Grid1.TextMatrix(i, 12)) & " where file1_10.item = " & MyParn(Grid1.TextMatrix(i, 1))
         If Val(Grid1.TextMatrix(i, 5)) > 0 Then con.Execute "update file1_10 set file1_10.DISCOUNT = " & Val(Grid1.TextMatrix(i, 5)) & " where file1_10.item = " & MyParn(Grid1.TextMatrix(i, 1))
         If Val(Grid1.TextMatrix(i, 13)) > 0 Then con.Execute "update file1_10 set file1_10.MAXDISC = " & Val(Grid1.TextMatrix(i, 13)) & " where file1_10.item = " & MyParn(Grid1.TextMatrix(i, 1))
         If Val(Grid1.TextMatrix(i, 4)) > 0 Then con.Execute "update file1_10 set file1_10.COSTIMP = " & Val(Grid1.TextMatrix(i, 4)) & " where file1_10.item = " & MyParn(Grid1.TextMatrix(i, 1))
         con.Execute "update file1_10 set file1_10.IMPORT  = 1 where file1_10.item = " & MyParn(Grid1.TextMatrix(i, 1))
     Next
     prog1.Visible = False
 End With
con.CommitTrans
myreplace = True
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Sub xfilter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd Grid1, xfilter.Text, 2
End If
End Sub
Private Sub xfilterItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd Grid1, xfilterItem.Text, 1
End If
End Sub

