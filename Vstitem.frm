VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form grditem1 
   Caption         =   "„ «»⁄… «·«’‰«›"
   ClientHeight    =   9030
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   15780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   15780
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   1125
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   630
      Width           =   4875
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "Vstitem.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3555
         Picture         =   "Vstitem.frx":27EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "Vstitem.frx":4CDD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2385
         Picture         =   "Vstitem.frx":7149
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Width           =   9690
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
         Left            =   2475
         MaxLength       =   15
         TabIndex        =   13
         Top             =   900
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   3615
      End
      Begin VB.TextBox xDescItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   900
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   7065
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   5490
         TabIndex        =   8
         Top             =   900
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroupMain 
         Height          =   315
         Left            =   135
         TabIndex        =   2
         Top             =   180
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   5490
         TabIndex        =   3
         Top             =   540
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ «·’‰› :"
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
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   945
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈”„ «·’‰› :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   4
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄… :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   8625
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   990
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì… :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label Label2 
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
         Height          =   240
         Index           =   3
         Left            =   8625
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   615
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   8625
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   780
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   8700
      Width           =   15780
      _ExtentX        =   27834
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   1260
      Top             =   405
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   -1845
      Top             =   -135
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   -1620
      Top             =   -225
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -1575
      Top             =   -180
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
      Height          =   7170
      Left            =   45
      TabIndex        =   20
      Top             =   1395
      Width           =   15675
      _cx             =   27649
      _cy             =   12647
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   300
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   21
      Top             =   8550
      Visible         =   0   'False
      Width           =   15780
      _ExtentX        =   27834
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "grditem1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFilesave As String
Dim con As New ADODB.Connection
Dim osearchitem As New Search31
Dim LastSalTable As New ADODB.Recordset
Dim LastImpTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub cmd_dem_Click()
'If Not MYVALID Then Exit Sub
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
cString = "SELECT FILE1_10.PACKAGE ,Sum(FILE1_11.[IN]- FILE1_11.[out]) AS Balance,FILE1_10.REORDER , FILE1_10.ITEM,FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP] AS FILE1_50GROUP,FILE1_10.[SECTION],FILE1_10.COST, file1_10.price3 , FILE1_50.DESCA AS FILE1_50DESCA,FILE1_50G.DESCA AS FILE1_50GDESCA,FILE1_10SC.DESCA AS FILE1_10SCDESCA " & _
            "FROM (((FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_50G ON FILE1_50.[GROUP] = FILE1_50G.CODE) LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE "
If xFact.Text <> "" Then cString = cString & turn(cString) & " fact = " & MyParn(xFact.Text)
If xGroup.BoundText <> "" Then cString = cString & turn(cString) & " file1_10.[GROUP]  = " & xGroup.BoundText
If xGroupMain.BoundText <> "" Then cString = cString & turn(cString) & " file1_50.[Group]  = " & xGroupMain.BoundText
If xSection.BoundText <> "" Then cString = cString & turn(cString) & " [Section] = " & xSection.BoundText
cString = cString & " GROUP BY FILE1_10.PACKAGE ,  file1_10.reorder , FILE1_10.REORDER , FILE1_10.ITEM,file1_10.price , FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP],FILE1_10.[SECTION],FILE1_10.COST, FILE1_50.DESCA,FILE1_50G.DESCA,FILE1_10SC.DESCA"
          
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
con.Execute "UPDATE FILE1_10 SET FILE1_10.COSTIMP = [FILE7_60].[PRICE]  FROM FILE1_10 INNER JOIN FILE7_60 ON FILE1_10.ITEM = FILE7_60.ITEM WHERE (((FILE7_60.ITEM) Is Not Null))"
With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!val10 = TurnValue(!Section, "", Null)
        temptable!str6 = TurnValue(!file1_10SCdesca, "", Null)
        temptable!val11 = TurnValue(!FILE1_50GROUP, "", Null)
        temptable!str7 = TurnValue(!file1_50GDESCA, "", Null)
        temptable!val12 = TurnValue(!Group, "", Null)
        temptable!str8 = TurnValue(!file1_50desca, "", Null)
        temptable!str1 = !Item
        temptable!str2 = !Desca
        
        temptable!val2 = TurnValue(!BALANCE, "", Null)
        temptable!val3 = TurnValue(!package, "", Null)
        temptable!val4 = Val(!COSTIMP & "")
        temptable!val5 = TurnValue(!dem, "", Null)
        
        temptable!str3 = TurnValue(!Fact, "", Null)
        temptable!str21 = "»Ì«‰ «’‰«› ·Â« ÿ·»Ì…  "
        temptable.Update
      .MoveNext
    Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.Report1.ReportFileName = App.Path & "\Reports\Item_dem.rpt"
    main.Report1.DataFiles(0) = tempFile
    main.Report1.Action = 1
End If
temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing

End Sub

Private Sub Cmd_Print_Click()
    
End Sub

Private Sub cmdDelinv_Click()
End Sub

Private Sub cmdExel_Click()
ToFileExel grid1, Array(1)
End Sub

Private Sub cmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
myload
End Sub

Private Sub CmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï √—’œ… „»Ì⁄«  „‘ —Ì«  ··√’‰«›  "
    If IsDate(xDate1.Text) Then cHead2 = "„‰ : " & Format(xDate1.Text, "DD-MM-YYYY")
    If IsDate(xDate2.Text) Then cHead2 = cHead2 & turn(cHead2, " ") & "Õ Ì : " & Format(xDate2.Text, "DD-MM-YYYY")
    PrintGrd.doprint Me.grid1, 0.8, -2, cHead1, cHead2, , False, False, 8, , Array(1)
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()
    openCon con
    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
    Set xSection.RowSource = data1
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File1_50G order by Desca"
    Set xGroupMain.RowSource = DATA2
    xGroupMain.ListField = "Desca"
    xGroupMain.BoundColumn = "Code"
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    Set grid1.DataSource = data4
    data4.ConnectionString = strCon
    Fixgrd
    grid1.Rows = 1
    LoadText Me
End Sub
Private Sub myload()
Dim cwhere As String

If IsDate(xDate1.Text) Then cwhere = " date < " & DateSq(xDate1.Text)
cField = myiif(cwhere, "[IN] - [OUT]") & " AS F_BAL"
cwhere = ""
If IsDate(xDate1.Text) Then cwhere = " date >= " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(xDate2.Text)

cField = cField & "," & _
        myiif(cwhere & turn(cwhere, " And ") & " ( TYPE = '2' OR TYPE = '7' )", "[IN] - [OUT]") & " AS PURCHASES"

cField = cField & "," & _
        myiif(cwhere & turn(cwhere, " And ") & "(TYPE = '11')", "[OUT] - [IN]") & " AS OUTPUT"

cField = cField & "," & _
        myiif(cwhere & turn(cwhere, " And ") & "(TYPE = '8')", "[OUT] - [IN]") & " AS DAMAGE"

cField = cField & "," & _
        myiif(cwhere & turn(cwhere, " And ") & "(TYPE = '10')", "[OUT] - [IN]") & " AS INPUT"

cField = cField & "," & _
        myiif(cwhere & turn(cwhere, " And ") & "(TYPE = 'Z')", "[OUT] - [IN]") & " AS STOCK"

cwhere = ""
If IsDate(xDate2.Text) Then cwhere = " date <= " & DateSq(xDate2.Text)
cField = cField & "," & _
         myiif(cwhere, "[IN] - [OUT]") & " AS endbal "

'                           0               1                 2                3
    cString = "  select file1_10.item , file1_10.desca,FILE1_10.COST," & _
                cField & _
                " from FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM  left join file1_50 on file1_10.[GROUP] = file1_50.code"

    If xGroup.BoundText <> "" Then cString = cString & turn(cString) & " file1_10.[GROUP]  = " & xGroup.BoundText
    If xGroupMain.BoundText <> "" Then cString = cString & turn(cString) & " file1_50.[Group]  = " & xGroupMain.BoundText
    If xSection.BoundText <> "" Then cString = cString & turn(cString) & "  [Section] = " & xSection.BoundText
    If Trim(xDesca.Text) <> "" Then cString = cString & turn(cString) & MyParnAnd(xDesca.Text, "file1_10.desca")
    With grid1
    cString = cString & " GROUP BY file1_10.item, file1_10.desca, file1_10.cost,FILE1_10.PRICE,file1_10.reorder,file1_10.package"
    data4.RecordSource = cString
    data4.Refresh
End With
Fixgrd
End Sub
Sub Fixgrd()
    With grid1
    .RowHeight(0) = 700
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·’‰›"
    .TextMatrix(0, 2) = "”⁄—  ﬂ·›…"
    .TextMatrix(0, 3) = "—’Ìœ" & Format(xDate1.Text, "dd-mm-yyyy")
    .TextMatrix(0, 4) = "„‘ —Ì« "
    .TextMatrix(0, 5) = "’«œ—"
    .TextMatrix(0, 6) = "Ê«—œ"
    .TextMatrix(0, 7) = "Â«·ﬂ"
    .TextMatrix(0, 8) = " ”ÊÌ… Ã—œ"
    .TextMatrix(0, 9) = "—’Ìœ " & Format(xDate2.Text, "dd-mm-yyyy")
    
    .FrozenCols = 2
    .ColWidth(0) = 2000
    .ColWidth(1) = 4000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1100
    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(8) = flexDTDouble
    .ColDataType(9) = flexDTDouble

    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 3, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 7, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 8, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 9, "#0", vbRed, vbYellow, True, "  "
    StatusBar1.Panels(1).Text = "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.Rows - 2
    If .Rows > 1 Then
        .TextMatrix(1, 0) = "«·≈Ã„«·Ì"
        .TextMatrix(1, 1) = "«·≈Ã„«·Ì"
        .MergeRow(1) = True
    End If
    .MergeCells = flexMergeFree

    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Unload Me
Set grditem1 = Nothing
End Sub

Private Sub Grid1_DblClick()
    If grid1.Row < 2 Then Exit Sub
    If grid1.Col = 1 Or grid1.Col = 0 Then
        Dim aData As Variant
        aData = AddFlag(Empty, "ITEM", grid1.TextMatrix(grid1.Row, 0))
        aData = AddFlag(aData, "DATE1", xDate1.Text)
        aData = AddFlag(aData, "DATE2", xDate2.Text)
        StoreMove.aData = aData
        StoreMove.Show
    Else
        Dim aWhere As Variant
        aWhere = AddFlag(Empty, "ITEM", grid1.TextMatrix(grid1.Row, 0))
        aWhere = AddFlag(aWhere, "DESCA", grid1.TextMatrix(grid1.Row, 1))
        aWhere = AddFlag(aWhere, "DATE1", xDate1.Text)
        aWhere = AddFlag(aWhere, "DATE2", xDate2.Text)
        ShowSalItem.aWhere = aWhere
'        cRepItem = grid1.TextMatrix(grid1.Row, 0)
'        DRepDate1 = xdate1.Text
'        DRepDate2 = xDate2.Text
        ShowSalItem.Show
    End If
End Sub
'Private Sub xDesca_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    FilterGrd grid1, xDesca.Text, 1
'End If
'End Sub
'Private Sub xitem_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then
'    ItemsLookupAll Me, oSearchItem
'End If
'End Sub
'Private Sub xITEM_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    FilterGrd grid1, xItem.Text, 0
'End If
'End Sub
'Sub myProc()
'xItem.Text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
'xDesca.Text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
'Unload oSearchItem
'End Sub
Private Sub Label2_Click(Index As Integer)

End Sub
Private Sub grid2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
ItemsLookupAll Me, osearchitem
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then Exit Sub
With grid2
If grid2.Row = grid2.Rows - 1 Then
    MyAddItem
End If
End With
End Sub
Private Sub grid2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If (Not validRow(OldRow, True, True)) And OldRow <> grid2.Rows - 1 And OldRow <> 0 And grid2.TextMatrix(OldRow, grid2.Cols - 1) = "" Then
    grid2.RemoveItem OldRow
    grid2.SaveGrid cFilesave, flexFileData
End If
End Sub
Private Sub grid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid2.Row <> grid2.Rows - 1 And grid2.Row <> 0 Then
    grid2.RemoveItem grid2.Row
    grid2.SaveGrid cFilesave, flexFileData
    grid2.Select grid2.Rows - 1, 1
    grid2.ShowCell grid2.Rows - 1, 1
End If
End Sub

Private Sub grid2_Validate(Cancel As Boolean)
If (Not validRow(grid2.Row, True, True)) And grid2.Row <> grid2.Rows - 1 And grid2.Row <> 0 And grid2.TextMatrix(grid2.Row, grid2.Cols - 1) = "" Then
    grid2.RemoveItem grid2.Row
    grid2.SaveGrid cFilesave, flexFileData
End If
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid2
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
End With
validRow = True
End Function
Sub myProc()
Dim nFound As Long
nFound = grid2.FindRow(osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 0), , 0)
If nFound > 0 Then
    MsgBox "«·’‰› „ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ " & nFound
    Exit Sub
End If
Dim bNew As Boolean
bNew = grid2.Row = grid2.Rows - 1
grid2.TextMatrix(grid2.Row, 0) = osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 0)
grid2.TextMatrix(grid2.Row, 1) = osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 1)
grid2_AfterEdit grid2.Row, 0
If Not bNew Then
    Unload osearchitem
Else
    grid2.Row = grid2.Rows - 1
End If
grid2.SaveGrid cFilesave, flexFileData
End Sub
Private Sub MyAddItem()
grid2.AddItem ""
grid2.ShowCell grid2.Rows - 1, 1
grid2.SaveGrid cFilesave, flexFileData
End Sub

Private Sub xDate1_LostFocus()
myValidDate xDate1
End Sub
Private Sub xdate2_LostFocus()
myValidDate xDate2
End Sub

