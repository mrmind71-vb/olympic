VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsImpItem 
   Caption         =   "„ «»⁄… «·«’‰«›"
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15240
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
   ScaleHeight     =   10365
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1860
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   -45
      Width           =   3930
      Begin VB.CommandButton CMD_R2 
         Caption         =   "ÿ»«⁄…  ≈Ã„«·Ï Õ—ﬂ… «’‰«› «·—”«·…"
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
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1440
         Width           =   3750
      End
      Begin VB.CommandButton FixCost 
         Caption         =   "÷»ÿ  ﬂ·›… «·«’‰«›"
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
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1035
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   3750
      End
      Begin VB.CommandButton CMD_R1 
         Caption         =   "ÿ»«⁄Ï —’Ìœ «·«’‰«›"
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
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   630
         Width           =   3750
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "⁄—÷"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton Cmd_Print 
         Caption         =   "ÿ»«⁄…"
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
         Left            =   1350
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmd_fix 
      Caption         =   "fix"
      Height          =   375
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   990
      Width           =   855
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   4725
      Top             =   900
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
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
      Left            =   4095
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   11085
      Begin VB.TextBox xDoc_No 
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
         Left            =   7605
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1815
      End
      Begin VB.TextBox XDATE2 
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
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   150
         Width           =   1815
      End
      Begin VB.TextBox xCode 
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
         Left            =   8355
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1065
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label xFactName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„” ‰œ"
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
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   300
         Width           =   960
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
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
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   690
         Width           =   570
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   4560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·—”«·…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5880
         TabIndex        =   4
         Top             =   270
         Width           =   1050
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   10035
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsImpItem.frx":0000
      Height          =   7860
      Left            =   150
      TabIndex        =   5
      Top             =   1845
      Width           =   14865
      _cx             =   26220
      _cy             =   13864
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
      BackColorSel    =   16777152
      ForeColorSel    =   8388608
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
      SelectionMode   =   1
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
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   225
      Left            =   45
      TabIndex        =   14
      Top             =   900
      Visible         =   0   'False
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «·—”«·…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   13710
      TabIndex        =   10
      Top             =   270
      Width           =   1050
   End
End
Attribute VB_Name = "VsImpItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim invTable As New ADODB.Recordset
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String

Private Sub CMD_FIX_Click()
    con.Execute " UPDATE FILE1_10 INNER JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM SET FILE6_20.price2 = [file1_10].[price] "
    con.Execute " UPDATE FILE1_10 INNER JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM SET FILE6_20.cost = [file1_10].[cost] "
    con.Execute " UPDATE FILE1_10 INNER JOIN FILE6_10 ON FILE1_10.ITEM = FILE6_10.ITEM SET FILE6_10.price2 = [file1_10].[price] "
    con.Execute " UPDATE FILE1_10 INNER JOIN FILE6_10 ON FILE1_10.ITEM = FILE6_10.ITEM SET FILE6_10.cost = [file1_10].[cost] "
End Sub
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰  ›’Ì·Ï √—’œ… «’‰«› —”«·… ≈” Ì—«œÌ… » «—ÌŒ " & xdate1.Text & " ··„Ê—œ  " & xCodeDesca.Caption
    
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 1, -2, cHead1, , , False, True, 10
    PrintGrd.Show 1
End Sub
Private Sub CMD_R1_Click()

If Not MYVALID Then Exit Sub
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT file1_10.package, Sum(FILE1_11.[IN]) - FILE1_11.[out]) AS Balance,FILE1_10.REORDER , FILE1_10.ITEM,FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP] AS FILE1_50GROUP,FILE1_10.[SECTION],FILE1_10.COST, file1_10.price , FILE1_50.DESCA AS FILE1_50DESCA,FILE1_50G.DESCA AS FILE1_50GDESCA,FILE1_10SC.DESCA AS FILE1_10SCDESCA" & _
            "FROM (((FILE1_10 LEFT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_50G ON FILE1_50.[GROUP] = FILE1_50G.CODE) LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE " & _
            " WHERE FILE1_10.ITEM IN (SELECT ITEM FROM FILE7_60 INNER JOIN FILE7_60H ON FILE7_60H.DOC_NO = FILE7_60.DOC_NO   WHERE FILE7_60H.doc_no = " & MyParn(xDoc_No.Text) & " ) "
cString = cString & " GROUP BY file1_10.package , FILE1_10.REORDER , FILE1_10.ITEM,file1_10.price , FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP],FILE1_10.[SECTION],FILE1_10.COST, FILE1_50.DESCA,FILE1_50G.DESCA,FILE1_10SC.DESCA"
          
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!Val10 = !Section
        temptable!str6 = !file1_10SCdesca
        temptable!val11 = !FILE1_50GROUP
        temptable!str7 = !file1_50GDESCA
        temptable!val12 = !Group
        temptable!str8 = !file1_50desca
        temptable!str1 = !Item
        temptable!str2 = !Desca
        temptable!val2 = !Balance
        temptable!val5 = !package
        temptable!Val3 = !price
        temptable!val4 = temptable!val2 * temptable!Val3
        temptable!str21 = "»Ì«‰  ›’Ì·Ï √—’œ… «’‰«› —”«·… ≈” Ì—«œÌ… » «—ÌŒ " & xdate1.Text & " " & XFACTNAME.Caption
        temptable.Update
      .MoveNext
    Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.REPORT1.ReportFileName = App.Path & "\Reports\Item1.rpt"
    main.REPORT1.DataFiles(0) = tempFile
    main.REPORT1.Action = 1
End If
temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub

Private Sub CMD_R2_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰  ›’Ì·Ï —”«·… —ﬁ„ " & xDoc_No.Text & "  » «—ÌŒ " & xdate1.Text & " " & XFACTNAME.Caption
    
    Load PrintGrd
    If bopt1 Then
        PrintGrd.doprint Me.grid1, 0.8, -2, cHead1, , , False, True, 7
    Else
        PrintGrd.doprint Me.grid1, 1, -2, cHead1, , , False, True, 8
    End If
    PrintGrd.Show 1

End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    If xDoc_No.Text <> "" Then myload
End Sub
Private Sub Form_Load()
    openCon con
'    invTable.Index = "nDate2"
    invTable.Open "FILE7_20", con, adOpenStatic, adLockReadOnly, adCmdTable
    
    
    Set grid1.DataSource = data1
    data1.ConnectionString = strCon
    xDate2.Text = Format(Date, "DD-MM-YYYY")
    FixGrid
    grid1.Rows = 1
End Sub
Private Sub myload()
Dim cwhere As String
If IsDate(xdate1.Text) Then
    cField5 = "0 AS F_BAL"
Else
    cwhere = " date < " & DateSq(xdate1.Text)
    cField5 = myiif(cwhere, "[IN] - [OUT]") & " AS F_BAL"
End If

If IsDate(xdate1.Text) Then cwhere = " date >= " & DateSq(xdate1.Text)
If IsDate(xDate2.Text) Then cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(xDate2.Text)

cField6 = myiif(cwhere & turn(cwhere, " And ") & " ( TYPE = '2' OR TYPE = '7' )", "[IN] - [OUT]") & " AS T_PURCH"

cField7 = myiif(cwhere & turn(cwhere, " And ") & "(TYPE = '6' OR TYPE = '3')", "[OUT] - [IN]") & " AS T_SALES"

cField8 = myiif(cwhere & turn(cwhere, " And ") & "(TYPE = '6' OR TYPE = '3')", "(FILE1_11.OUT - FILE1_11.[IN])* FILE1_11.PRICE * (1-(FILE1_11.DISCOUNT/100))") & " AS TV_SALES"

cField9 = myiif(cwhere & turn(cwhere, " And ") & "(TYPE = '6' OR TYPE = '3')", "(FILE1_11.OUT - FILE1_11.[IN] )* (FILE1_11.PRICE2) ") & " AS TV_PRICE"

cField14 = myiif(cwhere & turn(cwhere, " And ") & "(TYPE = '2' OR TYPE = '7')", "(FILE1_11.[IN] - FILE1_11.[OUT] )* FILE1_11.COST ") & " AS TCOST_PURCH"
 
cField15 = myiif(cwhere & turn(cwhere, " And ") & "(TYPE = '6' OR TYPE = '3')", "(FILE1_11.[IN] - FILE1_11.[OUT] )* FILE1_11.COST ") & " AS TCOST_SALES"

cField16 = myiif(cwhere, "(FILE1_11.[IN] - FILE1_11.[OUT] )* FILE1_11.COST ") & " AS TCOST_BAL"

If IsDate(xDate2.Text) Then cwhere = "date <= " & DateSq(xDate2.Text)
cField12 = myiif(cwhere, "[IN] - [OUT]") & " AS endbal "
 
With grid1
'                           0               1                 2                3                4
    cString = "  select file1_10.item , file1_10.desca , FILE1_10.PRICE , FILE1_10.PRICE2 , FILE1_10.COST  , " & _
                cField5 & " , " & cField6 & " , " & cField7 & " , " & cField8 & " , " & cField9 & ",  '  ' AS N10  , '  ' AS N11   , " & cField12 & " , '  ' AS  N13 , " & cField14 & " , " & cField15 & " , " & cField16 & _
                " from ( FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM )  inner join file1_50 on file1_10.[GROUP] = file1_50.code  " & _
                " WHERE FILE1_10.ITEM IN (SELECT ITEM FROM FILE7_60  INNER JOIN FILE7_60H ON FILE7_60H.DOC_NO = FILE7_60.DOC_NO   WHERE FILE7_60H.DOC_NO  = " & MyParn(xDoc_No.Text) & " ) "
    cString = cString & " GROUP BY FILE1_10.ITEM , FILE1_10.COST, FILE1_10.DESCA , FILE1_10.PRICE , FILE1_10.PRICE2 "
    data1.RecordSource = cString
    data1.Refresh
End With
FixGrid
End Sub
Sub FixGrid()
    With grid1
    .Cols = 20
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·’‰›"
    .TextMatrix(0, 2) = "”⁄— Ã„·…"
    .TextMatrix(0, 3) = "”⁄— ﬁÿ«⁄Ï"
    .TextMatrix(0, 4) = "”⁄—  ﬂ·›…"
    .TextMatrix(0, 5) = "—’Ìœ " & Format(xdate1.Text, "dd-mm-yyyy")
    .TextMatrix(0, 6) = "„‘ —Ì« "
    .TextMatrix(0, 7) = "„»Ì⁄« "
    .TextMatrix(0, 8) = "ﬁÌ„… „»Ì⁄«  ›⁄·Ì…"
    .TextMatrix(0, 9) = "ﬁÌ„… „»Ì⁄«  »”⁄— «·Ã„·…"
    .TextMatrix(0, 10) = "ﬁÌ„… Œ’„ „»Ì⁄« "
    .TextMatrix(0, 11) = "‰”»… «·Œ’„"
    .TextMatrix(0, 12) = "—’Ìœ " & Format(xDate2.Text, "dd-mm-yyyy")
    .TextMatrix(0, 13) = "‰”»… «·»Ì⁄"

    
    .TextMatrix(0, 14) = " ﬂ·›… „‘ —Ì« "
    .TextMatrix(0, 15) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 16) = " ﬂ·›… «·—’Ìœ"
    .TextMatrix(0, 17) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 18) = "‰”»… —»Õ „»Ì⁄« "
    .TextMatrix(0, 19) = "„Êﬁ› «·’‰›"
        
    .ColHidden(4) = Not bopt1
    .ColHidden(14) = Not bopt1
    .ColHidden(15) = Not bopt1
    .ColHidden(16) = Not bopt1
    .ColHidden(17) = Not bopt1
    .ColHidden(18) = Not bopt1
    .ColHidden(19) = Not bopt1
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 3000
    .ColWidth(2) = 800
    .ColWidth(3) = 800
    .ColWidth(4) = 800
    .ColWidth(5) = 700
    .ColWidth(6) = 700
    .ColWidth(7) = 900
    .ColWidth(8) = 1100
    .ColWidth(9) = 1000
    .ColWidth(10) = 900
    .ColWidth(11) = 800
    .ColWidth(12) = 700
    
    .ColWidth(13) = 1000
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    .ColWidth(16) = 1000
    .ColWidth(17) = 1000
    .ColWidth(18) = 1000
    .ColWidth(19) = 1000
    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(8) = flexDTDouble
    .ColDataType(9) = flexDTDouble
    .ColDataType(10) = flexDTDouble
    .ColDataType(11) = flexDTDouble
    
    .ColDataType(13) = flexDTDouble
    .ColDataType(14) = flexDTDouble
    .ColDataType(15) = flexDTDouble
    .ColDataType(16) = flexDTDouble
    .ColDataType(17) = flexDTDouble
    .ColDataType(18) = flexDTDouble
    .ColDataType(19) = flexDTDouble

    .ColFormat(2) = "#0.00"
    .ColFormat(3) = "#0.00"
    .ColFormat(4) = "#0.00"
    .ColFormat(5) = "#0"
    .ColFormat(6) = "#0"
    .ColFormat(7) = "#0"
    .ColFormat(8) = "#0"
    .ColFormat(9) = "#0"
    .ColFormat(10) = "#0"
    .ColFormat(11) = "#0"
    .ColFormat(12) = "#0"
    .ColFormat(13) = "#0"
    .ColFormat(14) = "#0"
    .ColFormat(15) = "#0"
    .ColFormat(16) = "#0"
    .ColFormat(17) = "#0"
    .ColFormat(18) = "#0"
    .ColFormat(19) = "#0"
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    For i = 1 To .Rows - 1
        .TextMatrix(i, 14) = Format(Val(.TextMatrix(i, 4)) * Val(.TextMatrix(i, 6)), "#0.00")
        
        .TextMatrix(i, 10) = Format(Val(.TextMatrix(i, 9)) - Val(.TextMatrix(i, 8)), "#0.00")
        If Val(.TextMatrix(i, 10)) <> 0 And Val(.TextMatrix(i, 9)) <> 0 Then .TextMatrix(i, 11) = Format(Val(.TextMatrix(i, 10)) / Val(.TextMatrix(i, 9)) * 100, "#0.00")
    
        If Val(.TextMatrix(i, 7)) <> 0 Then
            If (Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6))) <> 0 Then .TextMatrix(i, 13) = Format(Val(.TextMatrix(i, 7)) / (Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6))) * 100, "#0.00")
        End If
        
        .TextMatrix(i, 17) = Format(Val(.TextMatrix(i, 8)) - Val(.TextMatrix(i, 15)), "#0.00")
        If Val(.TextMatrix(i, 8)) <> 0 Then .TextMatrix(i, 18) = Format(Val(.TextMatrix(i, 17)) / Val(.TextMatrix(i, 8)) * 100, "#0.00")
        .TextMatrix(i, 19) = Format(Val(.TextMatrix(i, 8)) - Val(.TextMatrix(i, 14)), "#0.00")
        If Val(.TextMatrix(i, 19)) > 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H80FF80
        End If
        .TextMatrix(i, 16) = Format(.TextMatrix(i, 16), "#0")
'        LastSalTable.Filter = "item = " & MyParn(.TextMatrix(I, 0))
'        If Not LastSalTable.EOF Then .TextMatrix(I, 13) = Format(LastSalTable!m_date, "dd-mm-yyyy")
    Next i
    .SubtotalPosition = flexSTAbove
    
    .Subtotal flexSTSum, -1, 5, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 7, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 8, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 9, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 10, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 12, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 14, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 15, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 16, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 17, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 19, "#0", vbYellow, vbRed, True, "  "
    
    If .Rows >= 1 Then
        If Val(.TextMatrix(1, 10)) <> 0 And Val(.TextMatrix(1, 9)) <> 0 Then .TextMatrix(1, 11) = Format(Val(.TextMatrix(1, 10)) / Val(.TextMatrix(1, 9)) * 100, "#0.00")
        If Val(.TextMatrix(1, 7)) <> 0 Then
            If (Val(.TextMatrix(1, 5)) + Val(.TextMatrix(1, 6))) <> 0 Then .TextMatrix(1, 13) = Format(Val(.TextMatrix(1, 7)) / (Val(.TextMatrix(1, 5)) + Val(.TextMatrix(1, 6))) * 100, "#0.00")
        End If
        If Val(.TextMatrix(1, 8)) <> 0 Then .TextMatrix(1, 18) = Format(Val(.TextMatrix(1, 17)) / Val(.TextMatrix(1, 8)) * 100, "#0.00")
        
        .TextMatrix(1, 1) = "«·≈Ã„«·Ï"
    End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub grid1_dblClick()
    If grid1.Col <= 5 Then
        Load StoreMove
        StoreMove.xItem.Text = grid1.TextMatrix(grid1.Row, 0)
        StoreMove.Show
    Else
        cRepItem = grid1.TextMatrix(grid1.Row, 0)
        DRepDate1 = xdate1.Text
        DRepDate2 = xDate2.Text
        
        ShowSalItem.Show 1
    End If
End Sub
Private Sub xDOC_NO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(3, 1)
    
    Set Generalarray(0) = Me
    
    Generalarray(1) = "SELECT FILE7_60H.DOC_NO, file4_10.DESCA, FILE7_60H.date , FILE7_60H.FactName FROM file4_10 RIGHT JOIN FILE7_60H ON file4_10.CODE = FILE7_60H.code "
    Generalarray(2) = "Order by FILE7_60H.DATE DESC "
    Generalarray(3) = 5000
    Generalarray(5) = False
    
    listarray(0, 0) = "«·„Ê—œ - ≈”„ «·—”«·…"
    listarray(0, 1) = "(%%DESCA%% or %%FactName%%)"
    
    GrdArray(0, 0) = "—ﬁ„ „” ‰œ"
    GrdArray(0, 1) = 1500
    
    GrdArray(1, 0) = "«·«”„"
    GrdArray(1, 1) = 3000
    
    GrdArray(2, 0) = " «—ÌŒ"
    GrdArray(2, 1) = 1500
    
    GrdArray(3, 0) = "≈”„ «·—”«·…"
    GrdArray(3, 1) = 2500
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "≈” ⁄·«„ "
    Search3.Show 1
End If
End Sub
Sub myProc()
    xDoc_No.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    xdate1.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 2)
    xCodeDesca.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
    XFACTNAME.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 3)
    Unload Search3
Exit Sub
myerror:
Unload Search
End Sub
Private Sub xdate1_LostFocus()
'If IsDate(xDate1.Text) Then
'    Xcode.Text = GetDesca("select CODE  from FILE7_60H where DATE = " & DateSq(xDate1.Text))
'    xCodeDesca.Caption = GetDesca("select DESCA from FILE4_10 where CODE = " & MyParn(Xcode.Text))
'End If
End Sub

Private Sub FixCost_Click()
Dim datatable As New ADODB.Recordset
Dim nBalItem As Double, nCost As Double, i As Double
con.Execute "UPDATE FILE7_20 LEFT JOIN FILE7_20H ON FILE7_20.DOC_NO = FILE7_20H.DOC_NO SET FILE7_20.[DATE] = [FILE7_20H].[DATE] ,  FILE7_20.[imp_doc] = [FILE7_20H].[docimp] "
con.Execute "UPDATE FILE6_20 LEFT JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO SET FILE6_20.[DATE] = [FILE6_20H].[DATE] , FILE6_20.[rate] = [FILE6_20H].[rate] "
con.Execute "UPDATE FILE6_10 LEFT JOIN FILE6_10H ON FILE6_10.DOC_NO = FILE6_10H.DOC_NO SET FILE6_10.[DATE] = [FILE6_10H].[DATE] , FILE6_10.[rate] = [FILE6_10H].[rate] "


'datatable.Open "SELECT * FROM FILE6_20 ORDER BY ITEM", CON, adOpenKeyset, adLockPessimistic, adCmdText
'nRec = datatable.RecordCount
prog1.Visible = True
prog1.Value = 0
prog1.Value = 0
CalcImpCost
End Sub
Function CalcImpCost() As Double
    Dim nQ As Double
    Dim IMP_SALTable As New ADODB.Recordset
    Dim i As Double, nRec As Double
    Dim imp2table   As New ADODB.Recordset
    invTable.Requery
'    invTable.Index = "NDATE2"
    con.Execute " delete * from imp_sal "
    con.Execute "INSERT INTO IMP_SAL ( item, [date], quant, doc_sal  , price )  SELECT FILE6_20.ITEM, FILE6_20.DATE, FILE6_20.QUANT, FILE6_20.DOC_NO , file6_20.price * ((100 - file6_20.rate )/100) FROM FILE6_20  WHERE VAL(FILE6_20.QUANT & '') > 0 "
    con.Execute "INSERT INTO IMP_SAL ( item, [date], quant , doc_ret , price )  SELECT FILE6_10.ITEM, FILE6_10.DATE, FILE6_10.QUANT * -1, FILE6_10.DOC_NO , file6_10.price * ((100 - file6_10.rate )/100)  FROM FILE6_10  WHERE VAL(FILE6_10.QUANT & '') > 0 "
    With invTable
        invTable.MoveLast
        nRec = Val(GetDesca("select count(item) from FILE7_20") & "")
        .MoveFirst
        Do While Not .EOF
            i = i + 1
            prog1.Value = Round(i / nRec * 100)
            
            
            nQ = invTable!Quant
            If IMP_SALTable.State = adStateOpen Then IMP_SALTable.Close
                IMP_SALTable.Open "SELECT * FROM IMP_SAL WHERE ITEM = " & MyParn(!Item) & " AND VAL(QTY1 & '') + VAL(QTY2 & '') <> QUANT ORDER BY DATE , QUANT ", con, adOpenKeyset, adLockOptimistic
                If Not (IMP_SALTable.EOF And IMP_SALTable.BOF) Then
                    IMP_SALTable.MoveFirst
                    Do While Not IMP_SALTable.EOF
                        If Val(IMP_SALTable!Quant & "") > 0 Then
                            If Val(IMP_SALTable!Quant & "") <= nQ Then
                                If Val(IMP_SALTable!qty1 & "") = 0 Then
                                    If Val(IMP_SALTable!Quant & "") - nQ > 0 Then
                                        IMP_SALTable!qty1 = nQ
                                        IMP_SALTable!imp1 = invTable!IMP_DOC
                                        IMP_SALTable!cost = invTable!price
                                        IMP_SALTable.Update
                                        nQ = 0
                                    Else
                                        IMP_SALTable!qty1 = Val(IMP_SALTable!Quant & "")
                                        IMP_SALTable!imp1 = invTable!IMP_DOC
                                        IMP_SALTable!cost = invTable!price
                                        IMP_SALTable.Update
                                        nQ = nQ - Val(IMP_SALTable!Quant & "")
                                    End If
                                Else
                                    If Val(IMP_SALTable!qty1 & "") > 0 Then
                                        If (Val(IMP_SALTable!Quant & "") - Val(IMP_SALTable!qty1 & "")) - nQ > 0 Then
                                            IMP_SALTable!qty2 = nQ
                                            IMP_SALTable!imp2 = invTable!IMP_DOC
                                            IMP_SALTable!cost = ((IMP_SALTable!cost * IMP_SALTable!qty1) + (invTable!price * IMP_SALTable!qty2)) / IMP_SALTable!Quant
                                            IMP_SALTable.Update
                                            nQ = 0
                                        Else
                                            IMP_SALTable!qty2 = (Val(IMP_SALTable!Quant & "") - Val(IMP_SALTable!qty1 & ""))
                                            IMP_SALTable!imp2 = invTable!IMP_DOC
                                            IMP_SALTable!cost = ((IMP_SALTable!cost * IMP_SALTable!qty1) + (invTable!price * IMP_SALTable!qty2)) / IMP_SALTable!Quant
                                            IMP_SALTable.Update
                                            nQ = nQ - (Val(IMP_SALTable!Quant & "") - Val(IMP_SALTable!qty1 & ""))
                                        End If
                                    End If
                                End If
                            Else
                                IMP_SALTable!qty1 = nQ
                                IMP_SALTable!cost = invTable!price
                                IMP_SALTable!imp1 = invTable!IMP_DOC
                                IMP_SALTable.Update
                                nQ = 0
                            End If
                        End If
                        
                        
                        If Val(IMP_SALTable!Quant & "") < 0 Then
                            IMP_SALTable!qty1 = IMP_SALTable!Quant
                            IMP_SALTable!cost = invTable!price
                            IMP_SALTable!imp1 = invTable!IMP_DOC
                            IMP_SALTable.Update
                            nQ = nQ - IMP_SALTable!Quant
                        End If
                        
                        IMP_SALTable.MoveNext
                        If nQ <= 0 And Not IMP_SALTable.EOF Then
                            If IMP_SALTable!Quant > 0 Then Exit Do
                        End If
                    Loop
                End If
            
            .MoveNext
        Loop
    End With

    con.Execute " UPDATE FILE6_20 LEFT JOIN IMP_SAL ON (FILE6_20.ITEM = IMP_SAL.item) AND (FILE6_20.DOC_NO = IMP_SAL.doc_sal) SET FILE6_20.cost = [imp_sal].[cost] "
    con.Execute " UPDATE FILE6_10 LEFT JOIN IMP_SAL ON (FILE6_10.ITEM = IMP_SAL.item) AND (FILE6_10.DOC_NO = IMP_SAL.doc_ret) SET FILE6_10.cost = [imp_sal].[cost] "

    
    cStr1 = " select * from imp_sal where imp2 is not null "
    imp2table.Open cStr1, con, adOpenKeyset, adLockOptimistic, adCmdText
    With imp2table
        .MoveFirst
        Do While Not .EOF
            IMP_SALTable.AddNew
            IMP_SALTable!Item = !Item
            IMP_SALTable!imp1 = !imp2
            IMP_SALTable!qty1 = !qty2
            IMP_SALTable!doc_sal = !doc_sal
            IMP_SALTable!doc_ret = !doc_ret
            IMP_SALTable!Date = !Date
            IMP_SALTable!price = !price
            IMP_SALTable!cost = !cost
            IMP_SALTable.Update
            .MoveNext
        Loop
    End With
End Function




