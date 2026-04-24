VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ClientMove 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… «·⁄„·«¡"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12960
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
   ScaleHeight     =   6825
   ScaleWidth      =   12960
   WindowState     =   2  'Maximized
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9000
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1350
      Width           =   1515
   End
   Begin VB.TextBox xClient 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9000
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   900
      Width           =   1515
   End
   Begin VB.CommandButton Cmdgo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "⁄—÷ Õ—ﬂ… «·Õ”«»"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5175
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1275
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   540
      ScaleWidth      =   12960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   12960
      Begin VB.CommandButton CmdApply 
         Caption         =   "ÿ»«⁄… ﬂ‘› Õ”«» œ› —Ï"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   45
         Width           =   2490
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÿ»«⁄… ﬂ‘› Õ”«» ‰ﬁœÏ"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   45
         Width           =   2490
      End
      Begin VB.CommandButton CmdFix 
         BackColor       =   &H00C0FFFF&
         Caption         =   "÷»ÿ «·Õ—ﬂ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9975
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         Width           =   1590
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Œ—ÊÃ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   150
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   75
         Width           =   1365
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
      TabIndex        =   3
      Top             =   1920
      Width           =   405
   End
   Begin Crystal.CrystalReport Report1 
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
   Begin VSFlex7LCtl.VSFlexGrid InvGrid 
      Height          =   6345
      Left            =   300
      TabIndex        =   12
      Top             =   1920
      Width           =   11340
      _cx             =   20002
      _cy             =   11192
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   16777215
      BackColorSel    =   13292191
      ForeColorSel    =   255
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
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
      GridLineWidth   =   2
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Clmove.frx":0000
      ScrollTrack     =   -1  'True
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
      OutlineCol      =   1
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
      WordWrap        =   -1  'True
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
      WallPaperAlignment=   4
   End
   Begin VB.Label xDateInv 
      Caption         =   "Label4"
      Height          =   390
      Left            =   450
      TabIndex        =   16
      Top             =   825
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰  «—ÌŒ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10650
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1350
      Width           =   780
   End
   Begin VB.Label xInv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   2550
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   900
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label xType 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1575
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H006FA877&
      BorderWidth     =   3
      Height          =   1065
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   750
      Width           =   11625
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10650
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   900
      Width           =   330
   End
   Begin VB.Label xDesca 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5175
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   900
      Width           =   3165
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "≈”„"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   900
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H006FA877&
      BorderWidth     =   3
      Height          =   6645
      Left            =   150
      Top             =   1800
      Width           =   11625
   End
End
Attribute VB_Name = "ClientMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClientTable As Recordset, DocTable As Recordset
Public SubBalTable As Recordset
Dim FBAL2Table As Recordset
Dim FBAL3Table As Recordset
Dim datatable As Recordset
Dim CashTble As Recordset
Dim nCash1 As Double
Dim nCash2 As Double
Dim nFBal2 As Double
Dim nFBal3 As Double
Dim nCrow As Double
Dim lCardCust As Boolean
Sub fillgrd()
Dim nBal1 As Double
Dim nBal2 As Double
nBal1 = 0
nBal2 = 0
If DocTable.RecordCount <> 0 Then DocTable.MoveFirst
InvGrid.Rows = 1
i = 1
If DocTable.RecordCount > 0 Then
nBal1 = nFBal2
nBal2 = nFBal3

If nFBal2 <> 0 Or nFBal3 <> 0 Then
    nBal1 = nFBal2
    nBal2 = nFBal3
    InvGrid.AddItem ""
    InvGrid.TextMatrix(i, 0) = "—’Ìœ √Ê·"
    InvGrid.TextMatrix(i, 5) = Format(nBal1, "fixed")
    InvGrid.TextMatrix(i, 6) = Format(nBal2, "fixed")
    i = 1 + i
End If

Do
    If DocTable.Show = "1" Or DocTable.Show = "2" Then nBal1 = nBal1 + TurnValue(DocTable.sal, Null, 0) - TurnValue(DocTable.PAY, Null, 0)
    If DocTable.Show = "1" Or DocTable.Show = "3" Then nBal2 = nBal2 + TurnValue(DocTable.sal, Null, 0) - TurnValue(DocTable.PAY, Null, 0)
    InvGrid.AddItem ""
    If DocTable!Type <> "8" And DocTable!Type <> "9" Then
        InvGrid.TextMatrix(i, 0) = TurnValue(DocTable.desca, Null, "") & TurnValue(DocTable.Doc_Id, Null, "")
    Else
        InvGrid.TextMatrix(i, 0) = TurnValue(DocTable.desca, Null, "")
    End If
    InvGrid.TextMatrix(i, 1) = IIf(IsDate(DocTable!Date), Format(DocTable!Date, "dd-mm-yyyy"), "")
    InvGrid.TextMatrix(i, 2) = TurnValue(DocTable.Doc_Id, Null, "")
    InvGrid.TextMatrix(i, 3) = TurnValue(DocTable.PAY, Null, "")
    InvGrid.TextMatrix(i, 4) = TurnValue(DocTable.sal, Null, "")
    InvGrid.TextMatrix(i, 5) = Format(nBal1, "fixed")
    InvGrid.TextMatrix(i, 6) = Format(nBal2, "fixed")
    InvGrid.TextMatrix(i, 7) = DocTable!Type
    DocTable.MoveNext
    i = i + 1
Loop Until DocTable.EOF
End If
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
xDescA.Caption = GrdText(Search.Grid1, 1)
Unload Search
End Sub
Private Sub CmdApply_Click()
    Dim InvTable As Recordset
    Dim RInvTable As Recordset
'    Dim cHead1 As String
'    cHead1 = " ›’Ì·Ï ﬂ‘› Õ”«» «·⁄„Ì· " & Me.xDesca.Caption
'    Load PrintGrd
'    PrintGrd.Doprint InvGrid, 1, -2, cHead1, , , False, False, 10
'    PrintGrd.Show 1

Dim datatable As Recordset
Dim nBalance  As Double


If publicFlag = 1 Then
    cString = "Select * FROM FILE3_11  " & _
              " Where ( file3_11.SHOW = '1' or file3_11.SHOW = '2' ) AND File3_11.Code = " & MyParn(xclient.Text)
    If IsDate(xDate.Text) Then cString = cString & " AND FILE3_11.DATE >= " & DateSql(xDate.Text)
    cString = cString & " Order By Date,SAL "
Else
    cString = "Select * FROM FILE4_11  " & _
              " Where ( file4_11.SHOW = '1' or file4_11.SHOW = '2' ) AND File4_11.Code = " & MyParn(xclient.Text)
    If IsDate(xDate.Text) Then cString = cString & " AND FILE4_11.DATE >= " & DateSql(xDate.Text)
    cString = cString & " Order By Date,SAL "

End If
Set datatable = mydb.OpenRecordset(cString, dbOpenSnapshot)

If publicFlag = "2" Then

    cString = "Select total , doc_no FROM FILE7_20   " & _
              " Where file7_20.store = 'zz' and File7_20.Code = " & MyParn(xclient.Text)
    If IsDate(xDate.Text) Then cString = cString & " AND FILE7_20.DATE >= " & DateSql(xDate.Text)
    Set InvTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

    cString = "Select total , doc_no FROM FILE6_11   " & _
              " Where file6_11.store = 'zz' and File6_11.Code = " & MyParn(xclient.Text)
    If IsDate(xDate.Text) Then cString = cString & " AND FILE6_11.DATE >= " & DateSql(xDate.Text)
    Set RInvTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

End If
tempdb.Execute " DELETE * FROM TEMP "
Set TargetTable = tempdb.OpenRecordset("TEMP")

nBalance = 0
If nFBal3 <> 0 Then
    TargetTable.AddNew
    nBalance = nFBal3
    TargetTable.str1 = "—’Ìœ ”«»ﬁ"
    TargetTable.VAL6 = nBalance
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle
    TargetTable.Update
End If

With datatable
If .RecordCount > 0 Then
    Do While True
        TargetTable.AddNew
        nBalance = nBalance + TurnValue(.sal, Null, 0) - TurnValue(.PAY, Null, 0)
        If !Type = 1 Then
            TargetTable.str1 = "—’Ìœ √Ê·"
        Else
            TargetTable.str2 = .Doc_Id
        End If
        TargetTable.str8 = " ﬂ‘› Õ”«» " & xDescA.Caption
        TargetTable.str1 = .[desca] & " " & .Doc_Id
        If publicFlag = 2 Then
            If !Type = "4" Then
                InvTable.FindFirst " DOC_NO  =   " & MyParn(.Doc_Id)
                If Not InvTable.NoMatch Then TargetTable.VAL10 = InvTable.total * -1
            End If
            
            If !Type = "5" Then
                RInvTable.FindFirst " DOC_NO  =   " & MyParn(.Doc_Id)
                If Not RInvTable.NoMatch Then TargetTable.VAL10 = RInvTable.total
            End If
        End If
        If !Type = "4" Then TargetTable.VAL1 = .sal
        If !Type = "5" Then TargetTable.VAL2 = .PAY
        If !Type = "7" Then TargetTable.VAL3 = .PAY
        If !Type = "8" Then TargetTable.VAL4 = .PAY
        If !Type = "0" Then TargetTable.VAL5 = .PAY
        ClientTable.FindFirst "Code = " & MyParn(xclient.Text)
        TargetTable.VAL9 = ClientTable.DISC
        TargetTable.VAL6 = nBalance
        TargetTable.Date1 = .[Date]
        TargetTable.STR19 = firsttitle
        ' TargetTable.str20 = Secondtitle
        TargetTable.Update
        .MoveNext
        If .EOF Then Exit Do
    Loop
End If
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\rep_203.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End With
End Sub
Private Sub cmdFix_Click()
If publicFlag = 1 Then
    mydb.Execute "DELETE * FROM FILE3_11"
    cString = "INSERT INTO FILE3_11( " & _
           "[Type],Code,[Date],Sal,DescA,SHOW)" & _
           " Select '1',Code,f_Date,f_Bal1,'—’Ìœ √Ê·' , '2' " & _
           " From File3_10 where file3_10.f_Bal1 <> 0 "
    mydb.Execute cString
    
    cString = "INSERT INTO FILE3_11( " & _
           "[Type],Code,[Date],Sal,DescA,SHOW)" & _
           " Select '1',Code,f_Date,f_Bal2,'—’Ìœ √Ê·' , '3' " & _
           " From File3_10  where file3_10.f_Bal1 <> 0 "
    mydb.Execute cString
    
    cString = "Insert Into File3_11(" & _
              "[Type],Doc_Id, Code,[Date],Sal,DescA, SHOW )" & _
              " Select '4',Doc_No,FILE6_20.Code,[Date],Sum(total),'„»Ì⁄« ' , '1' From File6_20 LEFT JOIN FILE3_10 ON File6_20.CODE = FILE3_10.CODE WHERE NOT FILE3_10.CASH " & _
              " Group by Doc_No,FILE6_20.Code,[Date]"
    mydb.Execute cString
    
    cString = "Insert Into File3_11(" & _
              "[Type],Doc_Id, Code,[Date],Pay,DescA, SHOW )" & _
              " Select '5',Doc_No,FILE6_10.Code,[Date],Sum(total),'„—œÊœ „»Ì⁄« ' , '1' FROM File6_10 LEFT JOIN FILE3_10 ON File6_10.CODE = FILE3_10.CODE WHERE NOT FILE3_10.CASH " & _
              " Group by Doc_No,FILE6_10.Code,[Date]"
    mydb.Execute cString
    
    cString = "Insert Into File3_11(" & _
              "[Type],Doc_Id,Code,[Date],Pay,DescA,SHOW )" & _
              " Select '7',Doc_No,Code,[Date],[Value],'œ›⁄…' , '1' " & _
              " From File8_10"
    mydb.Execute cString

    cString = "Insert Into File3_11(" & _
              "[Type],Doc_Id,Code,[Date],Pay,DescA,SHOW )" & _
              " Select '0',Doc_No,Code,[Date],[Value],' ”ÊÌ…' , '1' " & _
              " From File8_20"
    mydb.Execute cString

    cString = "Insert Into File3_11(" & _
              "[Type],Doc_Id,Code,[Date],Pay,DescA,SHOW )" & _
              " Select '8',SER_NO,Code,[Date_R],[Value],'‘Ìﬂ Õﬁ' &  FORMAT(DATE_1,'dd-MM-yyyy'), '2' " & _
              " From File5_20"
    mydb.Execute cString

    cString = "INSERT INTO FILE3_11 ( PAY, TYPE, SHOW, DOC_ID, CODE, [DATE], DESCA )" & _
              "SELECT file5_22.VALUE, '9' AS Expr1, 3 AS Expr2, file5_22.SER_NO, FILE5_20.CODE, FILE5_22.[DATE],'œ›⁄… ‘Ìﬂ' &  FORMAT(DATE_1,'dd-MM-yyyy') " & _
              " FROM FILE5_20 RIGHT JOIN file5_22 ON FILE5_20.SER_NO = file5_22.ser_no "
    mydb.Execute cString

'    FIXCHQ1
Else
    mydb.Execute "DELETE * FROM FILE4_11"
    
    cString = "INSERT INTO FILE4_11( " & _
           "[Type],Code,[Date],Sal,DescA,SHOW)" & _
           " Select '1',Code,f_Date,f_Bal1,'—’Ìœ √Ê·' , '2' " & _
           " From File4_10  where file4_10.f_Bal1 <> 0 "
    mydb.Execute cString
    
    cString = "INSERT INTO FILE4_11( " & _
           "[Type],Code,[Date],Sal,DescA,SHOW)" & _
           " Select '1',Code,f_Date,f_Bal2,'—’Ìœ √Ê·' , '3' " & _
           " From File4_10  where file4_10.f_Bal1 <> 0 "
    mydb.Execute cString
    
    cString = "Insert Into File4_11(" & _
              "[Type],Doc_Id, Code,[Date],Sal,DescA, SHOW )" & _
              " Select '4',Doc_No,Code,[Date],Sum(total),'„‘ —Ì« ' , '1' From File7_20 " & _
              " Group by Doc_No,Code,[Date]"
    mydb.Execute cString
    
    cString = "Insert Into File4_11(" & _
              "[Type],Doc_Id, Code,[Date],Pay,DescA, SHOW )" & _
              " Select '5',Doc_No,Code,[Date],Sum(total),'„—œÊœ „‘ —Ì« ' , '1' From File6_11 " & _
              " Group by Doc_No,Code,[Date]"
    mydb.Execute cString
    
    cString = "Insert Into File4_11(" & _
              "[Type],Doc_Id,Code,[Date],Pay,DescA,SHOW )" & _
              " Select '0',Doc_No,Code,[Date],[Value],' ”ÊÌ…' , '1' " & _
              " From File8_40"
    mydb.Execute cString

    cString = "Insert Into File4_11(" & _
              "[Type],Doc_Id,Code,[Date],Pay,DescA,SHOW )" & _
              " Select '7',Doc_No,Code,[Date],[Value],'œ›⁄…' , '1' " & _
              " From File8_30"
    mydb.Execute cString

    cString = "Insert Into File4_11(" & _
              "[Type],Doc_Id,Code,[Date],Pay,DescA,SHOW )" & _
              " Select '8',SER_NO,Code,[Date_R],[Value],'‘Ìﬂ Õﬁ' &  FORMAT(DATE_1,'dd-MM-yyyy'), '2' " & _
              " From File5_21"
    mydb.Execute cString

    cString = "INSERT INTO FILE4_11 ( PAY, TYPE, SHOW, DOC_ID, CODE, [DATE], DESCA )" & _
              "SELECT file5_23.VALUE, '9' AS Expr1, 3 AS Expr2, file5_23.SER_NO, FILE5_21.CODE, FILE5_23.[DATE],'œ›⁄… ‘Ìﬂ' &  FORMAT(DATE_1,'dd-MM-yyyy') " & _
              " FROM FILE5_21 RIGHT JOIN file5_23 ON FILE5_21.SER_NO = file5_23.ser_no "
    mydb.Execute cString

'    FIXCHQ2
End If
End Sub
Private Sub CmdGo_Click()
Dim cStr2 As String
Dim cStr3 As String

If publicFlag = 1 Then
    cString = " SELECT Sum(FILE3_11.PAY) AS fPAY, Sum(FILE3_11.SAL) AS fSAL FROM FILE3_11 " & _
              " Where File3_11.Code = " & MyParn(xclient.Text)
    If IsDate(xDate.Text) Then cString = cString & " AND FILE3_11.DATE < " & DateSql(xDate.Text)
    cStr2 = cString & " AND ( FILE3_11.SHOW = '2' OR     FILE3_11.SHOW = '1' )"
    cStr3 = cString & " AND ( FILE3_11.SHOW = '3' OR     FILE3_11.SHOW = '1' )"
    
    Set FBAL2Table = mydb.OpenRecordset(cStr2)
    Set FBAL3Table = mydb.OpenRecordset(cStr3)
    nFBal2 = 0
    nFBal3 = 0
    If IsDate(xDate.Text) Then
        FBAL2Table.MoveFirst
        nFBal2 = TurnValue(FBAL2Table.FSAL, Null, 0) - TurnValue(FBAL2Table.FPAY, Null, 0)
        
        FBAL3Table.MoveFirst
        nFBal3 = TurnValue(FBAL3Table.FSAL, Null, 0) - TurnValue(FBAL3Table.FPAY, Null, 0)
    End If
    cQryStr = "Select * from File3_11 Where Code = " & MyParn(xclient.Text)
    If IsDate(xDate.Text) Then cQryStr = cQryStr & " AND FILE3_11.DATE >= " & DateSql(xDate.Text)
Else
    cString = " SELECT Sum(FILE4_11.PAY) AS fPAY, Sum(FILE4_11.SAL) AS fSAL FROM FILE4_11 " & _
              " Where File4_11.Code = " & MyParn(xclient.Text)
    If IsDate(xDate.Text) Then cString = cString & " AND FILE4_11.DATE < " & DateSql(xDate.Text)
    cStr2 = cString & " AND ( FILE4_11.SHOW = '2' OR     FILE4_11.SHOW = '1' )"
    cStr3 = cString & " AND ( FILE4_11.SHOW = '3' OR     FILE4_11.SHOW = '1' )"
    
    Set FBAL2Table = mydb.OpenRecordset(cStr2)
    Set FBAL3Table = mydb.OpenRecordset(cStr3)
    nFBal2 = 0
    nFBal3 = 0
    If IsDate(xDate.Text) Then
        FBAL2Table.MoveFirst
        nFBal2 = TurnValue(FBAL2Table.FSAL, Null, 0) - TurnValue(FBAL2Table.FPAY, Null, 0)
        
        FBAL3Table.MoveFirst
        nFBal3 = TurnValue(FBAL3Table.FSAL, Null, 0) - TurnValue(FBAL3Table.FPAY, Null, 0)
    End If
    cQryStr = "Select * from File4_11 Where Code = " & MyParn(xclient.Text)
    If IsDate(xDate.Text) Then cQryStr = cQryStr & " AND FILE4_11.DATE >= " & DateSql(xDate.Text)
End If
cQryStr = cQryStr & " Order by [Date],PAY"
Set DocTable = mydb.CreateSnapshot(cQryStr)
fillgrd
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Command1_Click()
Dim datatable As Recordset
Dim nBalance  As Double

If publicFlag = 1 Then
    cString = "Select * FROM FILE3_11  " & _
              " Where ( file3_11.SHOW = '1' or file3_11.SHOW = '3' ) AND File3_11.Code = " & MyParn(xclient.Text)
    If IsDate(xDate.Text) Then cString = cString & " AND FILE3_11.DATE >= " & DateSql(xDate.Text)
    cString = cString & " Order By Date,SAL "
Else
    cString = "Select * FROM FILE4_11  " & _
              " Where ( file4_11.SHOW = '1' or file4_11.SHOW = '3' ) AND File4_11.Code = " & MyParn(xclient.Text)
    If IsDate(xDate.Text) Then cString = cString & " AND FILE4_11.DATE >= " & DateSql(xDate.Text)
    cString = cString & " Order By Date,SAL "
End If
Set datatable = mydb.OpenRecordset(cString, dbOpenSnapshot)
tempdb.Execute " DELETE * FROM TEMP "
Set TargetTable = tempdb.OpenRecordset("TEMP")

nBalance = 0
If nFBal2 <> 0 Then
    TargetTable.AddNew
    nBalance = nFBal2
    TargetTable.str1 = "—’Ìœ ”«»ﬁ"
    TargetTable.VAL6 = nBalance
    TargetTable.STR19 = firsttitle
    ' TargetTable.str20 = Secondtitle
    TargetTable.Update
End If

With datatable
If .RecordCount > 0 Then
    Do While True
        TargetTable.AddNew
        nBalance = nBalance + TurnValue(.sal, Null, 0) - TurnValue(.PAY, Null, 0)
        If !Type = 1 Then
            TargetTable.str1 = "—’Ìœ √Ê·"
        Else
            TargetTable.str2 = .Doc_Id
        End If
        TargetTable.str8 = " ﬂ‘› Õ”«» " & xDescA.Caption
        TargetTable.str1 = .[desca]
        
        If !Type = "4" Then TargetTable.VAL1 = .sal
        If !Type = "5" Then TargetTable.VAL2 = .PAY
        If !Type = "7" Then TargetTable.VAL3 = .PAY
        If !Type = "9" Then TargetTable.VAL4 = .PAY
        If !Type = "0" Then TargetTable.VAL5 = .PAY
        TargetTable.VAL6 = nBalance
        TargetTable.Date1 = .[Date]
        TargetTable.STR19 = firsttitle
        ' TargetTable.str20 = Secondtitle
        TargetTable.Update
        .MoveNext
        If .EOF Then Exit Do
    Loop
End If
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\R_203.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End With
End Sub
Private Sub Form_Load()
CmdApply.Visible = bopt3
Command1.Visible = bopt3
If publicFlag = 1 Then
    Set ClientTable = mydb.OpenRecordset("file3_10", dbOpenSnapshot)
    Me.Caption = "Õ—ﬂ… «·⁄„·«¡"
Else
    Set ClientTable = mydb.OpenRecordset("file4_10", dbOpenSnapshot)
    Me.Caption = "Õ—ﬂ… «·„Ê—œÌ‰"
End If
With InvGrid
.Cols = 9
.TextMatrix(0, 0) = "»Ì«‰ «·Õ—ﬂ…"
.TextMatrix(0, 1) = " «—ÌŒ"
.TextMatrix(0, 2) = "—ﬁ„ Õ«”»"
.TextMatrix(0, 3) = "„œÌ‰"
.TextMatrix(0, 4) = "œ«∆‰"
.TextMatrix(0, 6) = "—’Ìœ ‰ﬁœÏ"
.TextMatrix(0, 5) = "—’Ìœ œ› —Ï"
.TextMatrix(0, 7) = "„” ‰œ"

.ColWidth(0) = 3000
.ColWidth(1) = 1500
.ColWidth(2) = 0
.ColWidth(3) = 1300
.ColWidth(4) = 1200
.ColWidth(5) = 1200
.ColWidth(6) = 1200
.ColWidth(7) = 0
.ColWidth(8) = 0
End With

End Sub
Private Sub invGrid_DblClick()
    xType.Caption = InvGrid.TextMatrix(InvGrid.Row, 7)
    xInv.Caption = InvGrid.TextMatrix(InvGrid.Row, 2)
    xDateInv.Caption = InvGrid.TextMatrix(InvGrid.Row, 1)
    If xType.Caption = "4" Or xType.Caption = "5" Or xType.Caption = "0" Then ViewInv.Show 1
End Sub

Private Sub xClient_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(3)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    If publicFlag = 1 Then
        Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê From File3_10"
    Else
        Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê From File4_10"
    End If
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 3000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo) Then SendKeys "{tAB}"
'If KeyAscii = 13 Then
'    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
'End If
End Sub
Private Sub xClient_LostFocus()
    ClientTable.FindFirst "Code = " & MyParn(xclient.Text)
    xDescA.Caption = IIf(ClientTable.NoMatch, "", TurnValue(ClientTable.desca, Null, ""))
    lCardCust = False
End Sub
