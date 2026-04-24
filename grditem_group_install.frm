VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form grditemGroup_installfrm 
   Caption         =   " ﬁ—Ì— «·”œ«œ «·ÌÊ„Ì ··⁄÷ÊÌ«  «·„ﬁ”ÿ…"
   ClientHeight    =   10290
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   19065
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
   ScaleHeight     =   10290
   ScaleWidth      =   19065
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   1035
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   540
      Width           =   4830
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2325
         Picture         =   "grditem_group_install.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grditem_group_install.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3510
         Picture         =   "grditem_group_install.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "grditem_group_install.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   5895
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   90
      Width           =   9285
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   585
         Width           =   1815
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "D"
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ"
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
         Index           =   0
         Left            =   8280
         TabIndex        =   10
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   8295
         TabIndex        =   9
         Top             =   270
         Width           =   660
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   9960
      Width           =   19065
      _ExtentX        =   33629
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   90
      Top             =   270
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   4950
      Top             =   90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   8250
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1305
      Width           =   15090
      _cx             =   26617
      _cy             =   14552
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
      Cols            =   7
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
Attribute VB_Name = "grditemGroup_installfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public sDate1 As String, sdate2 As String
Dim aHeader()
Private Sub cmdExel_Click()
Dim sHeader As String, nMargin As Integer

sHeader = Me.Caption
nMargin = 30
If retHeader(aHeader, 0, 3, "-") <> "" Then
    sHeader = sHeader & turn(sHeader, Chr(13)) & retHeader(aHeader, 0, 3)
    nMargin = nMargin + 15
End If
If retHeader(aHeader, 1, 3, "-") <> "" Then
    sHeader = sHeader & turn(sHeader, Chr(13)) & retHeader(aHeader, 1, 3, "-")
    nMargin = nMargin + 15
End If


'Dim aSplit As Variant
'aSplit = AddFlag(aSplit, "title_col", "A:B")
'aSplit = AddFlag(aSplit, "title_row", "1:1")
'aSplit = AddFlag(aSplit, "center_header", sHeader)
ToFileExel grid1, , , , , , aRowWidth, arowHeight, aSplit, , , , nMargin

'ToFileExel grid1, , , , , 1
End Sub

Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
myload
End Sub
Private Sub cmdPrint_Click()
Dim aRow(0) As Variant, nRate As Double
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 4)
aRow(0) = AddFlag(aRow(0), "TEXT", "«·≈Ã„«·Ì")

Dim nwidth As Double
For I = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(I) Then
        nwidth = grid1.ColWidth(I) + nwidth
    End If
Next
nRate = 11500 / nwidth
PrintGrdNew.doprint grid1, nRate, -2, Me.Caption, retHeader(aHeader, 0, 1), retHeader(aHeader, 1, 2), , False, False, 11, , aRow
PrintGrdNew.Show 1
End Sub

Private Sub Form_Load()
openCon con


Set grid1.DataSource = DATA11
Fixgrd
If IsDate(sDate1) Then xDate1.text = sDate1
If IsDate(sdate2) Then xDate2.text = sdate2
If IsDate(sDate1) Or IsDate(sdate2) Then myload
End Sub
Private Sub myload()
With grid1
    Dim cString As String
    ReDim aHeader(2)
    cString = "SELECT FILE6_30H.FORM_NO,CASE WHEN FILE6_30H.TOTAL_VALUE <> 0 THEN '”œ«œ «ﬁ”«ÿ' ELSE 'ﬂ«—‰ÌÂ« ' END,FILE6_30H.CODE,FILE2_10.DESCA,FILE6_30H.TOTAL_VALUE + FILE6_30H.CARD_VALUE + FILE6_30H.INTEREST + FILE6_30H.OTHER" & _
              " FROM FILE6_30H INNER JOIN FILE2_10 ON FILE6_30H.CODE = FILE2_10.CODE WHERE IsFawry = 0"
    cString = cString & turn(cString) & "(NOT FILE6_30H.FORM_NO IS NULL)"
    
    cWhere = "FILE6_30H.TOTAL_VALUE + FILE6_30H.CARD_VALUE + FILE6_30H.INTEREST + FILE6_30H.OTHER  <> 0 "

    If IsDate(xDate1.text) Then
        cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_30H.DATE >= " & DateSq(xDate1.text)
        aHeader(0) = "⁄‰ «·› —… „‰ : " & BetweenString(xDate1.text, xDate2.text)
    End If
    
    If IsDate(xDate2.text) Then
        cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_30H.DATE <= " & DateSq(xDate2.text)
        aHeader(0) = "⁄‰ «·› —… „‰ : " & BetweenString(xDate1.text, xDate2.text)
    End If
    If cWhere <> "" Then cString = cString & " AND " & cWhere
    'cString = cString & " ORDER by FILE6_30H.FORM_NO"
    
    
    cString = cString & _
              " UNION ALL " & _
              " SELECT FILE6_30H.FORM_NO2 AS FORM_NO,'ﬁÌ„… „÷«›…' ,FILE6_30H.CODE,FILE2_10.DESCA,FILE6_30H.TOTAL_TAX" & _
              " FROM FILE6_30H INNER JOIN FILE2_10 ON FILE6_30H.CODE = FILE2_10.CODE  WHERE IsFawry = 0"
    cString = cString & turn(cString) & "(NOT FILE6_30H.FORM_NO2 IS NULL) AND (FILE6_30H.FORM_NO2 <> 0) "
    cWhere = "FILE6_30H.TOTAL_TAX <> 0"
    If IsDate(xDate1.text) Then
        cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_30H.DATE >= " & DateSq(xDate1.text)
        aHeader(0) = "⁄‰ «·› —… „‰ : " & BetweenString(xDate1.text, xDate2.text)
    End If
    
    If IsDate(xDate2.text) Then
        cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_30H.DATE <= " & DateSq(xDate2.text)
        aHeader(0) = "⁄‰ «·› —… „‰ : " & BetweenString(xDate1.text, xDate2.text)
    End If
    
    If cWhere <> "" Then cString = cString & " AND " & cWhere
    cString = cString & " ORDER by FILE6_30H.FORM_NO"
            
    Set DATA11.Recordset = myRecordSet(cString, con)
End With
Fixgrd
StatusBar1.Panels(1).text = IIf(grid1.rows - 2 > 0, "⁄œœ «·«” „«—«  : " & grid1.rows - 2, "")
End Sub
Sub Fixgrd()
    With grid1
    .ColWidth(0) = 1000
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 4000
    .ColWidth(4) = 1400
'    .ColWidth(4) = 1200
'    .ColWidth(5) = 1400
    
    .TextMatrix(0, 0) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 1) = "«·‰Ê⁄"
    .TextMatrix(0, 2) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 3) = "«·√”„"
    .TextMatrix(0, 4) = "«·≈Ã„«·Ì"
    '.TextMatrix(0, 4) = "«·÷—Ì»…"
    '.TextMatrix(0, 5) = "«·≈Ã„«·Ì"
    
    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    For I = 4 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0.00", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ì"
    Next
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set grditemGroup_installfrm = Nothing
closeCon con
End Sub
Private Function MYVALID() As Boolean
MYVALID = True
End Function

Private Sub xDate1_DblClick()
Set datefrm.oDate = xDate1
datefrm.Show 1
End Sub

Private Sub xdate2_DblClick()
Set datefrm.oDate = xDate2
datefrm.Show 1
End Sub

Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xgroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
