VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form trust_addfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«÷«›… »Ê«·’ «·‘Õ‰"
   ClientHeight    =   9495
   ClientLeft      =   15
   ClientTop       =   390
   ClientWidth     =   17955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   17955
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "‰ﬁ·«  »œÊ‰ »Ê«·’ ‘Õ‰"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   5355
      Width           =   17745
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Height          =   3210
         Left            =   90
         TabIndex        =   14
         Top             =   315
         Width           =   17565
         _cx             =   30983
         _cy             =   5662
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "‰ﬁ·«  »»Ê«·’ ‘Õ‰"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4290
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1035
      Width           =   17745
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   3930
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   17565
         _cx             =   30983
         _cy             =   6932
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
         Cols            =   9
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
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   11745
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   90
      Width           =   6000
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "D"
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox xDate_policy1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "D"
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox xDate_Policy2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·⁄„·Ì… „‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   540
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·»Ê·Ì’… „‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   1395
      End
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   1215
      Top             =   585
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
   Begin MSAdodcLib.Adodc data12 
      Height          =   330
      Left            =   2700
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
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   8055
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   360
      Width           =   3660
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1260
         MaskColor       =   &H00FFFFFF&
         Picture         =   "travel_trust.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmdFilter 
         Height          =   510
         Left            =   2430
         Picture         =   "travel_trust.frx":2363
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "travel_trust.frx":4855
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
   End
End
Attribute VB_Name = "trust_addfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sbox As String, sBox_Desca As String
Public myForm As Form
Dim con As New ADODB.Connection
Private Sub cmdFilter_Click()
myload1
myload2
End Sub

Private Sub CmdPrint_Click()
Dim cHead1 As String, cHead2 As String, cHead3 As String
cHead1 = "„ «»⁄… ›Ì“« Œ·«· › —…"
cHead2 = retHeader(aHeader, 0, 1)
cHead3 = retHeader(aHeader, 1, 1)
PrintGrd.doprint Me.grid1, 1, -2, cHead1, cHead2, cHead3, False, False, 8, , Array(1), Array(1)
PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set travel_addfrm = Nothing
End Sub
Private Sub CmdUndo_Click()
Unload Me
End Sub
Private Sub CmdPrint2_Click()
doprint
End Sub
Private Sub cmdSave_Click()
If Not MsgBox("≈÷«›… «·»Ê«·’ «·Ì ›« Ê—… «·„»Ì⁄«   ?", vbOKCancel + vbDefaultButton2) = vbOK Then Exit Sub
Dim NROWS As Long, nRow As Long
For i = 1 To grid1.Rows - 1
    If Val(grid1.TextMatrix(i, grid1.Cols - 1)) <> 0 Then
        NROWS = NROWS + 1
        Exit For
    End If
Next
If NROWS = 0 Then
    MsgBox "·«  ÊÃœ ”Ã·«  ·«÷«› Â«"
    Exit Sub
End If
myForm.Addproc
End Sub
Private Sub Form_Load()
openCon con

Set grid1.DataSource = DATA11
DATA11.ConnectionString = strCon

Set grid2.DataSource = DATA12
DATA12.ConnectionString = strCon

myload1
myload2
End Sub
Private Sub myload1()
Dim cÚString As String
With grid1
cString = "SELECT TRAVEL_H.DOC_NO,Convert(VARCHAR(10),TRAVEL_H.[DATE],111),FILE3_10.Desca,TRAVEL_H.POLICY," & _
          " Convert(VARCHAR(10),TRAVEL_H.[DATE_POLICY],111),TRAVEL_H.PLACE1,TRAVEL_H.PLACE2" & _
          " ,TRAVEL_BAL.TRUST,TRAVEL_BAL.CHARGE,TRAVEL_BAL.BALANCE,-1 AS CHOICE " & _
          " FROM  TRAVEL_H INNER JOIN FILE3_10 ON TRAVEL_H.CODE  = FILE3_10.CODE" & _
          " LEFT JOIN DRIVER ON TRAVEL_H.DRIVER = DRIVER.CODE" & _
          " LEFT JOIN TRAVEL_TRUST_ADDED ON (TRAVEL_H.DOC_NO = TRAVEL_TRUST_ADDED.DOC_NO AND BOX = " & MyParn(Sbox) & ")" & _
          " LEFT JOIN TRAVEL_BAL ON TRAVEL_BAL.DOC_NO = TRAVEL_H.DOC_NO" & _
          " Where (TRAVEL_TRUST_ADDED.DOC_NO Is Null)" & _
          " AND (NOT  TRAVEL_H.DATE_POLICY IS NULL) AND (NOT TRAVEL_H.POLICY IS NULL)" & _
          " AND TRAVEL_BAL.BOX = " & MyParn(Sbox)

If IsDate(xDate1.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE  >= " & DateSq(xDate1.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE <= " & DateSq(xDate2.Text)
End If

If IsDate(xDate_policy1.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE_POLICY  >= " & DateSq(xDate_policy1.Text)
End If

If IsDate(xDate_Policy2.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE_POLICY <= " & DateSq(xDate_Policy2.Text)
End If

cString = cString & " Order by Convert(VARCHAR(10),TRAVEL_H.[DATE],111),TRAVEL_H.DOC_NO"
DATA11.RecordSource = cString
DATA11.Refresh
fixgrd1
End With
End Sub
Private Sub myload2()
Dim cÚString As String
With grid1
cString = "SELECT TRAVEL_H.DOC_NO,Convert(VARCHAR(10),TRAVEL_H.[DATE],111),FILE3_10.Desca,TRAVEL_H.POLICY," & _
          " Convert(VARCHAR(10),TRAVEL_H.[DATE_POLICY],111),TRAVEL_H.PLACE1,TRAVEL_H.PLACE2" & _
          " ,TRAVEL_BAL.TRUST,TRAVEL_BAL.CHARGE,TRAVEL_BAL.BALANCE " & _
          "  FROM  TRAVEL_H INNER JOIN FILE3_10 ON TRAVEL_H.CODE  = FILE3_10.CODE" & _
          "  LEFT JOIN DRIVER ON TRAVEL_H.DRIVER = DRIVER.CODE" & _
          "  LEFT JOIN TRAVEL_TRUST_ADDED ON TRAVEL_H.DOC_NO = TRAVEL_TRUST_ADDED.DOC_NO" & _
          "  LEFT JOIN TRAVEL_BAL ON TRAVEL_BAL.DOC_NO = TRAVEL_H.DOC_NO" & _
          "  Where (TRAVEL_TRUST_ADDED.DOC_NO Is Null)" & _
          "  AND (TRAVEL_H.DATE_POLICY IS NULL) AND (TRAVEL_H.POLICY IS NULL)" & _
          "  AND TRAVEL_BAL.BOX = " & MyParn(Sbox)

If IsDate(xDate1.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE  >= " & DateSq(xDate1.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE <= " & DateSq(xDate2.Text)
End If

If IsDate(xDate_policy1.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE_POLICY  >= " & DateSq(xDate_policy1.Text)
End If

If IsDate(xDate_Policy2.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE_POLICY <= " & DateSq(xDate_Policy2.Text)
End If

cString = cString & " Order by Convert(VARCHAR(10),TRAVEL_H.[DATE],111),TRAVEL_H.DOC_NO"
DATA12.RecordSource = cString
DATA12.Refresh
fixgrd2
End With
End Sub
Sub fixgrd1()
With grid1
'.Cols = 11
.TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
.TextMatrix(0, 1) = " «—ÌŒ"
.TextMatrix(0, 2) = "«·⁄„Ì·"
.TextMatrix(0, 3) = "—ﬁ„ «·»Ê·Ì’…"
.TextMatrix(0, 4) = " «—ÌŒ «·»Ê·Ì’…"
.TextMatrix(0, 5) = "„‰"
.TextMatrix(0, 6) = "≈·Ì"
.TextMatrix(0, 7) = "«·⁄Âœ…"
.TextMatrix(0, 8) = "«·„’—Ê›"
.TextMatrix(0, 9) = "«·»«ﬁÌ"
.TextMatrix(0, 10) = "«Œ Ì«—"
.ColDataType(10) = flexDTBoolean
.ColWidth(0) = 1000
.ColWidth(1) = 1400
.ColWidth(2) = 2500
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.ColWidth(5) = 1500
.ColWidth(6) = 1500
.ColWidth(7) = 1000
.ColWidth(8) = 1000
.ColWidth(9) = 1000
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Sub fixgrd2()
With grid2
.TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
.TextMatrix(0, 1) = " «—ÌŒ"
.TextMatrix(0, 2) = "«·⁄„Ì·"
.TextMatrix(0, 3) = "—ﬁ„ «·»Ê·Ì’…"
.TextMatrix(0, 4) = " «—ÌŒ «·»Ê·Ì’…"
.TextMatrix(0, 5) = "„‰"
.TextMatrix(0, 6) = "≈·Ì"
.TextMatrix(0, 7) = "«·⁄Âœ…"
.TextMatrix(0, 8) = "«·„’—Ê›"
.TextMatrix(0, 9) = "«·»«ﬁÌ"
.ColWidth(0) = 1000
.ColWidth(1) = 1400
.ColWidth(2) = 2500
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.ColWidth(5) = 1500
.ColWidth(6) = 1500
.ColWidth(7) = 1000
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColHidden(3) = True
.ColHidden(4) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set travel_addfrm = Nothing
End Sub
Private Function doprint()
On Error GoTo myerror
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
     
If IsDate(xDate1.Text) Then aHeader(0) = BetweenString(xDate1.Text, xDate2.Text)
If IsDate(xDate2.Text) Then aHeader(0) = BetweenString(xDate1.Text, xDate2.Text)
If xbox.MatchedWithList Then aHeader(1) = turn(xbox.Text, "«·ﬂ«‘Ì— : ") & xbox.Text

With grid1
For i = 2 To grid1.Rows - 1
    
    temptable.AddNew
    temptable!str1 = TurnValue(ArbString(grid1.TextMatrix(i, 1)))
    temptable!str2 = TurnValue(grid1.TextMatrix(i, 2))
    temptable!val1 = Val(grid1.TextMatrix(i, 5))
    temptable!str12 = Format(grid1.TextMatrix(i, 1), "yyyy-mm-dd")
    temptable!str11 = arbDay(grid1.TextMatrix(i, 1)) & turn(grid1.TextMatrix(i, 1), " ") & Format(grid1.TextMatrix(i, 1), "yyyy/mm/dd")
    temptable!str21 = TurnValue(aHeader(0))
    If Trim(aHeader(1)) <> "" Then temptable!str21 = temptable!str21 & turn(temptable!str21, vbCrLf) & aHeader(1)
    temptable.Update
Next
End With
temptable.Requery
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    temptable.Requery
    main.Report1.Reset
    'Report1.PrinterCopies = nCopies
    'Report1.Destination = crptToPrinter
    main.Report1.ReportFileName = App.Path & "\Reports\visa1.rpt"
    main.Report1.DataFiles(0) = tempFile
    main.Report1.Action = 1
End If
doprint = True
closeCon:
temptable.Close
Set temptable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GoTo closeCon
End Function

Private Sub Grid1_EnterCell()
If grid1.Col = 10 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub

Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
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
Private Sub xDate_policy1_GotFocus()
myGotFocus xDate_policy1
End Sub
Private Sub xDate_policy1_LostFocus()
myLostFocus xDate_policy1
myValidDate xDate_policy1
End Sub
Private Sub xDate_Policy2_GotFocus()
myGotFocus xDate_Policy2
End Sub
Private Sub xDate_Policy2_LostFocus()
myLostFocus xDate_Policy2
myValidDate xDate_Policy2
End Sub
