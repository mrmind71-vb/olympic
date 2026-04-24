VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form grdVisafrm2 
   Caption         =   "„ «»⁄… ›Ì“« ÌÊ„Ì…"
   ClientHeight    =   9705
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   12825
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
   ScaleHeight     =   9705
   ScaleWidth      =   12825
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "„ «»⁄… ‘Â—Ì…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8430
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1215
      Width           =   6180
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·«‘Â— »œÊ‰ ›Ì“«"
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
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   8010
         Width           =   2940
      End
      Begin VB.CommandButton cmdprint11 
         Enabled         =   0   'False
         Height          =   510
         Left            =   90
         Picture         =   "grdvisa2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7830
         Width           =   1590
      End
      Begin VSFlex7Ctl.VSFlexGrid GRID11 
         Height          =   7530
         Left            =   135
         TabIndex        =   13
         Top             =   270
         Width           =   6000
         _cx             =   10583
         _cy             =   13282
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
         Cols            =   6
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
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   3105
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   540
      Width           =   2490
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   1260
         Picture         =   "grdvisa2.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "grdvisa2.frx":5483
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   225
      Width           =   6855
      Begin VB.TextBox xDate1 
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1635
      End
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   135
         TabIndex        =   2
         Top             =   225
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬂ«‘Ì— :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ :"
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
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
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
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   1215
      Top             =   630
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
   Begin MSAdodcLib.Adodc DATA20 
      Height          =   330
      Left            =   1755
      Top             =   270
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
      Left            =   2790
      Top             =   495
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
   Begin VB.Frame Frame3 
      Caption         =   "„ «»⁄… ÌÊ„Ì…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8430
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1215
      Width           =   6180
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   510
         Left            =   90
         Picture         =   "grdvisa2.frx":78EF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7830
         Width           =   1455
      End
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   7485
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   6000
         _cx             =   10583
         _cy             =   13203
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
         Cols            =   6
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
End
Attribute VB_Name = "grdVisafrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim aHeader(1)
Dim oSearchClient As New Search3
Dim cStr1 As String, cStr2 As String
Private Sub CmdPrint_Click()
Dim cHead1 As String, cHead2 As String, cHead3 As String
cHead1 = "„ «»⁄… ÌÊ„Ì… ··›Ì“«"
cHead2 = retHeader(aHeader, 0, 1)
cHead3 = retHeader(aHeader, 1, 1)
PrintGrd.doprint grid1, 1, -2, cHead1, cHead2, cHead3, False, False, 8, , Array(1), Array(1)
PrintGrd.Show 1
End Sub
Private Sub cmdPrint11_Click()
Dim cHead1 As String, cHead2 As String, cHead3 As String
cHead1 = "„ «»⁄… ‘Â—Ì… ··›Ì“«"
cHead2 = retHeader(aHeader, 0, 1)
cHead3 = retHeader(aHeader, 1, 1)
PrintGrd.doprint GRID11, 1, -2, cHead1, cHead2, cHead3, False, False, 8, , Array(1), Array(1)
PrintGrd.Show 1
End Sub
Private Sub cmdExit_Click()
Unload Me
Set grdVisafrm2 = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
myload
myLoad11
End Sub

Private Sub CmdPrint2_Click()
doprint
End Sub
Private Sub Form_Load()
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "Select Code,DescA From FILE0_50 ORDER BY DESCA"
Set XBOX.RowSource = data1
XBOX.ListField = "Desca"
XBOX.BoundColumn = "Code"

Set grid1.DataSource = data10
data10.ConnectionString = strCon

Set GRID11.DataSource = DATA20
DATA20.ConnectionString = strCon

Fixgrd
Fixgrd11
grid1.Rows = 1
GRID11.Rows = 1
'grid11.Rows = 1
End Sub
Private Sub myload()
Dim cString As String
For i = 0 To UBound(aHeader)
    aHeader(i) = ""
Next
With grid1
    cString = "SELECT CONVERT(VARCHAR(10),INV_TOTAL.[DATE],111),SUM(TOTAL - DISCOUNT),SUM(VISA)  " & _
              " FROM INV_TOTAL"
    cString = cString & turn(cString) & "INV_TOTAL.PRINTED = 1"
    
    If IsDate(xDate1.Text) Then
        cString = cString & turn(cString) & " INV_TOTAL.DATE  >= " & DateSq(xDate1.Text)
        aHeader(0) = aHeader(0) & turn(cHead2, Space(3)) & "„‰ : " & Format(xDate1.Text, "yyyy/mm/dd")
    End If
       
    If IsDate(xDate2.Text) Then
        cString = cString & turn(cString) & " INV_TOTAL.DATE  <= " & DateSq(xDate2.Text)
        aHeader(0) = aHeader(0) & turn(cHead2, Space(3)) & "Õ Ì : " & Format(xDate2.Text, "yyyy/mm/dd")
    End If
       
    If XBOX.BoundText <> "" Then
         cString = cString & turn(cString) & " INV_TOTAL.[BOX]  = " & MyParn(XBOX.BoundText)
         aHeader(1) = "«·ﬂ«‘Ì— :" & XBOX.Text
    End If
    cString = cString & " GROUP BY  INV_TOTAL.DATE"
    cString = cString & turn(cString, " HAVING ", " AND ") & "SUM(INV_TOTAL.VISA) <> 0"
    cString = cString & " ORDER BY INV_TOTAL.DATE"
    data10.RecordSource = cString
    data10.Refresh
End With
Handlecontrols
Fixgrd
End Sub
Private Sub myLoad11()
Dim cString As String

For i = 0 To UBound(aHeader)
    aHeader(i) = ""
Next
With GRID11
    cString = "SELECT CONVERT(NVARCHAR(4),YEAR(DATE)) + '-' + CONVERT(VARCHAR(2),MONTH(date)) ,SUM(TOTAL - DISCOUNT),SUM(VISA)  " & _
              " FROM INV_TOTAL"
    cString = cString & turn(cString) & "INV_TOTAL.PRINTED = 1"
    
    If IsDate(xDate1.Text) Then
        cString = cString & turn(cString) & " INV_TOTAL.DATE  >= " & DateSq(xDate1.Text)
        aHeader(0) = aHeader(0) & turn(cHead2, Space(3)) & "„‰ : " & Format(xDate1.Text, "yyyy/mm/dd")
    End If
       
    If IsDate(xDate2.Text) Then
        cString = cString & turn(cString) & " INV_TOTAL.DATE  <= " & DateSq(xDate2.Text)
        aHeader(0) = aHeader(0) & turn(cHead2, Space(3)) & "Õ Ì : " & Format(xDate2.Text, "yyyy/mm/dd")
    End If
       
    If XBOX.BoundText <> "" Then
         cString = cString & turn(cString) & " INV_TOTAL.[BOX]  = " & MyParn(XBOX.BoundText)
         aHeader(1) = "«·ﬂ«‘Ì— :" & XBOX.Text
    End If
    cString = cString & " GROUP BY YEAR(DATE),MONTH(INV_TOTAL.DATE)"
    If Check1.Value = 0 Then cString = cString & turn(cString, " HAVING ", " AND ") & "SUM(INV_TOTAL.VISA) <> 0"
    cString = cString & " ORDER BY YEAR(DATE),MONTH(DATE)"
    DATA20.RecordSource = cString
    DATA20.Refresh
End With
Handlecontrols
Fixgrd11
End Sub

Sub Fixgrd()
Dim nTotal As Double, nSales As Double
  With grid1
 .RowHeight(0) = 1000
 .WordWrap = True
 .Cols = 4
 .TextMatrix(0, 0) = "«·ÌÊ„"
 .TextMatrix(0, 1) = "≈Ã„«·Ì «·ÌÊ„"
 .TextMatrix(0, 2) = "≈Ã„«·Ì «·›Ì“«"
 .TextMatrix(0, 3) = "‰”»… «·›Ì“«"
 .ColFormat(3) = "#0.00%"
 
 .ColWidth(0) = 2000
 .ColWidth(1) = 1200
 .ColWidth(2) = 1200
 .ColWidth(3) = 1200

 .ColDataType(1) = flexDTDouble
 .ColDataType(2) = flexDTDouble
 .ColDataType(3) = flexDTDouble
 
 
 .ExplorerBar = flexExSort
 .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
 
 .SubtotalPosition = flexSTAbove
 .Subtotal flexSTSum, -1, 1, "#0.00", vbRed, vbYellow, True, "  "
 .Subtotal flexSTSum, -1, 2, "#0.00", vbRed, vbYellow, True, "  "
 For i = 1 To grid1.Rows - 1
     If i > 1 Then grid1.TextMatrix(i, 0) = arbDay(grid1.TextMatrix(i, 1)) & " " & grid1.TextMatrix(i, 0)
     If Val(grid1.TextMatrix(i, 1)) <> 0 Then grid1.TextMatrix(i, 3) = Val(grid1.TextMatrix(i, 2)) / Val(grid1.TextMatrix(i, 1))
 Next
 
 If grid1.Rows > 1 Then
     .MergeCells = flexMergeFree
     For i = 0 To 0
         .TextMatrix(1, i) = "«·«Ã„«·Ì"
     Next
     .MergeRow(1) = True
 End If
 End With
End Sub
Sub Fixgrd11()
Dim nTotal As Double, nSales As Double
  With GRID11
 .RowHeight(0) = 1000
 .WordWrap = True
 .Cols = 4
 .TextMatrix(0, 0) = "«·‘Â—"
 .TextMatrix(0, 1) = "≈Ã„«·Ì «·‘Â—"
 .TextMatrix(0, 2) = "≈Ã„«·Ì «·›Ì“«"
 .TextMatrix(0, 3) = "‰”»… «·›Ì“«"
 .ColFormat(3) = "#0.00%"
 
 .ColWidth(0) = 2000
 .ColWidth(1) = 1200
 .ColWidth(2) = 1200
 .ColWidth(3) = 1200

 .ColDataType(1) = flexDTDouble
 .ColDataType(2) = flexDTDouble
 .ColDataType(3) = flexDTDouble
 
 
 .ExplorerBar = flexExSort
 .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
 
 .SubtotalPosition = flexSTAbove
 .Subtotal flexSTSum, -1, 1, "#0.00", vbRed, vbYellow, True, "  "
 .Subtotal flexSTSum, -1, 2, "#0.00", vbRed, vbYellow, True, "  "
 For i = 1 To .Rows - 1
     If Val(.TextMatrix(i, 1)) <> 0 Then .TextMatrix(i, 3) = Val(.TextMatrix(i, 2)) / Val(.TextMatrix(i, 1))
 Next
 
 If .Rows > 1 Then
     .MergeCells = flexMergeFree
     For i = 0 To 0
         .TextMatrix(1, i) = "«·«Ã„«·Ì"
     Next
     .MergeRow(1) = True
 End If
 End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set grditem1 = Nothing
End Sub
Private Sub xDesca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xdesca.Text, 1
End If
End Sub
Private Sub xITEM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, XITEM.Text, 0
End If
End Sub
Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub
Sub myProc()
ActiveControl.Text = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
Unload oSearchClient
End Sub
Private Sub xCode_LostFocus()
xCode.BackColor = &H80000005
xCodeDesca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text)
xCodeDesca.Caption = GetDesca("select desca from FILE3_10 where code = " & MyParn(xCode.Text)) & ""
End Sub
Private Sub xSection_Validate(Cancel As Boolean)
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
Private Sub xbox_Validate(Cancel As Boolean)
If Not XBOX.MatchedWithList Then XBOX.BoundText = ""
End Sub
Private Sub xgroup_Validate(Cancel As Boolean)
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub xGroup2_Validate(Cancel As Boolean)
If Not xGroup2.MatchedWithList Then xGroup2.BoundText = ""
End Sub
Private Sub xCode_GotFocus()
xCode.SelStart = 0
xCode.SelLength = Len(xCode.Text)
xCode.BackColor = &HC0FFFF
End Sub
Private Sub xDate1_GotFocus()
xDate1.SelStart = 0
xDate1.SelLength = Len(xDate1.Text)
xDate1.BackColor = &HC0FFFF
End Sub
Private Sub xdate2_GotFocus()
xDate2.SelStart = 0
xDate2.SelLength = Len(xDate2.Text)
xDate2.BackColor = &HC0FFFF
End Sub
Private Sub xGroup_GotFocus()
xGroup.BackColor = &HC0FFFF
End Sub
Private Sub xSection_GotFocus()
xSection.BackColor = &HC0FFFF
End Sub
Private Sub xbox_GotFocus()
XBOX.BackColor = &HC0FFFF
End Sub
Private Sub xGroup2_GotFocus()
xGroup2.BackColor = &HC0FFFF
End Sub
Private Sub xDate1_LostFocus()
xDate1.BackColor = &H80000005
End Sub
Private Sub xdate2_LostFocus()
xDate2.BackColor = &H80000005
End Sub
Private Sub xGroup_LostFocus()
xGroup.BackColor = &H80000005
End Sub
Private Sub xSection_LostFocus()
xSection.BackColor = &H80000005
End Sub
Private Sub xbox_LostFocus()
XBOX.BackColor = &H80000005
End Sub
Private Sub xgroup2_LostFocus()
xGroup2.BackColor = &H80000005
End Sub
Private Sub Handlecontrols()
cmdPrint.Enabled = grid1.Rows > 1
cmdprint11.Enabled = GRID11.Rows > 1
End Sub
Private Function doprint()
'On Error GoTo myerror
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
     
If IsDate(xDate1.Text) Then aHeader(0) = BetweenString(xDate1.Text, xDate2.Text)
If IsDate(xDate2.Text) Then aHeader(0) = BetweenString(xDate1.Text, xDate2.Text)
If XBOX.MatchedWithList Then aHeader(1) = turn(XBOX.Text, "«·ﬂ«‘Ì— : ") & XBOX.Text

With grid1
For i = 2 To grid1.Rows - 1
    
    temptable.AddNew
    temptable!str1 = TurnValue(ArbString(grid1.TextMatrix(i, 1)))
    temptable!str2 = TurnValue(grid1.TextMatrix(i, 2))
    temptable!val1 = Val(grid1.TextMatrix(i, 5))
    temptable!str12 = Format(grid1.TextMatrix(i, 1), "yyyy-mm-dd")
    temptable!Str11 = arbDay(grid1.TextMatrix(i, 1)) & turn(grid1.TextMatrix(i, 1), " ") & Format(grid1.TextMatrix(i, 1), "yyyy/mm/dd")
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
