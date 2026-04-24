VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form grdgasfrm2 
   Caption         =   " ›’Ì·Ì «Ê«„— «·‘⁄·"
   ClientHeight    =   10110
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   16785
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
   ScaleHeight     =   10110
   ScaleWidth      =   16785
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   7830
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   630
      Width           =   4830
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "grdgas2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2415
         Picture         =   "grdgas2.frx":3059
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grdgas2.frx":5483
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1215
         Picture         =   "grdgas2.frx":78EF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
   End
   Begin MSComctlLib.StatusBar SBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   9735
      Width           =   16785
      _ExtentX        =   29607
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
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
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   12690
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   4020
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
         Left            =   1170
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
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xcar 
         Height          =   360
         Left            =   90
         TabIndex        =   2
         Top             =   900
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "«·”Ì«—…"
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
         Index           =   1
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   990
         Width           =   1005
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   630
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   2520
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1890
      Top             =   45
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
      Left            =   45
      Top             =   135
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
      Height          =   7395
      Left            =   90
      TabIndex        =   7
      Top             =   1395
      Width           =   16620
      _cx             =   29316
      _cy             =   13044
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.Label xRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   8865
      Width           =   1860
   End
   Begin VB.Label Label7 
      Caption         =   "„⁄œ· «·«” Â·«ﬂ :"
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
      Left            =   2025
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   8910
      Width           =   2580
   End
   Begin VB.Label xTotal_Kilo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   9225
      Width           =   1860
   End
   Begin VB.Label Label4 
      Caption         =   "≈Ã„«·Ì «·ﬂÌ·Ê „ — «·„ﬁÿÊ⁄ :"
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
      Left            =   14130
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   9225
      Width           =   2580
   End
   Begin VB.Label xTotal_Gas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   8865
      Width           =   1860
   End
   Begin VB.Label Label3 
      Caption         =   "≈Ã„«·Ì «·ÊﬁÊœ «·„‰’—› :"
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
      Left            =   14130
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   8865
      Width           =   2580
   End
End
Attribute VB_Name = "grdgasfrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer
Dim con As New ADODB.Connection
Dim cString As String

Private Sub cmdExel_Click()
ToFileExel grid1, Array(1)
End Sub

Private Sub CmdPrint_Click()
Dim aHeader(3)
Dim cHead1 As String, cHead2 As String, cHead3 As String
cHead1 = " ›’Ì·Ì «Ê«„— ‘€· Œ·«· › —…"
If IsDate(xDate1.Text) Then aHeader(0) = BetweenString(Format(xDate1.Text, "YYYY-MM-DD"), Format(xDate2.Text, "YYYY-MM-DD"))
If IsDate(xDate2.Text) Then aHeader(0) = BetweenString(Format(xDate2.Text, "YYYY-MM-DD"), Format(xDate2.Text, "YYYY-MM-DD"))
If xCar.BoundText <> "" Then aHeader(1) = "«·”Ì«—… : " & xCar.Text
cHead2 = retHeader(aHeader, 0, 1)
cHead3 = retHeader(aHeader, 1, 2)

Dim aFooter(1) As Variant
cString = "≈Ã„«·Ì «·ÊﬁÊœ «·„‰’—› : " & xTotal_Gas.Caption
cString = cString & Space(100) & "„⁄œ· «·«” Â·«ﬂ : " & xRate.Caption
aFooter(0) = AddFlag(aFooter(0), "fontsize", 11)
aFooter(0) = AddFlag(aFooter(0), "text", cString)

aFooter(1) = AddFlag(aFooter(1), "fontsize", 11)
aFooter(1) = AddFlag(aFooter(1), "text", "≈Ã„«·Ì «·ﬂÌ·Ê „ — «·„ﬁÿÊ⁄ : " & Val(xTotal_Kilo.Caption))

printGrdNew.doprint Me.grid1, 1, -4, cHead1, cHead2, cHead3, , False, True, 10, , , , atotal, aFooter
printGrdNew.Show 1

'PrintGrd.doprint Me.grid1, 0.75, -2, cHead1, cHead2, cHead3, False, False, 8, , Array(1)
'PrintGrd.Show 1
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
Private Sub Form_Load()
Me.Top = 1000
Me.Left = 1000
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "Select Code,DescA From cars"
Set xCar.RowSource = data1
xCar.ListField = "Desca"
xCar.BoundColumn = "Code"

Set grid1.DataSource = data10
data10.ConnectionString = strCon
Fixgrd
LoadText Me
End Sub
Private Sub myload()
Dim cString As String, nPrevious As Double
With grid1
If Not MYVALID Then Exit Sub
grid1.Rows = 1
cString = " SELECT  ORDERS.CODE,ORDERS.ORDER_NO, ORDERS.DATE, ORDERS.COUNTER_OUT, ORDERS.COUNTER_IN, ORDERS.DISTANCE, DRIVER.DESCA,BON, GAS_ORDERS.QUANT," & _
          " GAS_ORDERS.COUNTER, DRIVER_1.DESCA AS DRIVER2 FROM ORDERS INNER JOIN  DRIVER ON ORDERS.DRIVER = DRIVER.CODE LEFT OUTER JOIN" & _
          " GAS_ORDERS ON ORDERS.CAR = GAS_ORDERS.CAR AND ORDERS.DATE = GAS_ORDERS.DATE LEFT OUTER JOIN " & _
          " DRIVER AS DRIVER_1 ON GAS_ORDERS.DRIVER = DRIVER_1.CODE"
cString = cString & turn(cString) & "ORDERS.Car = " & MyParn(xCar.BoundText)

If IsDate(xDate1.Text) Then
    cString = cString & turn(cString) & "ORDERS_GAS.Date >= " & DateSq(xDate1.Text)
End If
          
If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & "ORDERS_GAS.Date <= " & DateSq(xDate2.Text)
End If


cString = cString & " Order by ORDERS.DATE,ORDERS.CODE"
data10.RecordSource = cString
data10.Refresh
End With
Fixgrd
Handlecontrols
End Sub
Sub Fixgrd()
Dim nTotal_Sales As Double, nTotal_in As Double
    With grid1
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = "„"
    .TextMatrix(0, 1) = "«·ﬂÊœ"
    .TextMatrix(0, 2) = "—ﬁ„ «„— «·‘€·"
    .TextMatrix(0, 3) = "«· «—ÌŒ"
    .TextMatrix(0, 4) = "ﬁ—«¡… «·⁄œ«œ „‰"
    .TextMatrix(0, 5) = "ﬁ—«¡… «·⁄œ«œ ≈·Ì"
    .TextMatrix(0, 6) = "«·„”«›… «·„ﬁÿÊ⁄…"
    .TextMatrix(0, 7) = "≈”„ «·”«∆ﬁ"
    .TextMatrix(0, 8) = "—ﬁ„ «·»Ê‰"
    .TextMatrix(0, 9) = "«·ﬂ„Ì…"
    .TextMatrix(0, 10) = "—ﬁ„ «·⁄œ«œ ⁄‰ «· „ÊÌ‰"
    .TextMatrix(0, 11) = "≈”„ «·”«∆ﬁ"
        
    .ColWidth(0) = 900
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1300
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 3000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColWidth(11) = 3000
    .ColHidden(1) = True
'    Dim nTotal_Gas As Long, nTotal_kilo As Long, nRate As Double
'    For i = 1 To grid1.Rows - 1
'        .TextMatrix(i, 0) = 1
'        nTotal_Gas = Val(grid1.TextMatrix(i, 9)) + nTotal_Gas
'        nTotal_kilo = Val(grid1.TextMatrix(i, 6)) + nTotal_kilo
'    Next
'
'    xTotal_Gas.Caption = Myvalue(nTotal_Gas)
'    xTotal_Kilo.Caption = Myvalue(nTotal_kilo)
'    If nTotal_kilo = 0 Then
'        xRate.Caption = ""
'    Else
'        xRate.Caption = Myvalue(Round(nTotal_Gas / nTotal_kilo, 2))
'    End If
'    .ColDataType(2) = flexDTDouble
'    .ColDataType(3) = flexDTDouble
'    .ColDataType(4) = flexDTDouble

    
    For i = 0 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
               
'    If .Rows > 1 Then
'        .SubtotalPosition = flexSTAbove
'        .Subtotal flexSTSum, -1, 2, "#0.00", vbRed, vbYellow, True, "  "
'        .Subtotal flexSTSum, -1, 3, "#0.00", vbRed, vbYellow, True, "  "
'        .TextMatrix(1, 4) = Round(Val(.TextMatrix(1, 2)) - Val(.TextMatrix(1, 3)), 2)
'        For i = 0 To 1
'            .TextMatrix(1, i) = "«·≈Ã„«·Ï"
'        Next
'        .MergeCells = flexMergeFree
'        .MergeRow(1) = True
'    End If
    
    SBar1.Panels(1).Text = IIf(grid1.Rows > 2, "⁄œœ «·”Ã·«  : " & grid1.Rows - 2, "")
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdbankfrm1 = Nothing
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
'ItemsLookupAll Me, osearchitem, myFlag
End Sub

Private Sub xDesca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xDesca.Text, 1
End If
End Sub

Private Sub grid1_dblClick()
With grid1
If .TextMatrix(.Row, .Cols - 1) = "2" Then
    bankinoutfrm.sDoc_no = .TextMatrix(.Row, 1)
    bankinoutfrm.Show
ElseIf .TextMatrix(.Row, .Cols - 1) = "3" Or .TextMatrix(.Row, .Cols - 1) = "4" Then
    checkfrm1.sCode = .TextMatrix(.Row, 1)
    checkfrm1.Show
ElseIf .TextMatrix(.Row, .Cols - 1) = "5" Or .TextMatrix(.Row, .Cols - 1) = "6" Then
    checkfrm2.sCode = .TextMatrix(.Row, 1)
    checkfrm2.Show
End If
End With
End Sub

Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub
Private Sub Handlecontrols()
cmdPrint.Enabled = grid1.Rows > 1
End Sub

Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
End Sub
Private Sub xbox_GotFocus()
myGotFocus xbox
End Sub
Private Sub xbox_LostFocus()
myLostFocus xbox
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
End Sub
Sub myProc()
xCode.BoundText = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
Unload oSearchItem
End Sub
Private Function MYVALID() As Boolean
If Not xCar.MatchedWithList Then
    MsgBox "«·”Ì«—… €Ì— „”Ã·"
    Exit Function
End If
MYVALID = True
End Function
