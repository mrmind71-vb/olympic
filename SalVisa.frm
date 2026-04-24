VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form SalVisa 
   BackColor       =   &H00E0E0E0&
   Caption         =   "„»Ì⁄«  ›Ì“« Œ·«· › —…"
   ClientHeight    =   7140
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   7785
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   7785
   Begin Threed.SSCheck XPAY 
      Height          =   315
      Left            =   1425
      TabIndex        =   8
      Top             =   600
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   556
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   128
      PictureUseMask  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«·€Ì— „”œœ ›ﬁÿ"
      Alignment       =   1
   End
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00DEE7D3&
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   3225
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   225
      Width           =   915
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
      Height          =   315
      Left            =   4335
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   1785
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
      Height          =   315
      Left            =   4335
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   160
      Width           =   1785
   End
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00DEE7D3&
      Caption         =   "ÿ»«⁄…"
      Height          =   390
      Left            =   2235
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   225
      Width           =   915
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00DEE7D3&
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   195
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   225
      Width           =   1035
   End
   Begin VSFlex7LCtl.VSFlexGrid VsItem 
      Height          =   5520
      Left            =   225
      TabIndex        =   7
      Top             =   1350
      Width           =   7215
      _cx             =   12726
      _cy             =   9737
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
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
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
      TabBehavior     =   0
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   975
      Left            =   75
      Top             =   75
      Width           =   7515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰  «—ÌŒ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6225
      TabIndex        =   6
      Top             =   195
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï  «—ÌŒ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6225
      TabIndex        =   5
      Top             =   645
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   5850
      Left            =   75
      Top             =   1185
      Width           =   7515
   End
End
Attribute VB_Name = "SalVisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataTable As Recordset
Dim cString As String
Dim VisaPayTable As Recordset
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "„ «»⁄… „»Ì⁄«  ›Ì“« ·› —… "
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(xDate2.Text, "DD-MM-YYYY")
    
    Load PrintGrd
    PrintGrd.Doprint Me.VsItem, 1, -2, cHead1, cHead2, , False, True, 10
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdOk_Click()
cString = "SELECT FILE6_20.DATE, FILE6_22.DOC_NO, FILE6_22.VISA, FILE6_22.CASH FROM FILE6_20 RIGHT JOIN FILE6_22 ON FILE6_20.DOC_NO = FILE6_22.DOC_NO " & _
          " Where FILE6_22.VISA IS NOT NULL AND FILE6_22.VISA <> 0  "
If IsDate(xDate1.Text) Then cString = cString & " AND file6_20.DATE >= " & DateSql(xDate1.Text)
If IsDate(xDate2.Text) Then cString = cString & " AND file6_20.DATE <= " & DateSql(xDate2.Text)
cString = cString & " GROUP BY FILE6_20.DATE, FILE6_22.DOC_NO, FILE6_22.VISA, FILE6_22.CASH ORDER BY file6_20.DATE , FILE6_22.DOC_NO "
Set DataTable = mydb.OpenRecordset(cString)
Me.MousePointer = 11

If DataTable.RecordCount > 0 Then
    Fillgrd
Else
    VsItem.Rows = 1
End If
Me.MousePointer = 0
End Sub
Sub Fillgrd()
Dim nPack As Double
Dim nBal2 As Double
With VsItem
.FixedRows = 1

.ExplorerBar = flexExSortShow
.Rows = 1
DataTable.MoveFirst
.SubtotalPosition = flexSTBelow
Do While Not DataTable.EOF
     .AddItem ""
    .TextMatrix(.Rows - 1, 0) = Format(DataTable.Date, "DD-MM-YYYY")
    .TextMatrix(.Rows - 1, 1) = DataTable.DOC_NO
    .TextMatrix(.Rows - 1, 2) = Format(DataTable.VISA, "#0.00")
    VisaPayTable.Seek "=", DataTable.DOC_NO
    If Not VisaPayTable.NoMatch Then
        .TextMatrix(.Rows - 1, 3) = Format(VisaPayTable.Date, "DD-MM-YYYY")
        If XPAY.Value Then .RemoveItem .Rows - 1
    End If
    DataTable.MoveNext
Loop
.SubtotalPosition = flexSTAbove
.Subtotal flexSTSum, -1, 2, "#0.00", , vbRed, , " "
End With
End Sub
Private Sub Form_Load()
Set VisaPayTable = mydb.OpenRecordset("FILE8_00")
VisaPayTable.Index = "nDoc"
xDate1.Text = ""
xDate2.Text = ""
With VsItem
.Cols = 4
.Rows = 1

.TextMatrix(0, 0) = " «—ÌŒ"
.TextMatrix(0, 1) = "„” ‰œ"
.TextMatrix(0, 2) = "ﬁÌ„… «·›Ì“«"
.TextMatrix(0, 3) = " «—ÌŒ «·”œ«œ"

.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 1800
.MergeCells = flexMergeFree
.MergeCol(0) = True
.ColDataType(2) = flexDTDouble
.ColDataType(0) = flexDTDate
.ColDataType(3) = flexDTDate
End With
End Sub
Private Sub VsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If VsItem.TextMatrix(Row, 3) = "" Then
        VsItem.TextMatrix(Row, 3) = Format(Date, "dd-mm-yyyy")
    End If
End Sub

Private Sub VsItem_EnterCell()
    With VsItem
        If .Col = 3 Then
            .Editable = flexEDKbdMouse
            
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub
Private Sub VsItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With VsItem
    If .IsSubtotal(Row) Then Exit Sub
        If IsDate(.EditText) Then
            VisaPayTable.Seek "=", .TextMatrix(Row, 1)
            If VisaPayTable.NoMatch Then
                VisaPayTable.AddNew
            Else
                VisaPayTable.Edit
            End If
            VisaPayTable.DOC_NO = .TextMatrix(Row, 1)
            VisaPayTable.Date = DateValue(.EditText)
            VisaPayTable.Value = Val(.TextMatrix(Row, 2))
            VisaPayTable.Update
        Else
            VisaPayTable.Seek "=", .TextMatrix(Row, 1)
            If Not VisaPayTable.NoMatch Then VisaPayTable.Delete
        End If
    End With
End Sub
