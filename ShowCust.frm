VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form ShowCust 
   BackColor       =   &H00E0E0E0&
   Caption         =   " ⁄œÌ· »Ì«‰«  «·⁄„·«¡"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E x i t"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   75
      Width           =   2115
   End
   Begin VSFlex7LCtl.VSFlexGrid VsItem 
      Height          =   6000
      Left            =   150
      TabIndex        =   0
      Top             =   600
      Width           =   11715
      _cx             =   20664
      _cy             =   10583
      _ConvInfo       =   1
      Appearance      =   1
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
      SelectionMode   =   0
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
      Editable        =   2
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
End
Attribute VB_Name = "ShowCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CardTable As Recordset
Dim TSalTable As Recordset
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Dim cStr1 As String
Set CardTable = mydb.OpenRecordset("SELECT * FROM FILE3_20")
CFIELD1 = myiif("STORE <>  'zz' ", "[TOTAL]") & " as T_SAL ,"
CFIELD2 = myiif("STORE =   'zz' ", "[TOTAL]") & " as T_DISC "

cStr1 = " SELECT FILE6_20.CUST , Max(FILE6_20.DATE) AS L_Date ," & _
        CFIELD1 & CFIELD2 & _
        " FROM FILE6_20 GROUP BY FILE6_20.CUST "
Set TSalTable = mydb.OpenRecordset(cStr1)

With VsItem
    .WordWrap = True
    .ExplorerBar = flexExSortShow
    .Rows = 1
    .Cols = 8
    .RowHeight(0) = 500
    .FixedCols = 1
    .FixedRows = 1
    .ExplorerBar = flexExSortShow
    .TextMatrix(0, 0) = "—ﬁ„  ·Ì›Ê‰"
    .TextMatrix(0, 1) = "≈”„"
    .TextMatrix(0, 2) = "⁄‰Ê«‰"
    .TextMatrix(0, 3) = "«·„œ—”…"
    .TextMatrix(0, 4) = "«·”‰… «·œ—«”Ì… "
    .TextMatrix(0, 5) = "≈Ã„«·Ï ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 6) = "≈Ã„«·Ï ﬁÌ„… Œ’„"
    .TextMatrix(0, 7) = "√Œ— „»Ì⁄«  "
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 2000
    .ColWidth(2) = 2000
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .FrozenCols = 2

    .Rows = 1
    If CardTable.RecordCount > 0 Then
        CardTable.MoveFirst
        Do While Not CardTable.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = CardTable.CODE
            .TextMatrix(.Rows - 1, 1) = CardTable.desca
            .TextMatrix(.Rows - 1, 2) = TurnValue(CardTable.ADDRESS, Null, "")
            .TextMatrix(.Rows - 1, 3) = TurnValue(CardTable.SCOOL, Null, "")
            .TextMatrix(.Rows - 1, 4) = TurnValue(CardTable.Class, Null, "")
            
            TSalTable.FindFirst " CUST = " & MyParn(CardTable.CODE)
            If Not TSalTable.NoMatch Then
                .TextMatrix(.Rows - 1, 5) = Format(TSalTable.T_SAL, "#0.00")
                .TextMatrix(.Rows - 1, 6) = Format(TSalTable.T_DISC, "#0.00")
                .TextMatrix(.Rows - 1, 7) = Format(TSalTable.L_DATE, "DD-MM-YYYY")
            End If
            CardTable.MoveNext
        Loop
    End If
End With
End Sub
Private Sub VsItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo myerror
    With VsItem
        CardTable.FindFirst " CODE = " & MyParn(.TextMatrix(Row, 0))
        If Not CardTable.NoMatch Then
            CardTable.Edit
            Select Case Col
                Case 1
                    If .EditText <> "" Then CardTable.desca = .EditText
                Case 2
                    If .EditText <> "" Then CardTable.ADDRESS = .EditText
                Case 3
                    If .EditText <> "" Then CardTable.SCOOL = .EditText
                Case 4
                    If .EditText <> "" Then CardTable.Class = .EditText
            End Select
            CardTable.Update
        End If
    End With
Exit Sub
myerror:
MsgBox Err.Description
MsgBox "·„ Ì „ «·Õ›Ÿ"
Err.Clear
End Sub
