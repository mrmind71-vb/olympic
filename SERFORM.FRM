VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form SerForm 
   BackColor       =   &H00E0E0E0&
   Caption         =   "»Õ’ ⁄‰ —ﬁ„ „”·”· ·’‰› "
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_Search 
      BackColor       =   &H00E0E0E0&
      Caption         =   "»ÕÀ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1365
   End
   Begin VB.TextBox xSerNo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3750
      MaxLength       =   15
      TabIndex        =   3
      Top             =   750
      Width           =   4215
   End
   Begin VB.TextBox xItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6750
      MaxLength       =   15
      TabIndex        =   0
      Top             =   225
      Width           =   1215
   End
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   2940
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   9000
      _cx             =   15875
      _cy             =   5186
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
      MousePointer    =   1
      BackColor       =   14737632
      ForeColor       =   -2147483640
      BackColorFixed  =   12615808
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   14737632
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
      FormatString    =   $"Serform.frx":0000
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
      TextStyleFixed  =   1
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·„”·”·"
      Height          =   195
      Left            =   8190
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ «·’‰›"
      Height          =   195
      Left            =   8190
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label xItemDesc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "SerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SerTable As Recordset
Dim itemTable As Recordset
Dim cString As String
Private Sub Cmd_Search_Click()
    With ItemInv
    .Rows = 1
    cString = " SELECT file6_201.doc_no, FILE6_20.DATE, FILE3_10.DESCA" & _
              " FROM (FILE6_20 RIGHT JOIN file6_201 ON (FILE6_20.ITEM = file6_201.item) AND (FILE6_20.DOC_NO =  " & _
              " file6_201.doc_no)) LEFT JOIN FILE3_10 ON FILE6_20.CODE = FILE3_10.CODE" & _
              " WHERE file6_201.serial = " & MyParn(xSerNo.Text) & " AND FILE6_201.ITEM = " & MyParn(xItem.Text)
    Set SerTable = mydb.OpenRecordset(cString, dbOpenDynaset)
    If SerTable.RecordCount > 0 Then
        Do While True
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Format(SerTable!Date, "DD-MM-YYYY")
            .TextMatrix(.Rows - 1, 1) = "„»Ì⁄« "
            .TextMatrix(.Rows - 1, 2) = SerTable.DOC_NO
            .TextMatrix(.Rows - 1, 3) = SerTable.DESCA
            .TextMatrix(.Rows - 1, 4) = 2
            SerTable.MoveNext
            If SerTable.EOF Then Exit Do
        Loop
    End If

    cString = " SELECT file6_101.doc_no, FILE6_10.DATE, FILE3_10.DESCA" & _
              " FROM (FILE6_10 RIGHT JOIN file6_101 ON (FILE6_10.ITEM = file6_101.item) AND (FILE6_10.DOC_NO =  " & _
              " file6_101.doc_no)) LEFT JOIN FILE3_10 ON FILE6_10.CODE = FILE3_10.CODE" & _
              " WHERE file6_101.serial = " & MyParn(xSerNo.Text) & " AND FILE6_101.ITEM = " & MyParn(xItem.Text)
    Set SerTable = mydb.OpenRecordset(cString, dbOpenDynaset)
    If SerTable.RecordCount > 0 Then
        Do While True
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Format(SerTable!Date, "DD-MM-YYYY")
            .TextMatrix(.Rows - 1, 1) = "„— Ã⁄ „»Ì⁄« "
            .TextMatrix(.Rows - 1, 2) = SerTable.DOC_NO
            .TextMatrix(.Rows - 1, 3) = SerTable.DESCA
            .TextMatrix(.Rows - 1, 4) = 3
            
            SerTable.MoveNext
            If SerTable.EOF Then Exit Do
        Loop
    End If

    cString = " SELECT file6_111.doc_no, FILE6_11.DATE, FILE4_10.DESCA" & _
              " FROM (FILE6_11 RIGHT JOIN file6_111 ON (FILE6_11.ITEM = file6_111.item) AND (FILE6_11.DOC_NO =  " & _
              " file6_111.doc_no)) LEFT JOIN FILE4_10 ON FILE6_11.CODE = FILE4_10.CODE" & _
              " WHERE file6_111.serial = " & MyParn(xSerNo.Text) & " AND FILE6_111.ITEM = " & MyParn(xItem.Text)
    Set SerTable = mydb.OpenRecordset(cString, dbOpenDynaset)
    If SerTable.RecordCount > 0 Then
        Do While True
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Format(SerTable!Date, "DD-MM-YYYY")
            .TextMatrix(.Rows - 1, 1) = "„— Ã⁄ „‘ —Ì« "
            .TextMatrix(.Rows - 1, 2) = SerTable.DOC_NO
            .TextMatrix(.Rows - 1, 3) = SerTable.DESCA
            .TextMatrix(.Rows - 1, 4) = 4
            SerTable.MoveNext
            If SerTable.EOF Then Exit Do
        Loop
    End If

    cString = " SELECT file7_201.doc_no, FILE7_20.DATE, FILE4_10.DESCA" & _
              " FROM (FILE7_20 RIGHT JOIN file7_201 ON (FILE7_20.ITEM = file7_201.item) AND (FILE7_20.DOC_NO =  " & _
              " file7_201.doc_no)) LEFT JOIN FILE4_10 ON FILE7_20.CODE = FILE4_10.CODE" & _
              " WHERE file7_201.serial = " & MyParn(xSerNo.Text) & " AND FILE7_201.ITEM = " & MyParn(xItem.Text)
    Set SerTable = mydb.OpenRecordset(cString, dbOpenDynaset)
    
    If SerTable.RecordCount > 0 Then
        Do While True
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Format(SerTable!Date, "DD-MM-YYYY")
            .TextMatrix(.Rows - 1, 1) = "„‘ —Ì« "
            .TextMatrix(.Rows - 1, 2) = SerTable.DOC_NO
            .TextMatrix(.Rows - 1, 3) = SerTable.DESCA
            .TextMatrix(.Rows - 1, 4) = 1
            SerTable.MoveNext
            If SerTable.EOF Then Exit Do
        Loop
    End If
    
    
    If .Rows > 1 Then
        .ColSort(0) = flexSortGenericAscending
'       .ColSort(4) = flexSortGenericAscending
        .Select 1, 0, 1, 1
'        .Sort = flexSortStringAscending
    End If
    End With
End Sub
Private Sub Form_Load()
    Set itemTable = mydb.OpenRecordset("FILE1_10", dbOpenDynaset)
    With ItemInv
    .Cols = 5
    .Rows = 1
    .Editable = flexEDNone
    .FormatString = " «—ÌŒ|" & "‰Ê⁄|" & "—ﬁ„ «·›« Ê—…|" & "≈”„"
    .ColWidth(0) = 1200
    .ColWidth(1) = 1500
    .ColWidth(2) = 1000
    .ColWidth(3) = 3000
    .ColWidth(4) = 0
    
    .ColDataType(0) = flexDTDate
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTDouble
    
    .ColAlignment(0) = flexAlignCenterCenter
    .ColAlignment(1) = flexAlignCenterCenter
    .ColAlignment(2) = flexAlignCenterCenter
    .ColAlignment(3) = flexAlignCenterCenter
   End With
End Sub
Private Sub xItem_Change()
    Dim nRow As Double
    itemTable.FindFirst " ITEM = " & MyParn(xItem.Text)
    If Not itemTable.NoMatch Then
        xItemDesc.Caption = itemTable.DESCA
    End If
End Sub
Private Sub xItem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
    Dim Generalarray(4)
    Dim GrdArray(2)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Item as «·’‰›,DescA as [«”„ «·’‰›] From file1_10 "
    Generalarray(3) = " Where DescA Like('*cFilter*')"
    Generalarray(4) = "Order by Item"
    GrdArray(1) = 1000
    GrdArray(2) = 4500
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Sub myProc()
If ActiveControl.Name = xItem.Name Then
    xItem.Text = GrdText(Search.Grid1, 0)
End If
Unload Search
End Sub
