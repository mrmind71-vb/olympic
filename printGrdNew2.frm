VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form printGrdNew 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ÿ»«⁄…"
   ClientHeight    =   6720
   ClientLeft      =   735
   ClientTop       =   3000
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   6720
   ScaleWidth      =   5775
   WindowState     =   2  'Maximized
   Begin VSPrinter7LibCtl.VSPrinter Vp 
      Height          =   4515
      Left            =   150
      TabIndex        =   1
      Top             =   75
      Width           =   4665
      _cx             =   8229
      _cy             =   7964
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   -1  'True
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   0
      MarginTop       =   0
      MarginRight     =   0
      MarginBottom    =   0
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   23.3901515151515
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "Positioning"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   3795
   End
End
Attribute VB_Name = "PrintGrdNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Form
Dim cHeader As String, cBody As String, cFormat As String, cPageHeader1 As String, cPageHeader2 As String, cPageHeader3 As String, cPageHeader4 As String, cFormat2 As String
Dim nFontSize As Integer
Dim aHRow, ahCol
Dim aGrid As Variant
Private Sub Form_Resize()
    Dim v!
    
    '----------------------------------------------------
    ' set height
    '----------------------------------------------------
    v = ScaleHeight - vp.Top - 100
    If v > 0 Then vp.Height = v
    
    '----------------------------------------------------
    ' set width
    '----------------------------------------------------
    v = ScaleWidth - vp.Left
    If v > 0 Then vp.Width = v
    
End Sub
Sub doPrint(grid1, Optional nRate As Double = 1, Optional ptotal As Integer = -1, Optional pString1 As String = "", Optional pString2 As String = "", Optional pString3 As String = "", Optional pString4 As String = "", Optional bLeft As Boolean = False, Optional bLand As Boolean = False, Optional pFontSize As Integer = 11, Optional pFontName As String = "Arial", Optional ByVal aRowSpan, Optional aColSpan As Variant, Optional nRowHeight As Integer = -1, Optional nMarginLeft As Integer = 500, Optional nMarginRight As Integer = 400)
'Sub doprint(grid1, Optional nRate As Double = 1, Optional ptotal As Integer = -1, Optional pString1 As String = "", Optional pString2 As String = "", Optional pString3 As String = "", Optional bLeft As Boolean = False, Optional bLand As Boolean = False, Optional pFontSize As Integer = 11, Optional pFontName As String = "Simplified Arabic", Optional ByVal aSpan, Optional nRowSpan1 As Long = 0, Optional aColSpan As Variant)
nFontSize = pFontSize
'If Not IsMissing(phRow) Then aHRow = phRow
'If Not IsMissing(phCol) Then ahCol = phCol
cPageHeader1 = pString1
cPageHeader2 = pString2
cPageHeader3 = pString3
cPageHeader4 = pString4

aGrid = AddFlag(Empty, "left", bLeft)
nBegin = IIf(bLeft, 0, grid1.Cols - 1)
nEnd = IIf(bLeft, grid1.Cols - 1, 0)
nStep = IIf(bLeft, 1, -1)
vp.Orientation = IIf(bLand, orLandscape, orPortrait)
'VP.MarginLeft =
vp = " "
With vp
.ExportFormat = vpxRTF
'.ExportFile = "D:\Ali Mail" & "\RepDoc.RTF"

vp.FontSize = pFontSize
vp.FontBold = True
vp.FontName = pFontName
'vp.MarginLeft = 500
'vp.MarginRight = 500
vp.TextAlign = taCenterTop
vp.MarginLeft = nMarginLeft
vp.MarginRight = nMarginRight
vp.MarginTop = 750
vp.MarginBottom = 500

Dim nSpan As Long, nSpan2 As Long
Dim I As Long
For nCol = nBegin To nEnd Step nStep
    If Not grid1.ColHidden(nCol) Then
        cFormat = cFormat & turn(cFormat, "|") & IIf(bLeft, "<+", "+>") & (nRate * grid1.ColWidth(nCol))
    End If
Next

If Not IsMissing(aColSpan) And retFlag(aGrid, "Left") = False Then
    For I = 0 To UBound(aColSpan)
        aColSpan(I) = Printcol(aColSpan(I), grid1)
    Next
End If

If Not IsMissing(aRowSpan) Then
    Dim nColBgn As Long
    If (Not IsEmpty(aRowSpan)) And retFlag(aGrid, "Left") = False Then
        For I = 0 To UBound(aRowSpan)
            nColBgn = Val(Printcol(retFlag(aRowSpan(I), "col"), grid1)) - Val(retFlag(aRowSpan(I), "cols")) + 1
            aRowSpan(I) = AddFlag(aRowSpan(I), "col", nColBgn, True)
        Next
    End If
End If
If Not myForm Is Nothing Then
    myForm.prog1.Visible = True
    myForm.prog1.Value = 0
End If
For nRow = 0 To grid1.rows - 1
    If grid1.rows > 1 And Not myForm Is Nothing Then myForm.prog1.Value = Round(nRow / (grid1.rows - 1), 2) * 100
    If Not grid1.RowHidden(nRow) Then
        For nCol = nBegin To nEnd Step nStep
            If Not grid1.ColHidden(nCol) Then
                If nRow < 1 Then
                    cHeader = cHeader & turn(cHeader, "|") & grid1.TextMatrix(nRow, nCol)
                Else
                    If bLeft Then
                        cBody = cBody & grid1.Cell(flexcpTextDisplay, nRow, nCol) & "|"
                    Else
                        If grid1.ColDataType(nCol) = flexDTBoolean And nRow > grid1.FixedRows - 1 Then
                            cBody = cBody & IIf(Val(grid1.TextMatrix(nRow, nCol)) = 0, "", "‰⁄„") & "|"
                        ElseIf grid1.ColDataType(nCol) = flexDTDouble And nRow > grid1.FixedRows - 1 Then
                            cBody = cBody & grid1.Cell(flexcpTextDisplay, nRow, nCol) & "|"
                        Else
                            cBody = cBody & ArbString(grid1.Cell(flexcpTextDisplay, nRow, nCol)) & "|"
                        End If
                    End If
                End If
            End If
        Next
        cBody = cBody & turn(cBody, ";")
    End If
Next

If Not myForm Is Nothing Then
    myForm.prog1.Visible = False
    myForm.prog1.Value = 0
End If

.StartDoc
.FontSize = 11
.FontBold = True
.FontUnderline = True
'If cPageHeader1 <> "" Then .Paragraph = cPageHeader1
'If cPageHeader2 <> "" Then .Paragraph = cPageHeader2
'If cPageHeader3 <> "" Then .Paragraph = cPageHeader3
'If cPageHeader4 <> "" Then .Paragraph = cPageHeader4

.FontSize = pFontSize
.FontBold = True
.FontUnderline = False

cHeader = cHeader & turn(cHeader, ";")
.StartTable

.AddTable cFormat, cHeader, cBody, , , True
.TableCell(tcFontBold, 0, 1, 0, vp.TableCell(tcCols)) = taCenterMiddle
.TableCell(tcAlign, 0, 1, 0, vp.TableCell(tcCols)) = taCenterMiddle
.TableCell(tcBackColor, 0, 1, 0, vp.TableCell(tcCols)) = &H8000000F


'If Not IsMissing(aSpan) Then
'    For i = 0 To nRowSpan1
'        .TableCell(tcColSpan, aSpan(0) + i, nSpan + 1) = aSpan(2)
'        .TableCell(tcColAlign, aSpan(0) + i, nSpan, aSpan(0), nSpan + aSpan(2) - 1) = taCenterMiddle
'    Next
'End If


If ptotal = -1 Then
    .TableCell(tcFontBold, vp.TableCell(tcRows), 0, vp.TableCell(tcRows), vp.TableCell(tcCols)) = True
    .TableCell(tcBackColor, vp.TableCell(tcRows), 0, vp.TableCell(tcRows), vp.TableCell(tcCols)) = &HC0FFFF
ElseIf ptotal = -2 Then
    .TableCell(tcFontBold, 1, 1, 1, vp.TableCell(tcCols)) = True
    .TableCell(tcBackColor, 1, 1, 1, vp.TableCell(tcCols)) = &HC0FFFF
End If

If nRowHeight > -1 Then .TableCell(tcRowHeight, 0, 1, vp.TableCell(tcRows)) = nRowHeight

If (Not IsMissing(aRowSpan)) Then
    If Not IsEmpty(aRowSpan) Then MergeRows aRowSpan
End If
If (Not IsMissing(aColSpan)) Then MergeCols aColSpan
.EndTable
.EndDoc
For I = 1 To vp.PageCount
    vp.StartOverlay I
    vp.FontName = "Arial"
    vp.FontSize = 10
    vp.CurrentX = vp.MarginLeft + 300
    vp.CurrentY = vp.MarginTop - 300
    vp.TextAlign = taLeftTop
    vp.Paragraph = "’›Õ… " & I & " „‰ " & vp.PageCount
    vp.EndOverlay
Next
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set PrintGrdNew = Nothing
End Sub
Private Sub vp_NewPage()
With vp
.FontSize = 10
.FontBold = True
.TextAlign = taLeftTop
'If retFlag(aGrid, "left") Then
'    .TextBox "Page : " & vp.CurrentPage, 600, 600, 4000, 1000
'Else
'    .TextBox "’›Õ… : " & vp.CurrentPage, 600, 600, 4000, 1000
'End If
.TextAlign = taRightTop
.TextBox "«· «—ÌŒ : " & myFormat_p(Date), vp.PageWidth - 4500, vp.MarginTop - 300, 4000, 1000
.TextAlign = taCenterMiddle
.FontSize = 12
.FontBold = True
.FontUnderline = True
If cPageHeader1 <> "" Then .Paragraph = cPageHeader1
.FontUnderline = False
If cPageHeader2 <> "" Then .Paragraph = cPageHeader2
If cPageHeader3 <> "" Then .Paragraph = cPageHeader3
If cPageHeader4 <> "" Then .Paragraph = cPageHeader4
.FontSize = 3
.Paragraph = ""
.FontSize = nFontSize
.FontBold = True
.FontUnderline = False
.FontSize = nFontSize
End With
End Sub
Private Sub MergeRows(aRowSpan)
If IsEmpty(aRowSpan) Then Exit Sub
For I = 0 To UBound(aRowSpan)
   MergeRow aRowSpan(I)
Next
End Sub
Private Sub MergeCols(aColSpan)
For I = 0 To UBound(aColSpan)
   MergeCol Abs(aColSpan(I))
Next
End Sub
Private Sub MergeRow(aRow As Variant)
Dim nValue As Integer, cString As String
'cValue = "dummy"
'NCOLS = Vp.TableCell(tcCols)
'For i = nSpanBegin To NCOLS
'     If Trim(Vp.TableCell(tcText, nRow, i)) <> Trim(cValue) Then
'        If nValue > 1 Then
'            Vp.TableCell(tcColSpan, nRow, i - (nValue)) = nValue
'        End If
'        cValue = Vp.TableCell(tcText, nRow, i)
'        nValue = 1
'    Else
'        nValue = nValue + 1
'    End If
'Next
'If nValue > 1 Then
'    Vp.TableCell(tcColSpan, nRow, i - (nValue)) = nValue
'End If
If Trim(retFlag(aRow, "text")) <> "" Then
    vp.TableCell(tcText, retFlag(aRow, "row"), retFlag(aRow, "col"), retFlag(aRow, "row"), retFlag(aRow, "col")) = Trim(retFlag(aRow, "text"))
End If
vp.TableCell(tcColSpan, retFlag(aRow, "row"), retFlag(aRow, "col"), retFlag(aRow, "row"), retFlag(aRow, "col")) = retFlag(aRow, "cols")
vp.TableCell(tcColAlign, retFlag(aRow, "row"), retFlag(aRow, "col")) = taRightMiddle
vp.TableCell(tcFontBold, retFlag(aRow, "row"), 1, retFlag(aRow, "row"), vp.TableCell(tcCols)) = True
vp.TableCell(tcBackColor, retFlag(aRow, "row"), 1, retFlag(aRow, "row"), vp.TableCell(tcCols)) = &HC0FFFF
End Sub
Private Sub MergeCol(nCol)
Dim nValue As Integer
'cValue = "Dummy"
nRows = vp.TableCell(tcRows)
'For i = 1 To nRows
'    If Trim(Vp.TableCell(tcText, i, nCol)) <> Trim(cValue) Then
'        If nValue > 1 Then
'            Vp.TableCell(tcRowSpan, i - (nValue), nCol) = nValue
'        End If
'        cValue = Vp.TableCell(tcText, i, nCol)
'        nValue = 1
'    Else
'        nValue = nValue + 1
'    End If
'Next
Dim aCol As Long, nBegin As Long, cString As String
'cValue = "Dummy"
If nRows > 0 Then
    cValue = Trim(vp.TableCell(tcText, 1, nCol))
    nBegin = 1
End If
For I = 1 To nRows
    If Trim(vp.TableCell(tcText, I, nCol)) <> Trim(cValue) Then
        cValue = Trim(vp.TableCell(tcText, I, nCol))
        vp.TableCell(tcRowSpan, nBegin, nCol, nBegin, nCol) = nValue
        nValue = 1
        nBegin = I
    Else
        nValue = nValue + 1
    End If
Next
vp.TableCell(tcRowSpan, nBegin, nCol, nBegin, nCol) = nValue
End Sub
Private Function Printcol(nCol, pGrid) As Long
With pGrid
For I = .Cols - 1 To nCol Step -1
    If Not pGrid.ColHidden(I) Then
         Printcol = Printcol + 1
    End If
Next
End With
End Function

