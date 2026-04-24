VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form PrintGrd 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ÿ»«⁄…"
   ClientHeight    =   6720
   ClientLeft      =   735
   ClientTop       =   3000
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
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
      Top             =   90
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
      DefaultDevice   =   0   'False
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
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   3795
   End
End
Attribute VB_Name = "PrintGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cHeader As String, cBody As String, cFormat As String, cPageHeader1 As String, cPageHeader2 As String, cPageHeader3 As String, cFormat2 As String
Public sPrinter As String
Dim nFontSize As Integer
Dim aHRow, ahCol
Private Sub Form_Load()
If sPrinter <> "" Then Vp.Device = sPrinter
End Sub
Private Sub Form_Resize()
    Dim v!
    '----------------------------------------------------
    ' set height
    '----------------------------------------------------
    v = ScaleHeight - Vp.Top - 100
    If v > 0 Then Vp.Height = v
    
    '----------------------------------------------------
    ' set width
    '----------------------------------------------------
    v = ScaleWidth - Vp.Left
    If v > 0 Then Vp.Width = v
    
End Sub
Sub doprint(grid1, Optional nRate As Double = 1, Optional ptotal As Integer = -1, Optional pString1 As String = "", Optional pString2 As String = "", Optional pString3 As String = "", Optional bLeft As Boolean = False, Optional bLand As Boolean = False, Optional pFontSize As Integer = 11, Optional pFontName As String = "Simplified Arabic", Optional ByVal aRowSpan, Optional aColSpan As Variant)
nFontSize = pFontSize
cPageHeader1 = ArbString(pString1)
cPageHeader2 = ArbString(pString2)
cPageHeader3 = ArbString(pString3)

nBegin = IIf(bLeft, 0, grid1.Cols - 1)
nEnd = IIf(bLeft, grid1.Cols - 1, 0)
nStep = IIf(bLeft, 1, -1)
Vp.Orientation = IIf(bLand, orLandscape, orPortrait)
Vp = " "
With Vp
.ExportFormat = vpxRTF

'.ExportFile = "D:\Ali Mail" & "\RepDoc.RTF"

Vp.FontSize = pFontSize
Vp.FontName = pFontName
Vp.MarginLeft = 300
Vp.MarginRight = 300
Vp.TextAlign = taCenterTop
Vp.MarginTop = 300
Vp.MarginBottom = 300

Dim nSpan As Long, nSpan2 As Long

Dim I As Long

For nCol = nBegin To nEnd Step nStep
    If Not grid1.ColHidden(nCol) Then
        cFormat = cFormat & turn(cFormat, "|") & "+>" & (nRate * grid1.ColWidth(nCol))
    End If
Next

If Not IsMissing(aColSpan) Then
    For I = 0 To UBound(aColSpan)
        aColSpan(I) = Printcol(aColSpan(I), grid1)
    Next
End If

For nRow = 0 To grid1.rows - 1
    If Not grid1.RowHidden(nRow) Then
        For nCol = nBegin To nEnd Step nStep
            If Not grid1.ColHidden(nCol) Then
                If nRow < grid1.FixedRows Then
                    cHeader = cHeader & turn(cHeader, "|") & grid1.TextMatrix(nRow, nCol)
                Else
                    If InStr(1, grid1.ColFormat(nCol), "%") <> 0 Then
                        cBody = cBody & turn(grid1.TextMatrix(nRow, nCol), Round(Val(grid1.TextMatrix(nRow, nCol)) * 100, 2) & "%") & "|"
                    Else
                        If IsNumeric(grid1.TextMatrix(nRow, nCol)) Then
                            cBody = cBody & grid1.TextMatrix(nRow, nCol) & "|"
                        Else
                            cBody = cBody & ArbString(grid1.TextMatrix(nRow, nCol)) & "|"
                        End If
                    End If
                End If
            End If
        Next
        cBody = cBody & turn(cBody, ";")
    End If
Next

.StartDoc
.FontSize = 12
.FontBold = True
.FontUnderline = True

.FontSize = pFontSize
.FontBold = False
.FontUnderline = False

cHeader = cHeader & turn(cHeader, ";")
.StartTable

.AddTable cFormat, cHeader, cBody, , , True
.TableCell(tcFontBold, 0, 1, 0, Vp.TableCell(tcCols)) = taCenterMiddle
.TableCell(tcAlign, 0, 1, 0, Vp.TableCell(tcCols)) = taCenterMiddle
.TableCell(tcBackColor, 0, 1, 0, Vp.TableCell(tcCols)) = &H8000000F


'If Not IsMissing(aSpan) Then
'    For i = 0 To nRowSpan1
'        .TableCell(tcColSpan, aSpan(0) + i, nSpan + 1) = aSpan(2)
'        .TableCell(tcColAlign, aSpan(0) + i, nSpan, aSpan(0), nSpan + aSpan(2) - 1) = taCenterMiddle
'    Next
'End If


If ptotal = -1 Then
    .TableCell(tcFontBold, Vp.TableCell(tcRows), 0, Vp.TableCell(tcRows), Vp.TableCell(tcCols)) = True
    .TableCell(tcBackColor, Vp.TableCell(tcRows), 0, Vp.TableCell(tcRows), Vp.TableCell(tcCols)) = &HC0FFFF
ElseIf ptotal = -2 Then
    .TableCell(tcFontBold, 1, 1, 1, Vp.TableCell(tcCols)) = True
    .TableCell(tcBackColor, 1, 1, 1, Vp.TableCell(tcCols)) = &HC0FFFF
ElseIf ptotal = -3 And Not IsMissing(aRowSpan) Then
    For I = 0 To UBound(aRowSpan)
        .TableCell(tcFontBold, aRowSpan(I), 1, aRowSpan(I), Vp.TableCell(tcCols)) = True
        .TableCell(tcBackColor, aRowSpan(I), 1, aRowSpan(I), Vp.TableCell(tcCols)) = &HC0FFFF
    Next
End If

Dim nRowMerge As Long
If Not IsMissing(aRowSpan) Then
    For I = 0 To UBound(aRowSpan)
        nRowMerge = aRowSpan(I)
        MergeRow nRowMerge
    Next
End If

If Not IsMissing(aColSpan) Then MergeCols aColSpan

.EndTable
.EndDoc
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set PrintGrd = Nothing
End Sub
Private Sub vp_NewPage()
With Vp
.FontSize = 10
.FontBold = False
.TextAlign = taLeftTop
.TextBox "’›Õ… " & Vp.CurrentPage, 400, 20, 4000, 1000
.TextAlign = taRightTop
.TextBox "«· «—ÌŒ : " & Format(Date, "YYYY-MM-DD"), Vp.PageWidth - (4000 + 400), 20, 4000, 1000
.TextAlign = taCenterMiddle
If cPageHeader1 & cPageHeader2 & cPageHeader3 <> "" Then
    .FontSize = 14
    .FontBold = True
    .FontUnderline = True
    
    If cPageHeader1 <> "" Then .Paragraph = cPageHeader1
    If cPageHeader2 <> "" Then .Paragraph = cPageHeader2
    If cPageHeader3 <> "" Then .Paragraph = cPageHeader3
    
    .FontBold = False
    .FontUnderline = False
    .FontBold = False
End If
.FontSize = nFontSize

End With
End Sub
Private Sub MergeCols(aColSpan)
For I = 0 To UBound(aColSpan)
   MergeCol Abs(aColSpan(I))
Next
End Sub
Private Sub MergeRow(nRow As Long)
Dim nValue As Integer, cString As String, nCols As Long
cValue = "dummy"
For I = 1 To Vp.TableCell(tcCols)
    If Trim(Vp.TableCell(tcText, nRow, I)) <> Trim(cValue) Then
        If nValue > 1 Then
            If (Not IsNumeric(cValue)) And Trim(cValue) <> "" Then Vp.TableCell(tcColSpan, nRow, I - (nValue)) = nValue
        End If
        cValue = Vp.TableCell(tcText, nRow, I)
        nValue = 1
    Else
        nValue = nValue + 1
    End If
Next
If nValue > 1 Then
    If (Not IsNumeric(cValue)) Then Vp.TableCell(tcColSpan, nRow, I - (nValue)) = nValue
End If
End Sub
Private Sub MergeCol(nCol)
Dim nValue As Integer
cValue = "Dummy"
NROWS = Vp.TableCell(tcRows)
For I = 1 To NROWS
    If Trim(Vp.TableCell(tcText, I, nCol)) <> Trim(cValue) Then
        If nValue > 1 Then
            Vp.TableCell(tcRowSpan, I - (nValue), nCol) = nValue
        End If
        cValue = Vp.TableCell(tcText, I, nCol)
        nValue = 1
    Else
        nValue = nValue + 1
    End If
Next
If nValue > 1 Then
    Vp.TableCell(tcRowSpan, I - (nValue), nCol) = nValue
End If
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

