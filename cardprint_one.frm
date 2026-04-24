VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{0BFA85A1-F9B8-11CF-8939-444553540000}#1.0#0"; "barcode.ocx"
Begin VB.Form cardPrint_one 
   Caption         =   "ØÈÇÚÉ ßÇÑäíåÇÊ"
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
      Height          =   6540
      Left            =   90
      TabIndex        =   0
      Top             =   225
      Width           =   5640
      _cx             =   9948
      _cy             =   11536
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
      Zoom            =   117.538461538462
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
   Begin BARCODELib.Barcode Barcode1 
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5265
      _Version        =   65536
      _ExtentX        =   9287
      _ExtentY        =   1535
      _StockProps     =   25
      Text            =   "1"
      Type            =   14
      TypeName        =   "Code 128"
      Text            =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowText        =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   525
      Top             =   2025
      Width           =   2190
   End
End
Attribute VB_Name = "cardPrint_one"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Form
Dim nLeftMargin, nCardHeight, nPageWidth, nRightMargin
Dim I, i2, nrow, nUpMargin As Integer, nCardWidth As Integer
Dim nCol As Integer
Dim tCard As ADODB.Recordset
Dim nTextWidth, nTextHeight
Private Sub Form_Load()
'Vp.PaperSize = pprA4
'Vp.PaperSize = pprLegal
Vp.PaperSize = pprUser
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
Private Function myRight(nRight, nWidth)
myRight = nRight - nWidth
End Function
Sub PrintArray()
Dim bNewRow, bNewPage
Dim nRate As Double, nHeight As Double
'On Error GoTo PrintError
'On Error Resume Next
'Set tCard = tempdb.OpenRecordset("Select * From Card order by CardNo")
Set tCard = New ADODB.Recordset
tCard.Open "Select * From Card order by CardNo", contemp, adOpenStatic, adLockReadOnly, adCmdText

nUpMargin = SettingArray(cUpMargin)
nRightMargin = SettingArray(cRightMargin)
nCardWidth = SettingArray(cCardWidth)
nCardHeight = SettingArray(cCardHeight)
nPrintRow = SettingArray(cRows)

nrow = 0
nCol = 0
nCols = SettingArray(cCols)
nPageWidth = SettingArray(cPageWidth)

'SetOriginalSettings
Vp.ZoomMode = zmWholePage
Vp.Orientation = orLandscape
Vp.StartDoc
Dim nRecordcount As Long, I As Long
If Not (tCard.EOF And tCard.BOF) Then
    tCard.MoveLast
    nRecordcount = tCard.RecordCount
    tCard.MoveFirst
    myForm.prog1.Visible = True
    myForm.prog1.Value = 0
End If
With Vp
    Do Until tCard.EOF
        I = I + 1
        myForm.prog1.Value = Round(I / (nRecordcount), 2) * 100
       .PenStyle = psTransparent
       .BrushStyle = bsTransparent
       If nCardNo <> tCard!CardNo Then
            nCol = IIf(nCol = nCols, 1, nCol + 1)
            nrow = IIf(nCol = 1, nrow + 1, nrow)
            nCardNo = tCard!CardNo
        End If
        
       If nrow > SettingArray(cRows) Then
           .NewPage
           nrow = 1
       End If
                
        If Not tCard!isPhoto And Not tCard!ISBARCODE Then
            If Val(tCard!BrushColor & "") = 0 Then .BrushStyle = bsTransparent Else .BrushStyle = bsSolid
            If Not IsNull(tCard!Text) Then
                 .FontName = tCard!FontName
                 .FontBold = tCard!FontBold
                 .FontSize = tCard!FontSize
                 .FontUnderline = tCard!FontUnderline
                 .FontItalic = tCard!FontItalic
                 If Not IsNull(tCard!PenColor) Then
                    .PenColor = vbBlack
                 End If
                 
                 If Not IsNull(tCard!PenWidth) Then
                    .PenWidth = tCard!PenWidth
                 End If
                 
                 .TextColor = TurnValue(tCard!ForeColor, Null, vbBlack)
                 nFieldWidth = IIf(tCard!Width = 0, .TextWidth(tCard!Text), tCard!Width)
                 nFieldHeight = IIf(tCard!Height = 0, .TextHeight(tCard!Text), tCard!Width)
                 Vp.TextAlign = IIf(IsNull(tCard!TextAlign), taRightTop, tCard!TextAlign)
                 If Not IsNull(tCard!TextAngle) Then .TextAngle = tCard!TextAngle
                 If (Not IsNull(tCard!BrushColor)) And tCard!BrushColor <> 0 Then
                     'Vp.BorderStyle = bsSingle
                     Vp.PenStyle = psTransparent
                     'Vp.BrushColor = tCard!BrushColor
                    .TextBox tCard!Text, Calcx, CalcY, nFieldWidth, nFieldHeight
                Else
                    .TextBox tCard!Text, Calcx, CalcY, nFieldWidth, nFieldHeight
                End If

            End If
         End If
         If tCard!ISBARCODE Then
             nFieldWidth = tCard!Width
             nFieldHeight = tCard!Height
             Barcode1.Text = tCard!Text
             DoEvents
             Barcode1.CreatePictureBySize tCard!Width, tCard!Height
            .DrawPicture Barcode1.Picture, Calcx, CalcY, tCard!Width, tCard!Height
        End If
        If tCard!isPhoto And Not IsNull(tCard!Text) Then
            .BrushStyle = bsTransparent
            If Val(tCard!PenWidth & "") = 0 Then
                .PenStyle = psTransparent
            Else
                .PenStyle = psSolid
                .PenWidth = Val(loctable!PenWidth & "")
            End If
'            nFieldWidth = tCard!Width
'            nFieldHeight = tCard!Height
        

            aRet = retDim(tCard!Width, tCard!Height)
            nFieldWidth = retFlag(aRet, "width")
            nadd1 = Int((tCard!Width - retFlag(aRet, "width")) / 2)
            nadd2 = Int(tCard!Height - retFlag(aRet, "height"))
                            
            .DrawPicture Image1.Picture, Calcx + nadd1, CalcY + nadd2, retFlag(aRet, "width"), retFlag(aRet, "height"), 11
            
        '.DrawPicture LoadPicture(tCard!Text, , 1), Calcx, CalcY, tCard!Width, tCard!Height, 10
        ElseIf tCard!isBox Then
            .BrushStyle = bsTransparent
            .PenStyle = psSolid
            .PenWidth = tCard!PenWidth
            .PenColor = Val(tCard!BrushColor & "")
            .PenWidth = tCard!FontSize
             nFieldWidth = tCard!Width
             nFieldHeight = tCard!Height
            .DrawRectangle Calcx, CalcY, Calcx + tCard!Width, CalcY + tCard!Height, 300, 300
'        ElseIf tCard!isRect Then
'            .BrushStyle = bsTransparent
'            .DrawRectangle CalcLeft, CalcTop, CalcLeft + tCard!Width, CalcTop + tCard!Height
        End If
        tCard.MoveNext
    Loop
End With
myForm.prog1.Visible = False
Vp.EndDoc
Exit Sub
PrintError:
myForm.prog1.Visible = False
MsgBox "ÎØÃ ãÇ ÞÏ ÍÏË ÇËäÇÁ ÇáØÈÇÚÉ " & Err.Description
Vp.EndDoc
End Sub
Private Function CalcY()
CalcY = ((nrow - 1) * nCardHeight) + nUpMargin + tCard!Top
End Function
Private Function Calcx()
If tCard!Width = 0 Then nFieldWidth = Vp.TextWidth(tCard!Text) Else nFieldWidth = tCard!Width
'nFieldWidth = IIf(tCard!Width = 0, Vp.TextWidth(tCard!Text), tCard!Width)
Calcx = nRightMargin - nFieldWidth - tCard!Right - ((nCol - 1) * nCardWidth)
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tCard.Close
Err.Clear
End Sub
Private Function CalcLeft()
CalcLeft = tCard!Left + ((nCol - 1) * nCardWidth) + (nRightMargin)
End Function
Private Function CalcTop()
CalcTop = tCard!Top + ((nrow - 1) * nCardHeight) + nUpMargin
End Function
Private Function retDim(pWidth As Long, pHeight As Long) As Variant
Dim nRate As Double
Set Image1.Picture = LoadPicture(tCard!Text)
nRate = (Image1.Picture.Width / Image1.Picture.Height)
nWidth = Int(pHeight * nRate)
If nWidth > pWidth Then
    nRate = (Image1.Picture.Height / Image1.Picture.Width)
    nHeight = Int(pWidth * nRate)
    nWidth = pWidth
Else
    nHeight = pHeight
End If
retDim = AddFlag(Empty, "width", nWidth)
retDim = AddFlag(retDim, "height", nHeight)
End Function

