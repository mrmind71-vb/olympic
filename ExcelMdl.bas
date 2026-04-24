Attribute VB_Name = "ExcelMdl"
Public Sub ToFileExel2(MyGrid, Optional aIg As Variant, Optional nRowHead As Long = 0, Optional aRowMerge As Variant = Empty, Optional aCol As Variant = Empty, Optional nRate As Double = 0, Optional aWidth As Variant = Empty, Optional arowHeight As Variant = Empty, Optional aSetUp As Variant = Empty, Optional nSize As Integer = 12, Optional acolSplit As Variant = Empty, Optional myForm As Form, Optional pHeader As Variant = Empty)
    Dim irow As Long, i As Long, i2 As Long, nRows As Long, nCols As Long, nFixedCols As Long, nFixedRows As Long, n As Long
    Dim icol As Long
    Dim objExcl As Excel.Application
    Dim objWk As Excel.Workbook
    Dim objSht As Excel.Worksheet
    Dim iHead As Long
    Dim vHead As Variant
    Dim sText As String
    
    'On Error Resume Next
    
    Set objExcl = Excel.Application
    objExcl.Application.Visible = False
    Set objWk = objExcl.Workbooks.Add
    Set objSht = objWk.Sheets(1)
    objExcl.Application.DisplayAlerts = True
        
    objSht.PageSetup.TopMargin = 10
    objSht.PageSetup.LeftMargin = 10
    objSht.PageSetup.HeaderMargin = 20
    objSht.PageSetup.CenterHeader = "&B &14"
    
        
    For i = 0 To MyGrid.FixedRows - 1
        If Not MyGrid.RowHidden(i) Then
            nFixedRows = nFixedRows + 1
        End If
    Next
                        
    For i = 0 To MyGrid.FixedCols - 1
        If Not MyGrid.ColHidden(i) Then
            nFixedCols = nFixedCols + 1
        End If
    Next
            
    
    For icol = 0 To MyGrid.Cols - 1
        If Not MyGrid.ColHidden(icol) Then
            nCols = nCols + 1
            If nFixedRows > 0 Then objSht.Range(objSht.Cells(1, nCols), objSht.Cells(nFixedRows, nCols)).NumberFormat = "@"
            If MyGrid.rows > nFixedRows Then
                If Not (MyGrid.ColDataType(icol) = flexDTDouble) Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCols), objSht.Cells(MyGrid.rows, nCols)).NumberFormat = "@"
                Else
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCols), objSht.Cells(MyGrid.rows, nCols)).NumberFormat = ""
                End If
            End If
        End If
    Next icol
    
    If Not myForm Is Nothing Then
        myForm.prog1.Visible = True
        myForm.prog1.Value = 0
        sCaption = myForm.Caption
    End If
    
    For irow = 0 To MyGrid.rows - 1
        If (Not myForm Is Nothing) And MyGrid.rows > 1 Then
            myForm.prog1.Value = IIf((irow / (MyGrid.rows - 1)) * 100 > 100, 100, (irow / (MyGrid.rows - 1)) * 100)
            myForm.Caption = sCaption & irow & " „‰ " & MyGrid.rows - 1
        End If
        If Not MyGrid.RowHidden(irow) Then
            nRows = nRows + 1
            nCols = 0
            For icol = 0 To MyGrid.Cols - 1
                If Not MyGrid.ColHidden(icol) Then
                    nCols = nCols + 1
                    If MyGrid.ColDataType(icol) = flexDTDate Then
                        objSht.Cells(nRows, nCols) = myFormat_p(MyGrid.Cell(flexcpTextDisplay, irow, icol))
                    ElseIf MyGrid.ColDataType(icol) = flexDTBoolean And irow > MyGrid.FixedRows - 1 Then
                        objSht.Cells(nRows, nCols) = IIf(Val(MyGrid.Cell(flexcpTextDisplay, irow, icol)) = 0, "·«", "‰⁄„")
                    Else
                        objSht.Cells(nRows, nCols) = MyGrid.Cell(flexcpTextDisplay, irow, icol)
                    End If
                End If
            Next icol
        End If
    Next irow
                                
    Dim nRow2 As Long
    If Not IsEmpty(aCol) Then
        For nCol = 0 To UBound(aCol)
            nValue = 0
            For nRow2 = 1 To nRows
                If Trim(objSht.Cells(nRow2, aCol(nCol))) <> Trim(cValue & "") Then
                    If nValue <> 0 Then
                        objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
                    End If
                    cValue = Trim(objSht.Cells(nRow2, aCol(nCol)))
                    nValue = 0
                    nBegin = nRow2
                Else
                    nValue = nValue + 1
                End If
            Next
            If nValue <> 0 Then
                objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
            End If
        Next
    End If
  
    If Not IsEmpty(acolSplit) Then
        For i = 0 To UBound(acolSplit)
            objSht.Columns(retFlag(acolSplit(i), "col")).PageBreak = xlPageBreakManual
        Next
    End If
    If nFixedRows > 0 Then
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedRows, nCols)).HorizontalAlignment = xlCenter
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedRows, nCols)).Interior.ColorIndex = 40
    End If
    If nRows > 0 Then
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).VerticalAlignment = xlCenter
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Borders.ColorIndex = 0
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Font.Size = nSize
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Font.Bold = True
    End If
    
    nCol = 0
    For icol = 0 To MyGrid.Cols - 1
        If Not MyGrid.ColHidden(icol) Then
            nCol = nCol + 1
            If nRows > 0 Then
                If MyGrid.ColFormat(icol) = "(##,##.##" Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).NumberFormat = "_(#,###.00_);[Red](#,###.00);0.00"
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlRight
                ElseIf MyGrid.ColAlignment(icol) = flexAlignLeftCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlLeft
                ElseIf MyGrid.ColAlignment(icol) = flexAlignRightCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlRight
                ElseIf MyGrid.ColAlignment(icol) = flexAlignCenterCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlCenter
                Else
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlLeft
                End If
            End If
            If nRate > 0 Then objSht.Columns(nCol).ColumnWidth = (MyGrid.ColWidth(icol) / 100) * nRate
        End If
    Next icol
               
    If Not IsEmpty(aWidth) Then
        For i = 0 To UBound(aWidth)
            If Val(retFlag(aWidth(i), "width")) = 0 Then
                objSht.Range(objSht.Cells(1, retFlag(aWidth(i), "col")), objSht.Cells(nRows, retFlag(aWidth(i), "col"))).Columns.AutoFit
            Else
                objSht.Columns(retFlag(aWidth(i), "col")).ColumnWidth = Val(retFlag(aWidth(i), "width")) / 100
            End If
        Next
    End If
    
    If Not IsEmpty(aRowMerge) Then
        For i = 0 To UBound(aRowMerge)
            If Not IsEmpty(retFlag(aRowMerge(i), "cols")) Then
                If Not IsEmpty(retFlag(aRowMerge(i), "text")) Then
                    objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))) = retFlag(aRowMerge(i), "text")
                End If
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))).Merge
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Interior.ColorIndex = 19
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Borders.ColorIndex = 0
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Bold = True
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Size = nSize
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).HorizontalAlignment = xlCenter
            End If
            
            If retFlag(aRowMerge(i), "split") Then
                objSht.rows(retFlag(aRowMerge(i), "row") + 1).PageBreak = xlPageBreakManual
            End If
            
            If retFlag(aRowMerge(i), "word_wrap") Then
                'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).WrapText = True
                objSht.rows(retFlag(aRowMerge(i), "row") + 1).WrapText = True
            End If
            If Not IsEmpty(retFlag(aRowMerge(i), "height")) Then
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).RowHeight = retFlag(arowHeight(i), "height")
               objSht.rows(retFlag(aRowMerge(i), "row") + 1).RowHeight = retFlag(aRowMerge(i), "height")
            End If
             If Not IsEmpty(retFlag(aRowMerge(i), "back_color")) Then
               objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Interior.ColorIndex = retFlag(aRowMerge(i), "back_color")
            End If
            
            If retFlag(aRowMerge(i), "bold") Then
               objSht.rows(retFlag(aRowMerge(i), "row") + 1).Font.Bold = True
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Bold = True
            End If
        Next
    End If
    
    objSht.PageSetup.Orientation = xlLandscape
    If Not IsEmpty(aSetUp) Then
        If Not IsEmpty(retFlag(aSetUp, "title_col")) Then
            objSht.PageSetup.PrintTitleColumns = objSht.Columns(retFlag(aSetUp, "title_col")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "title_row")) Then
            objSht.PageSetup.PrintTitleRows = objSht.rows(retFlag(aSetUp, "title_row")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "freeze")) Then
            ActiveWindow.SplitColumn = retFlag(aSetUp, "freeze")
            ActiveWindow.SplitRow = 0
            ActiveWindow.FreezePanes = True
        End If
        
        If retFlag(aSetUp, "autoSize") Then
            objSht.PageSetup.FitToPagesWide = True
        End If
                
        If retFlag(aSetUp, "center_header") <> "" Then
            objSht.PageSetup.CenterHeader = "&14&B" & retFlag(aSetUp, "center_header")
        End If
    End If
    
    
    If Not IsMissing(aIg) Then
        If IsEmpty(aIg) Then
            For i = 0 To UBound(aIg)
                objSht.rows(aIg(i)).Hidden = True
            Next
        End If
    End If
    
    
    If Not IsEmpty(pHeader) Then
        For i = 0 To UBound(pHeader)
            If Trim(pHeader(i)) <> "" Then
                objSht.Range("A1", "S1").Insert
                objSht.Range("A1", Chr(64 + nCols) & "1").Merge
                objSht.Range("A1", Chr(64 + nCols) & "1").Font.Size = nSize + 1
                objSht.Range("A1", Chr(64 + nCols) & "1").Font.Bold = True
                objSht.Range("A1", Chr(64 + nCols) & "1").VerticalAlignment = xlCenter
                objSht.Range("A1", Chr(64 + nCols) & "1").HorizontalAlignment = xlCenter
            End If
        Next
        
        For i = 0 To UBound(pHeader)
            If Trim(pHeader(i)) <> "" Then
                n = n + 1
                objSht.Cells(n, 1) = pHeader(i)
            End If
        Next
    
    End If
    If Not myForm Is Nothing Then
        myForm.prog1.Visible = True
        myForm.prog1.Value = 0
        myForm.Caption = sCaption
    End If
    
    objExcl.Application.Visible = True
''''''''''''''''''''''''''''
Set objSht = Nothing
Set objWk = Nothing
Set objExcl = Nothing
If Not myForm Is Nothing Then myForm.prog1.Visible = False
End Sub
Public Sub ToFileExel(MyGrid, Optional aIg As Variant, Optional nRowHead As Long = 0, Optional aRowMerge As Variant = Empty, Optional aCol As Variant = Empty, Optional nRate As Double = 1, Optional aWidth As Variant = Empty, Optional arowHeight As Variant = Empty, Optional aSetUp As Variant = Empty, Optional nSize As Integer = 12, Optional acolSplit As Variant = Empty, Optional myForm As Form, Optional nTopMargin As Integer = 30)
    Dim irow As Integer, i As Long, i2 As Long, nCols As Long
    Dim icol As Integer
    Dim objExcl As Excel.Application
    Dim objWk As Excel.Workbook
    Dim objSht As Excel.Worksheet
    Dim iHead As Integer
    Dim vHead As Variant
    On Error Resume Next
    Set objExcl = Excel.Application
    objExcl.Application.Visible = False
    Set objWk = objExcl.Workbooks.Add
    Set objSht = objWk.Sheets(1)
    objExcl.Application.DisplayAlerts = False
    Dim nRows As Long, nFixed As Long
        
    objSht.PageSetup.TopMargin = nTopMargin
    objSht.PageSetup.LeftMargin = 10
    objSht.PageSetup.HeaderMargin = nTopMargin
    objSht.PageSetup.CenterHeader = "&B &" & nSize
    
    objSht.Cells.NumberFormat = "@"
    
    If Not myForm Is Nothing Then
        myForm.prog1.Visible = True
        myForm.prog1.Value = 0
    End If
    
    For irow = 0 To MyGrid.rows - 1
        If Not MyGrid.RowHidden(irow) Then
            nRows = nRows + 1
            nCols = 0
            If Not myForm Is Nothing Then
                If myForm.prog1.Value <> Int((irow / MyGrid.rows - 1) * 100) Then
                    myForm.prog1.Value = IIf(Round(irow / (MyGrid.rows), 2) > 1, 1, Round(irow / (MyGrid.rows), 2)) * 100
                End If
            End If
            For icol = 0 To MyGrid.Cols - 1
                If Not MyGrid.ColHidden(icol) Then
                    nCols = nCols + 1
                    If MyGrid.ColDataType(icol) = flexDTDate And irow > MyGrid.FixedRows - 1 Then
                        objSht.Cells(nRows, nCols) = myFormat_p(MyGrid.Cell(flexcpTextDisplay, irow, icol))
                    ElseIf MyGrid.ColDataType(icol) = flexDTBoolean And irow > MyGrid.FixedRows - 1 Then
                        objSht.Cells(nRows, nCols) = IIf(Val(MyGrid.TextMatrix(irow, icol)) = 0, "·«", "‰⁄„")
                    ElseIf MyGrid.ColDataType(icol) = flexDTDouble And irow > MyGrid.FixedRows - 1 Then
                        objSht.Cells.NumberFormat = ""
                    Else
                        objSht.Cells(nRows, nCols) = MyGrid.Cell(flexcpTextDisplay, irow, icol) & ""
                    End If
                End If
            Next icol
        End If
    Next irow

    If Not myForm Is Nothing Then
        myForm.prog1.Visible = True
        myForm.prog1.Value = 0
    End If

    nFixed = 0
    For i = 0 To MyGrid.FixedRows - 1
        If Not MyGrid.RowHidden(i) Then
            nFixed = nFixed + 1
        End If
    Next
                        
    nFixedCols = 0
    For i = 0 To MyGrid.FixedCols - 1
        If Not MyGrid.ColHidden(i) Then
            nFixedCols = nFixedCols + 1
        End If
    Next
                    
            
    Dim nRow2 As Long
    If Not IsEmpty(aCol) Then
        For nCol = 0 To UBound(aCol)
            nValue = 0
            For nRow2 = 1 To nRows
                If Trim(objSht.Cells(nRow2, aCol(nCol))) <> Trim(cValue & "") Then
                    If nValue <> 0 Then
                        objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
                    End If
                    cValue = Trim(objSht.Cells(nRow2, aCol(nCol)))
                    nValue = 0
                    nBegin = nRow2
                Else
                    nValue = nValue + 1
                End If
            Next
            If nValue <> 0 Then
                objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
            End If
        Next
    End If
 
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixed, nCol)).Interior.ColorIndex = 40
'        objSht.Range(objSht.Cells(nFixed, 1), objSht.Cells(nFixed, nCol)).Font.bold = True
    
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Borders.ColorIndex = 0
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).AutoFit = True
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).VerticalAlignment = xlCenter
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedrows + 1, nCols)).HorizontalAlignment = xlCenter
'        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedrows + 1, nCols)).Interior.ColorIndex = 40
 
 
    If Not IsEmpty(aRowMerge) Then
        For i = 0 To UBound(aRowMerge)
            If Not IsEmpty(retFlag(aRowMerge(i), "cols")) Then
                If Not IsEmpty(retFlag(aRowMerge(i), "text")) Then
                    objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))) = retFlag(aRowMerge(i), "text")
                End If
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))).Merge
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Interior.ColorIndex = 19
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Borders.ColorIndex = 0
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Bold = True
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Size = nSize
            End If
            
            If retFlag(aRowMerge(i), "split") Then
                objSht.rows(retFlag(aRowMerge(i), "row") + 1).PageBreak = xlPageBreakManual
            End If
            
            If retFlag(aRowMerge(i), "word_wrap") Then
                'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).WrapText = True
                objSht.rows(retFlag(aRowMerge(i), "row") + 1).WrapText = True
            End If
            If Not IsEmpty(retFlag(aRowMerge(i), "height")) Then
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).RowHeight = retFlag(arowHeight(i), "height")
               objSht.rows(retFlag(aRowMerge(i), "row") + 1).RowHeight = retFlag(aRowMerge(i), "height")
            End If
             If Not IsEmpty(retFlag(aRowMerge(i), "back_color")) Then
               objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Interior.ColorIndex = retFlag(aRowMerge(i), "back_color")
            End If
            
            If retFlag(aRowMerge(i), "bold") Then
               objSht.rows(retFlag(aRowMerge(i), "row") + 1).Font.Bold = True
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Bold = True
            End If
        Next
    End If

    If Not IsEmpty(acolSplit) Then
        For i = 0 To UBound(acolSplit)
            objSht.Columns(retFlag(acolSplit(i), "col")).PageBreak = xlPageBreakManual
        Next
    End If


'    If Not IsEmpty(arowHeight) Then
'        For i = 0 To UBound(arowHeight)
'            objSht.Range(objSht.Cells(retFlag(arowHeight(i), "row") + 1, 1), objSht.Cells(retFlag(arowHeight(i), "row") + 1, NCOLS)).RowHeight = retFlag(arowHeight(i), "height")
'            If retFlag(arowHeight(i), "word_wrap") Then objSht.Range(objSht.Cells(retFlag(arowHeight(i), "row") + 1, 1), objSht.Cells(retFlag(arowHeight(i), "row") + 1, NCOLS)).WrapText = True
'            If retFlag(arowHeight(i), "bold") Then objSht.Range(objSht.Cells(retFlag(arowHeight(i), "row") + 1, 1), objSht.Cells(retFlag(arowHeight(i), "row") + 1, NCOLS)).WrapText = True
'        Next
'    End If

    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).VerticalAlignment = xlCenter
    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixed, nCols)).HorizontalAlignment = xlCenter
    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixed, nCols)).Interior.ColorIndex = 40
    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Borders.ColorIndex = 0
    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Font.Size = nSize
    objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Font.Bold = True
    
    nCol = 0
    For icol = 0 To MyGrid.Cols - 1
        If Not MyGrid.ColHidden(icol) Then
            nCol = nCol + 1
            If MyGrid.ColFormat(icol) = "(##,##.##" Then
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).NumberFormat = "_(#,###.00_);[Red](#,###.00);0.00"
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlRight
            ElseIf MyGrid.ColAlignment(icol) = flexAlignLeftCenter Then
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlLeft
            ElseIf MyGrid.ColAlignment(icol) = flexAlignRightCenter Then
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlRight
            ElseIf MyGrid.ColAlignment(icol) = flexAlignCenterCenter Then
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlCenter
            Else
                objSht.Range(objSht.Cells(nFixed + 1, nCol), objSht.Cells(nRows, nCols)).HorizontalAlignment = xlLeft
            End If
            If nRate > 0 Then objSht.Columns(nCol).ColumnWidth = (MyGrid.ColWidth(icol) / 100) * nRate
        End If
    Next icol
        
    If Not IsEmpty(aWidth) Then
        For i = 0 To UBound(aWidth)
            If Val(retFlag(aWidth(i), "width")) = 0 Then
                objSht.Range(objSht.Cells(1, retFlag(aWidth(i), "col")), objSht.Cells(nRows, retFlag(aWidth(i), "col"))).Columns.AutoFit
            Else
                objSht.Columns(retFlag(aWidth(i), "col")).ColumnWidth = Val(retFlag(aWidth(i), "width")) / 100
            End If
        Next
    End If
    
    objSht.PageSetup.Orientation = xlLandscape
    If Not IsEmpty(aSetUp) Then
        
        If Not IsEmpty(retFlag(aSetUp, "title_col")) Then
            objSht.PageSetup.PrintTitleColumns = objSht.Columns(retFlag(aSetUp, "title_col")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "title_row")) Then
            objSht.PageSetup.PrintTitleRows = objSht.rows(retFlag(aSetUp, "title_row")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "freeze")) Then
            ActiveWindow.SplitColumn = retFlag(aSetUp, "freeze")
            ActiveWindow.SplitRow = 0
            ActiveWindow.FreezePanes = True
        End If
        
        If retFlag(aSetUp, "autoSize") Then
            objSht.PageSetup.FitToPagesWide = True
        End If
                
        If retFlag(aSetUp, "center_header") <> "" Then
            objSht.PageSetup.CenterHeader = "&" & nSize & "&B" & retFlag(aSetUp, "center_header")
            objSht.PageSetup.CenterHeader = "&" & nSize & "&B" & retFlag(aSetUp, "center_header")
            objSht.PageSetup.CenterHeader = "&" & nSize & "&B" & retFlag(aSetUp, "center_header")
        End If
    End If
    
    'objSht.Range(objSht.Cells(0, 1), objSht.Cells(0, NCOLS)).WrapText = True
    
    'objSht.Range(objSht.Cells(1, 1), objSht.Cells(nRows, nCols)).Columns.AutoFit
    If Not IsMissing(aIg) Then
        For i = 0 To UBound(aIg)
            objSht.rows(aIg(i)).Hidden = True
        Next
    End If
    If Not myForm Is Nothing Then
        myForm.prog1.Visible = True
        myForm.prog1.Value = 0
    End If
    objExcl.Application.Visible = True
'    If Not IsEmpty(aRowMerge) Then
'        For i = 0 To UBound(aRowMerge)
'            objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols") + 1)).Merge
'        Next
'    End If

'''''''''''''''''''''''''''''
'›Ì Õ«·  —Ìœ Õ›Ÿ Ê—ﬁ… «·√ﬂ”·
'objWk.SaveAs "c:\Book1.xls"
'objWk.Close
'«·”ÿ— «· «·Ì Ì€·ﬁ »—‰«„Ã «·√ﬂ”·
'objExcl.Quit
''''''''''''''''''''''''''''
Set objSht = Nothing
Set objWk = Nothing
Set objExcl = Nothing
If Not myForm Is Nothing Then myForm.prog1.Visible = False
End Sub

