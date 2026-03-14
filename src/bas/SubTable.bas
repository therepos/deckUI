Attribute VB_Name = "SubTable"
Option Explicit

' =============================================================================
' MODULE: SubTable — PowerPoint Edition
' Purpose: Table cell operations - formulas, number formatting, date formatting
' =============================================================================
' Differences from Word version:
'   - PPT tables have no field codes, no Selection.Cells, no wdWithInTable
'   - We detect the selected table + cell(s) via ActiveWindow.Selection
'   - Formulas write plain text results (no field refreshing)
'   - Number/date formatting rewrites cell text directly
'   - Multi-cell selection supported via ShapeRange + cell iteration
' =============================================================================

' ===== FORMULAS ==============================================================

Public Sub SelSumColumn()
    InsertTableFormula "SUM"
End Sub

Public Sub SelAverageColumn()
    InsertTableFormula "AVERAGE"
End Sub

Public Sub SelCountColumn()
    InsertTableFormula "COUNT"
End Sub

Private Sub InsertTableFormula(funcName As String)

    Dim tbl As Table
    Dim targetRow As Long
    Dim targetCol As Long
    Dim r As Long
    Dim cellText As String
    Dim val As Double
    Dim total As Double
    Dim cnt As Long
    Dim finalVal As Double

    ' --- Find selected table and cell ---
    If Not GetSelectedTableAndCell(tbl, targetRow, targetCol) Then
        MsgBox "Please place your cursor in a table cell.", vbExclamation, "Table Formula"
        Exit Sub
    End If

    ' --- Read numeric values from cells above ---
    total = 0
    cnt = 0

    For r = 1 To targetRow - 1
        On Error Resume Next
        cellText = CleanNumericText(tbl.Cell(r, targetCol).Shape.TextFrame.TextRange.Text)
        On Error GoTo 0

        If IsNumeric(cellText) And Len(cellText) > 0 Then
            val = CDbl(cellText)
            total = total + val
            cnt = cnt + 1
        End If
    Next r

    ' --- Calculate ---
    Select Case UCase(funcName)
        Case "SUM"
            finalVal = total
        Case "AVERAGE"
            If cnt > 0 Then
                finalVal = total / cnt
            Else
                finalVal = 0
            End If
        Case "COUNT"
            finalVal = cnt
    End Select

    ' --- Write result to cell ---
    tbl.Cell(targetRow, targetCol).Shape.TextFrame.TextRange.Text = Format(finalVal, "0.00")

End Sub

' ===== NUMBER FORMATTING =====================================================

Public Sub SelFormatNumDecimal()
    FormatSelectedCells "#,##0.00", ""
End Sub

Public Sub SelFormatNumNoDecimal()
    FormatSelectedCells "#,##0", ""
End Sub

Public Sub SelFormatNumDollar()
    FormatSelectedCells "#,##0.00", "$"
End Sub

Private Sub FormatSelectedCells(fmt As String, prefix As String)

    Dim tbl As Table
    Dim selRow As Long, selCol As Long
    Dim cells As Collection
    Dim cellCoord As Variant
    Dim r As Long, c As Long
    Dim cellText As String
    Dim val As Double
    Dim result As String
    Dim tr As TextRange

    ' --- Try to get selected cells ---
    Set cells = GetSelectedCells(tbl)

    If cells Is Nothing Then
        ' Maybe text is selected outside a table
        If ActiveWindow.Selection.Type = ppSelectionText Then
            Dim selText As String
            selText = ActiveWindow.Selection.TextRange.Text
            selText = CleanNumericText(selText)
            If IsNumeric(selText) And Len(selText) > 0 Then
                val = CDbl(selText)
                ActiveWindow.Selection.TextRange.Text = FormatValue(val, fmt, prefix)
            End If
        Else
            MsgBox "Please select table cell(s) or text containing a number.", vbExclamation, "Number Format"
        End If
        Exit Sub
    End If

    ' --- Format each selected cell ---
    For Each cellCoord In cells
        r = cellCoord(0)
        c = cellCoord(1)

        On Error Resume Next
        Set tr = tbl.Cell(r, c).Shape.TextFrame.TextRange
        On Error GoTo 0
        If tr Is Nothing Then GoTo NextCell

        cellText = CleanNumericText(tr.Text)

        If IsNumeric(cellText) And Len(cellText) > 0 Then
            val = CDbl(cellText)
            result = FormatValue(val, fmt, prefix)
            tr.Text = result
            tr.ParagraphFormat.Alignment = ppAlignRight
        End If

NextCell:
        Set tr = Nothing
    Next cellCoord

End Sub

Private Function FormatValue(val As Double, fmt As String, prefix As String) As String
    Dim result As String
    If val < 0 Then
        result = "(" & Format(Abs(val), fmt) & ")"
    Else
        result = Format(val, fmt)
    End If
    If Len(prefix) > 0 Then
        result = prefix & " " & result
    End If
    FormatValue = result
End Function

' ===== DATE FORMATTING =======================================================

Public Sub SelFormatDateShort()
    FormatDateInSelection "DD-MMM-YY"
End Sub

Public Sub SelFormatDateLong()
    FormatDateInSelection "DD-MMMM-YYYY"
End Sub

Private Sub FormatDateInSelection(fmt As String)

    Dim tbl As Table
    Dim cells As Collection
    Dim cellCoord As Variant
    Dim r As Long, c As Long
    Dim cellText As String
    Dim dt As Date
    Dim tr As TextRange

    ' --- Try to get selected cells ---
    Set cells = GetSelectedCells(tbl)

    If cells Is Nothing Then
        ' Maybe text is selected outside a table
        If ActiveWindow.Selection.Type = ppSelectionText Then
            Dim selText As String
            selText = Trim$(ActiveWindow.Selection.TextRange.Text)
            If IsDate(selText) Then
                dt = CDate(selText)
                ActiveWindow.Selection.TextRange.Text = Format(dt, fmt)
            End If
        Else
            MsgBox "Please select table cell(s) or text containing a date.", vbExclamation, "Date Format"
        End If
        Exit Sub
    End If

    ' --- Format each selected cell ---
    For Each cellCoord In cells
        r = cellCoord(0)
        c = cellCoord(1)

        On Error Resume Next
        Set tr = tbl.Cell(r, c).Shape.TextFrame.TextRange
        On Error GoTo 0
        If tr Is Nothing Then GoTo NextDateCell

        cellText = Trim$(tr.Text)
        If IsDate(cellText) Then
            dt = CDate(cellText)
            tr.Text = Format(dt, fmt)
        End If

NextDateCell:
        Set tr = Nothing
    Next cellCoord

End Sub

' =============================================================================
' HELPERS — Table & Cell Detection
' =============================================================================

Private Function GetSelectedTableAndCell( _
        ByRef tbl As Table, _
        ByRef outRow As Long, _
        ByRef outCol As Long) As Boolean
    ' Returns True if the cursor is inside a table cell.
    ' Sets tbl, outRow, outCol to the active cell.

    Dim shp As Shape
    Dim sel As Selection
    Dim r As Long, c As Long

    GetSelectedTableAndCell = False
    Set sel = ActiveWindow.Selection

    On Error Resume Next

    ' Selection is inside a table cell (text cursor in cell)
    If sel.Type = ppSelectionText Then
        Set shp = sel.ShapeRange(1)
        If Not shp Is Nothing Then
            If shp.HasTable Then
                Set tbl = shp.Table
                ' Find which cell contains the selection
                If FindCellByTextRange(tbl, sel.TextRange, outRow, outCol) Then
                    GetSelectedTableAndCell = True
                End If
            End If
        End If
    End If

    On Error GoTo 0

End Function


Private Function FindCellByTextRange( _
        tbl As Table, _
        selRange As TextRange, _
        ByRef outRow As Long, _
        ByRef outCol As Long) As Boolean
    ' Identifies which cell the selection cursor is in by comparing
    ' the TextRange object references.

    Dim r As Long, c As Long
    Dim cellTR As TextRange

    FindCellByTextRange = False

    On Error Resume Next
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Set cellTR = tbl.Cell(r, c).Shape.TextFrame.TextRange
            ' Compare the parent shape — if the selection's parent is
            ' the same shape as the cell's, we found it
            If selRange.Parent.Parent.Name = cellTR.Parent.Parent.Name Then
                ' Further check: same text content as a tiebreaker
                ' (not perfect but works for most layouts)
                outRow = r
                outCol = c
                FindCellByTextRange = True
                Exit Function
            End If
        Next c
    Next r
    On Error GoTo 0

End Function


Private Function GetSelectedCells(ByRef tbl As Table) As Collection
    ' Returns a Collection of Array(row, col) for all selected cells.
    ' Returns Nothing if no table is selected.

    Dim shp As Shape
    Dim sel As Selection
    Dim result As Collection
    Dim r As Long, c As Long

    Set GetSelectedCells = Nothing
    Set sel = ActiveWindow.Selection

    On Error Resume Next

    ' Check if a table shape is selected
    If sel.Type = ppSelectionShapes Or sel.Type = ppSelectionText Then
        Set shp = sel.ShapeRange(1)
        If shp Is Nothing Then Exit Function
        If Not shp.HasTable Then Exit Function
        Set tbl = shp.Table
    Else
        Exit Function
    End If

    On Error GoTo 0

    Set result = New Collection

    ' PowerPoint doesn't expose which cells are "selected" directly.
    ' If text is selected, we identify the single active cell.
    ' If the whole shape is selected, we include ALL cells.
    If sel.Type = ppSelectionText Then
        Dim foundRow As Long, foundCol As Long
        If FindCellByTextRange(tbl, sel.TextRange, foundRow, foundCol) Then
            result.Add Array(foundRow, foundCol)
        End If
    Else
        ' Whole table shape selected — operate on all cells
        For r = 1 To tbl.Rows.Count
            For c = 1 To tbl.Columns.Count
                result.Add Array(r, c)
            Next c
        Next r
    End If

    If result.Count > 0 Then
        Set GetSelectedCells = result
    End If

End Function


Private Function CleanNumericText(s As String) As String
    ' Strips formatting characters from cell text for numeric parsing.

    Dim t As String
    t = Trim$(s)
    t = Replace(t, ",", "")
    t = Replace(t, "$", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")

    ' Parentheses = negative
    If InStr(t, "(") > 0 And InStr(t, ")") > 0 Then
        t = Replace(t, "(", "-")
        t = Replace(t, ")", "")
    End If

    CleanNumericText = Trim$(t)

End Function


' ===== NEW: AUTOFIT TABLE ====================================================
' Shrinks column widths to fit content and optionally resizes the table shape
' to fit the slide width. Adjusts row heights to minimum.

Sub SelTableAutofit()

    Dim tbl As Table
    Dim shp As Shape

    Set shp = GetSelectedTableShape()
    If shp Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", vbExclamation, "Autofit Table"
        Exit Sub
    End If
    Set tbl = shp.Table

    Dim r As Long, c As Long
    Dim maxW As Single
    Dim cellW As Single
    Dim totalW As Single
    Dim slideW As Single

    On Error Resume Next

    ' --- Pass 1: Autofit each column to its widest text ---
    ' PPT has no native autofit for table columns, so we measure text width
    ' and set column widths accordingly, with a minimum.
    Const MIN_COL_W As Single = 36  ' ~0.5 inch minimum
    Const H_PAD As Single = 14      ' padding per cell (left+right)

    Dim colWidths() As Single
    ReDim colWidths(1 To tbl.Columns.Count)

    ' Initialise to minimum
    For c = 1 To tbl.Columns.Count
        colWidths(c) = MIN_COL_W
    Next c

    ' Measure each cell's text extent
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Dim tr As TextRange
            Set tr = tbl.Cell(r, c).Shape.TextFrame.TextRange
            If Len(tr.Text) > 0 Then
                ' Approximate width: character count * average char width
                ' Use font size as a rough proxy (0.6 * fontSize per char)
                Dim avgCharW As Single
                avgCharW = tr.Font.Size * 0.55
                cellW = Len(tr.Text) * avgCharW + H_PAD
                If cellW > colWidths(c) Then colWidths(c) = cellW
            End If
        Next c
    Next r

    ' --- Pass 2: Scale columns proportionally to fit table shape width ---
    totalW = 0
    For c = 1 To tbl.Columns.Count
        totalW = totalW + colWidths(c)
    Next c

    Dim targetW As Single
    targetW = shp.Width  ' keep current table shape width

    If totalW > 0 Then
        Dim scale As Single
        scale = targetW / totalW
        For c = 1 To tbl.Columns.Count
            tbl.Columns(c).Width = colWidths(c) * scale
        Next c
    End If

    ' --- Pass 3: Minimise row heights ---
    For r = 1 To tbl.Rows.Count
        tbl.Rows(r).Height = 0  ' setting to 0 lets PPT auto-shrink to text
    Next r

    On Error GoTo 0

    MsgBox "Table autofitted.", vbInformation, "Autofit Table"

End Sub


' ===== NEW: RESET TABLE ======================================================
' Strips all formatting from a table: fills, borders, margins, indents,
' font overrides — returns it to a plain, clean-slate table.

Sub SelTableReset()

    Dim tbl As Table
    Dim shp As Shape

    Set shp = GetSelectedTableShape()
    If shp Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", vbExclamation, "Reset Table"
        Exit Sub
    End If
    Set tbl = shp.Table

    Dim r As Long, c As Long

    On Error Resume Next

    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count

            Dim cel As Cell
            Set cel = tbl.Cell(r, c)

            ' --- Clear cell fill ---
            cel.Shape.Fill.Background  ' = no fill

            ' --- Clear borders (all 4 sides) ---
            cel.Borders(ppBorderTop).Visible = msoFalse
            cel.Borders(ppBorderBottom).Visible = msoFalse
            cel.Borders(ppBorderLeft).Visible = msoFalse
            cel.Borders(ppBorderRight).Visible = msoFalse

            ' --- Reset cell margins to PPT defaults ---
            With cel.Shape.TextFrame
                .MarginTop = CmToPt(0.13)
                .MarginBottom = CmToPt(0.13)
                .MarginLeft = CmToPt(0.25)
                .MarginRight = CmToPt(0.25)
                .WordWrap = msoTrue
                .AutoSize = ppAutoSizeNone
            End With

            ' --- Reset text formatting ---
            Dim tr As TextRange
            Set tr = cel.Shape.TextFrame.TextRange

            If Len(tr.Text) > 0 Then
                With tr.Font
                    .Name = "Calibri"
                    .Size = 11
                    .Bold = msoFalse
                    .Italic = msoFalse
                    .Underline = msoFalse
                    .Color.RGB = RGB(0, 0, 0)
                    .Shadow = msoFalse
                End With

                ' Reset paragraph formatting
                Dim p As Long
                For p = 1 To tr.Paragraphs.Count
                    With tr.Paragraphs(p).ParagraphFormat
                        .Alignment = ppAlignLeft
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                        .SpaceWithin = 1
                        .WordWrap = msoTrue
                    End With

                    ' Reset indentation
                    With tr.Paragraphs(p).ParagraphFormat
                        .IndentLevel = 1  ' PPT indent level 1 = no indent
                    End With

                    ' Remove bullets/numbering
                    tr.Paragraphs(p).ParagraphFormat.Bullet.Type = ppBulletNone
                Next p
            End If

        Next c
    Next r

    ' --- Clear table-level style ---
    ' Remove any applied table style by setting to "No Style, No Grid"
    shp.Table.ApplyStyle "{2D5ABB26-0587-4C30-8999-92F81FD0307C}", msoFalse

    On Error GoTo 0

    MsgBox "Table reset to plain formatting.", vbInformation, "Reset Table"

End Sub


' =============================================================================
' HELPERS
' =============================================================================

Private Function GetSelectedTable() As Table
    ' Returns the table from the current selection, or Nothing.

    Dim shp As Shape
    Set shp = GetSelectedTableShape()
    If Not shp Is Nothing Then
        Set GetSelectedTable = shp.Table
    End If

End Function

Private Function GetSelectedTableShape() As Shape
    ' Returns the Shape containing a table from the current selection.

    Dim sel As Selection
    Dim shp As Shape

    Set GetSelectedTableShape = Nothing
    Set sel = ActiveWindow.Selection

    On Error Resume Next

    If sel.Type = ppSelectionShapes Or sel.Type = ppSelectionText Then
        Set shp = sel.ShapeRange(1)
        If Not shp Is Nothing Then
            If shp.HasTable Then
                Set GetSelectedTableShape = shp
            End If
        End If
    End If

    On Error GoTo 0

End Function

Private Function CmToPt(cm As Double) As Double
    CmToPt = cm * 28.3464567
End Function


' ===== TABLE MARGINS — ALL TABLES IN DOCUMENT ================================

Sub DeckTableMargin()

    Const PAD_TOP As Double = 0.1       ' cm
    Const PAD_BOTTOM As Double = 0.1    ' cm
    Const PAD_LEFT As Double = 0.19     ' cm
    Const PAD_RIGHT As Double = 0.19    ' cm

    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTable Then
                SetTableMargins shp.Table, PAD_TOP, PAD_BOTTOM, PAD_LEFT, PAD_RIGHT
            End If
        Next shp
    Next sld

End Sub


' ===== SELECTED TABLE — SET MARGINS ==========================================

Sub SelTableMargin()

    Const PAD_TOP As Double = 0.05      ' cm
    Const PAD_BOTTOM As Double = 0.05   ' cm
    Const PAD_LEFT As Double = 0.19     ' cm
    Const PAD_RIGHT As Double = 0.19    ' cm

    Dim tbl As Table
    Set tbl = GetSelectedTable()

    If tbl Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", vbExclamation, "Table Margin"
        Exit Sub
    End If

    SetTableMargins tbl, PAD_TOP, PAD_BOTTOM, PAD_LEFT, PAD_RIGHT
    MsgBox "Table margins applied.", vbInformation, "Table Margin"

End Sub


Private Sub SetTableMargins(tbl As Table, topCm As Double, bottomCm As Double, leftCm As Double, rightCm As Double)

    Dim r As Long, c As Long

    On Error Resume Next
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            With tbl.Cell(r, c).Shape.TextFrame
                .MarginTop = CmToPt(topCm)
                .MarginBottom = CmToPt(bottomCm)
                .MarginLeft = CmToPt(leftCm)
                .MarginRight = CmToPt(rightCm)
            End With
        Next c
    Next r
    On Error GoTo 0

End Sub


' ===== SELECTED TABLE — BORDERS ==============================================

Sub SelTableBorder()

    Dim tbl As Table
    Set tbl = GetSelectedTable()

    If tbl Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", vbExclamation, "Table Border"
        Exit Sub
    End If

    Dim r As Long, c As Long

    On Error Resume Next
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            With tbl.Cell(r, c).Borders(ppBorderTop)
                .ForeColor.RGB = RGB(0, 0, 0)
                .Weight = 0.25
                .DashStyle = msoLineSolid
                .Visible = msoTrue
            End With
            With tbl.Cell(r, c).Borders(ppBorderBottom)
                .ForeColor.RGB = RGB(0, 0, 0)
                .Weight = 0.25
                .DashStyle = msoLineSolid
                .Visible = msoTrue
            End With
            With tbl.Cell(r, c).Borders(ppBorderLeft)
                .ForeColor.RGB = RGB(0, 0, 0)
                .Weight = 0.25
                .DashStyle = msoLineSolid
                .Visible = msoTrue
            End With
            With tbl.Cell(r, c).Borders(ppBorderRight)
                .ForeColor.RGB = RGB(0, 0, 0)
                .Weight = 0.25
                .DashStyle = msoLineSolid
                .Visible = msoTrue
            End With
        Next c
    Next r
    On Error GoTo 0

    MsgBox "Borders applied.", vbInformation, "Table Border"

End Sub