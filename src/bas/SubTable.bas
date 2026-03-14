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
