Option Explicit

' 1 cm = 28.3464567 points (PPT has no built-in CentimetersToPoints)
Private Function CmToPt(cm As Double) As Double
    CmToPt = cm * 28.3464567
End Function

' =============================================================================
' MODULE: SubTable
' Purpose: Table-specific operations for PowerPoint (formulas, borders, margins)
'          Requires cursor/selection to be inside a table.
'
' Contents:
'   - SelSumColumn / SelAverageColumn / SelCountColumn
'   - SelTableBorder
'   - SelTableMargin
'   - DeckTableMargin
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

    Dim sel As Selection
    Dim shp As Shape
    Dim tbl As Table
    Dim targetRow As Long
    Dim targetCol As Long
    Dim r As Long
    Dim cellText As String
    Dim val As Double
    Dim total As Double
    Dim cnt As Long
    Dim finalVal As Double

    Set sel = ActiveWindow.Selection

    ' We need to find the active table and the cell the user is in.
    ' PPT doesn't expose the "active cell" directly, so we look for
    ' a text selection inside a table shape.

    If sel.Type <> ppSelectionText Then
        MsgBox "Please click inside a table cell, then run this command.", vbExclamation
        Exit Sub
    End If

    ' Walk up to find the table shape
    Set shp = sel.ShapeRange(1)

    If Not shp.HasTable Then
        ' The parent of the text frame might be a cell shape inside a table
        ' Try the parent shape
        On Error GoTo notTable
        Set shp = shp.ParentGroup
        On Error GoTo 0
        If Not shp.HasTable Then GoTo notTable
    End If

    Set tbl = shp.Table

    ' Find which cell contains the selection by matching text range
    Dim tr As TextRange
    Set tr = sel.TextRange

    ' Determine cell position by iterating
    Dim found As Boolean
    found = False
    For r = 1 To tbl.Rows.Count
        Dim c As Long
        For c = 1 To tbl.Columns.Count
            If tbl.Cell(r, c).Shape.TextFrame.TextRange.Text = tr.Text Then
                ' Heuristic: if the text matches, assume it's our cell
                ' (works well when user clicks into a specific cell)
                targetRow = r
                targetCol = c
                found = True
                Exit For
            End If
        Next c
        If found Then Exit For
    Next r

    If Not found Then
        ' Fallback: use last row, last column with selection
        MsgBox "Could not determine the active cell. " & _
               "Please place your cursor in the target cell.", vbExclamation
        Exit Sub
    End If

    ' Calculate from rows above the target
    total = 0
    cnt = 0

    For r = 1 To targetRow - 1
        cellText = Trim$(tbl.Cell(r, targetCol).Shape.TextFrame.TextRange.Text)
        cellText = CleanNumericText(cellText)
        If IsNumeric(cellText) And Len(cellText) > 0 Then
            val = CDbl(cellText)
            total = total + val
            cnt = cnt + 1
        End If
    Next r

    ' Calculate result
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

    ' Write result into the target cell
    tbl.Cell(targetRow, targetCol).Shape.TextFrame.TextRange.Text = Format(finalVal, "0.00")

    Exit Sub

notTable:
    MsgBox "Please place your cursor inside a table cell.", vbExclamation

End Sub


' ===== BORDERS ===============================================================

Sub SelTableBorder()

    Dim sel As Selection
    Dim shp As Shape
    Dim tbl As Table
    Dim r As Long, c As Long
    Dim cel As Cell
    Dim bdr As Long

    Set sel = ActiveWindow.Selection

    If sel.Type < ppSelectionShapes Then
        MsgBox "Please select a table shape.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set shp = sel.ShapeRange(1)
    On Error GoTo 0

    If shp Is Nothing Then
        MsgBox "Please select a table shape.", vbExclamation
        Exit Sub
    End If

    If Not shp.HasTable Then
        MsgBox "The selected shape is not a table.", vbExclamation
        Exit Sub
    End If

    Set tbl = shp.Table

    On Error Resume Next
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Set cel = tbl.Cell(r, c)
            For bdr = ppBorderTop To ppBorderRight
                With cel.Borders(bdr)
                    .ForeColor.RGB = RGB(0, 0, 0)
                    .Weight = 0.5
                    .Visible = msoTrue
                    .DashStyle = msoLineSolid
                End With
            Next bdr
            ' Hide diagonals
            cel.Borders(ppBorderDiagonalDown).Visible = msoFalse
            cel.Borders(ppBorderDiagonalUp).Visible = msoFalse
        Next c
    Next r
    On Error GoTo 0

End Sub


' ===== MARGINS — SELECTED TABLE ==============================================

Sub SelTableMargin()

    Const PAD_TOP_CM As Double = 0.05
    Const PAD_BOTTOM_CM As Double = 0.05
    Const PAD_LEFT_CM As Double = 0.19
    Const PAD_RIGHT_CM As Double = 0.19

    Dim sel As Selection
    Dim shp As Shape
    Dim tbl As Table
    Dim r As Long, c As Long

    Set sel = ActiveWindow.Selection

    If sel.Type < ppSelectionShapes Then
        MsgBox "Please select a table shape.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set shp = sel.ShapeRange(1)
    On Error GoTo 0

    If shp Is Nothing Or Not shp.HasTable Then
        MsgBox "Please select a table shape.", vbExclamation
        Exit Sub
    End If

    Set tbl = shp.Table

    On Error Resume Next
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            With tbl.Cell(r, c).Shape.TextFrame2
                .MarginTop = CmToPt(PAD_TOP_CM)
                .MarginBottom = CmToPt(PAD_BOTTOM_CM)
                .MarginLeft = CmToPt(PAD_LEFT_CM)
                .MarginRight = CmToPt(PAD_RIGHT_CM)
            End With
        Next c
    Next r
    On Error GoTo 0

End Sub


' ===== MARGINS — ALL TABLES IN PRESENTATION ==================================

Sub DeckTableMargin()

    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim r As Long, c As Long

    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTable Then
                Set tbl = shp.Table
                For r = 1 To tbl.Rows.Count
                    For c = 1 To tbl.Columns.Count
                        With tbl.Cell(r, c).Shape.TextFrame2
                            .MarginTop = CmToPt(0.1)
                            .MarginBottom = CmToPt(0.1)
                            .MarginLeft = CmToPt(0.19)
                            .MarginRight = CmToPt(0.19)
                        End With
                    Next c
                Next r
            End If
        Next shp
    Next sld
    On Error GoTo 0

End Sub


' ===== HELPER ================================================================

Private Function CleanNumericText(s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, ",", "")
    t = Replace(t, "$", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, Chr(13), "")
    If InStr(t, "(") > 0 And InStr(t, ")") > 0 Then
        t = Replace(t, "(", "-")
        t = Replace(t, ")", "")
    End If
    CleanNumericText = Trim$(t)
End Function
