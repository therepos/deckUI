Attribute VB_Name = "SubTable"
Option Explicit

' =============================================================================
' MODULE: SubTable — PowerPoint Edition
' Purpose: Table-specific operations (formulas, borders, margins, autofit, reset)
'          Requires a table to be selected.
'
' Contents:
'   - SelSumColumn / SelAverageColumn / SelCountColumn
'   - SelTableBorder
'   - SelTableMargin
'   - DocTableMargin
'   - SelTableAutofit
'   - SelTableReset
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

    If Not GetSelectedTableAndCell(tbl, targetRow, targetCol) Then
        MsgBox "Please place your cursor in a table cell.", _
               vbExclamation, "Table Formula"
        Exit Sub
    End If

    ' Read numeric values from cells above
    total = 0
    cnt = 0

    For r = 1 To targetRow - 1
        On Error Resume Next
        cellText = CleanNumericText( _
            tbl.Cell(r, targetCol).Shape.TextFrame.TextRange.Text)
        On Error GoTo 0

        If IsNumeric(cellText) And Len(cellText) > 0 Then
            val = CDbl(cellText)
            total = total + val
            cnt = cnt + 1
        End If
    Next r

    ' Calculate
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

    ' Write result
    tbl.Cell(targetRow, targetCol).Shape.TextFrame.TextRange.Text = _
        Format(finalVal, "0.00")

End Sub


' ===== BORDERS ===============================================================

Sub SelTableBorder()

    Dim tbl As Table
    Set tbl = GetSelectedTable()

    If tbl Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", _
               vbExclamation, "Table Border"
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


' ===== MARGINS — SELECTED TABLE ==============================================

Sub SelTableMargin()

    Const PAD_TOP As Double = 0.05
    Const PAD_BOTTOM As Double = 0.05
    Const PAD_LEFT As Double = 0.19
    Const PAD_RIGHT As Double = 0.19

    Dim tbl As Table
    Set tbl = GetSelectedTable()

    If tbl Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", _
               vbExclamation, "Table Margin"
        Exit Sub
    End If

    SetTableMargins tbl, PAD_TOP, PAD_BOTTOM, PAD_LEFT, PAD_RIGHT
    MsgBox "Table margins applied.", vbInformation, "Table Margin"

End Sub


' ===== MARGINS — ALL TABLES IN PRESENTATION ==================================

Sub DocTableMargin()

    Const PAD_TOP As Double = 0.1
    Const PAD_BOTTOM As Double = 0.1
    Const PAD_LEFT As Double = 0.19
    Const PAD_RIGHT As Double = 0.19

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

Private Sub SetTableMargins(tbl As Table, topCm As Double, bottomCm As Double, _
                            leftCm As Double, rightCm As Double)
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


' ===== AUTOFIT TABLE =========================================================

Sub SelTableAutofit()

    Dim tbl As Table
    Dim shp As Shape

    Set shp = GetSelectedTableShape()
    If shp Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", _
               vbExclamation, "Autofit Table"
        Exit Sub
    End If
    Set tbl = shp.Table

    Dim r As Long, c As Long
    Dim cellW As Single
    Dim totalW As Single

    On Error Resume Next

    Const MIN_COL_W As Single = 36
    Const H_PAD As Single = 14

    Dim colWidths() As Single
    ReDim colWidths(1 To tbl.Columns.Count)

    For c = 1 To tbl.Columns.Count
        colWidths(c) = MIN_COL_W
    Next c

    ' Measure each cell's text extent
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Dim tr As TextRange
            Set tr = tbl.Cell(r, c).Shape.TextFrame.TextRange
            If Len(tr.Text) > 0 Then
                Dim avgCharW As Single
                avgCharW = tr.Font.Size * 0.55
                cellW = Len(tr.Text) * avgCharW + H_PAD
                If cellW > colWidths(c) Then colWidths(c) = cellW
            End If
        Next c
    Next r

    ' Scale columns proportionally to fit table shape width
    totalW = 0
    For c = 1 To tbl.Columns.Count
        totalW = totalW + colWidths(c)
    Next c

    If totalW > 0 Then
        Dim scale As Single
        scale = shp.Width / totalW
        For c = 1 To tbl.Columns.Count
            tbl.Columns(c).Width = colWidths(c) * scale
        Next c
    End If

    ' Minimise row heights
    For r = 1 To tbl.Rows.Count
        tbl.Rows(r).Height = 0
    Next r

    On Error GoTo 0

    MsgBox "Table autofitted.", vbInformation, "Autofit Table"

End Sub


' ===== RESET TABLE ===========================================================

Sub SelTableReset()

    Dim tbl As Table
    Dim shp As Shape

    Set shp = GetSelectedTableShape()
    If shp Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", _
               vbExclamation, "Reset Table"
        Exit Sub
    End If
    Set tbl = shp.Table

    Dim r As Long, c As Long

    On Error Resume Next

    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count

            Dim cel As Cell
            Set cel = tbl.Cell(r, c)

            ' Clear cell fill
            cel.Shape.Fill.Background

            ' Clear borders
            cel.Borders(ppBorderTop).Visible = msoFalse
            cel.Borders(ppBorderBottom).Visible = msoFalse
            cel.Borders(ppBorderLeft).Visible = msoFalse
            cel.Borders(ppBorderRight).Visible = msoFalse

            ' Reset cell margins to PPT defaults
            With cel.Shape.TextFrame
                .MarginTop = CmToPt(0.13)
                .MarginBottom = CmToPt(0.13)
                .MarginLeft = CmToPt(0.25)
                .MarginRight = CmToPt(0.25)
                .WordWrap = msoTrue
                .AutoSize = ppAutoSizeNone
            End With

            ' Reset text formatting
            Dim cellTR As TextRange
            Set cellTR = cel.Shape.TextFrame.TextRange

            If Len(cellTR.Text) > 0 Then
                With cellTR.Font
                    .Name = "Calibri"
                    .Size = 11
                    .Bold = msoFalse
                    .Italic = msoFalse
                    .Underline = msoFalse
                    .Color.RGB = RGB(0, 0, 0)
                    .Shadow = msoFalse
                End With

                Dim p As Long
                For p = 1 To cellTR.Paragraphs.Count
                    With cellTR.Paragraphs(p).ParagraphFormat
                        .Alignment = ppAlignLeft
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                        .SpaceWithin = 1
                        .WordWrap = msoTrue
                        .IndentLevel = 1
                        .Bullet.Type = ppBulletNone
                    End With
                Next p
            End If

        Next c
    Next r

    ' Clear table-level style
    shp.Table.ApplyStyle "{2D5ABB26-0587-4C30-8999-92F81FD0307C}", msoFalse

    On Error GoTo 0

    MsgBox "Table reset to plain formatting.", vbInformation, "Reset Table"

End Sub


' =============================================================================
' HELPERS
' =============================================================================

Private Function GetSelectedTable() As Table
    Dim shp As Shape
    Set shp = GetSelectedTableShape()
    If Not shp Is Nothing Then
        Set GetSelectedTable = shp.Table
    End If
End Function

Private Function GetSelectedTableShape() As Shape
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

Private Function GetSelectedTableAndCell( _
        ByRef tbl As Table, _
        ByRef outRow As Long, _
        ByRef outCol As Long) As Boolean

    Dim shp As Shape
    Dim sel As Selection

    GetSelectedTableAndCell = False
    Set sel = ActiveWindow.Selection

    On Error Resume Next
    If sel.Type = ppSelectionText Then
        Set shp = sel.ShapeRange(1)
        If Not shp Is Nothing Then
            If shp.HasTable Then
                Set tbl = shp.Table
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

    Dim r As Long, c As Long
    Dim cellTR As TextRange

    FindCellByTextRange = False

    On Error Resume Next
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Set cellTR = tbl.Cell(r, c).Shape.TextFrame.TextRange
            If selRange.Parent.Parent.Name = cellTR.Parent.Parent.Name Then
                outRow = r
                outCol = c
                FindCellByTextRange = True
                Exit Function
            End If
        Next c
    Next r
    On Error GoTo 0

End Function

Private Function CleanNumericText(s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, ",", "")
    t = Replace(t, "$", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    If InStr(t, "(") > 0 And InStr(t, ")") > 0 Then
        t = Replace(t, "(", "-")
        t = Replace(t, ")", "")
    End If
    CleanNumericText = Trim$(t)
End Function

Private Function CmToPt(cm As Double) As Double
    CmToPt = cm * 28.3464567
End Function
