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

Private Sub InsertTableFormula(ByVal funcName As String)

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
        MsgBox "Please place your cursor in a table cell.", vbExclamation, "Table Formula"
        Exit Sub
    End If

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

    Select Case UCase$(funcName)
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
        Case Else
            finalVal = 0
    End Select

    tbl.Cell(targetRow, targetCol).Shape.TextFrame.TextRange.Text = Format$(finalVal, "0.00")

End Sub


' ===== BORDERS ===============================================================

Public Sub SelTableBorder()

    Dim tbl As Table
    Dim r As Long
    Dim c As Long

    Set tbl = GetSelectedTable()

    If tbl Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", vbExclamation, "Table Border"
        Exit Sub
    End If

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

End Sub


' ===== MARGINS — SELECTED TABLE ==============================================

Public Sub SelTableMargin()

    Const PAD_TOP As Double = 0.05
    Const PAD_BOTTOM As Double = 0.05
    Const PAD_LEFT As Double = 0.19
    Const PAD_RIGHT As Double = 0.19

    Dim tbl As Table

    Set tbl = GetSelectedTable()

    If tbl Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", vbExclamation, "Table Margin"
        Exit Sub
    End If

    SetTableMargins tbl, PAD_TOP, PAD_BOTTOM, PAD_LEFT, PAD_RIGHT

End Sub


' ===== MARGINS — ALL TABLES IN PRESENTATION ==================================

Public Sub DocTableMargin()

    Const PAD_TOP As Double = 0.1
    Const PAD_BOTTOM As Double = 0.1
    Const PAD_LEFT As Double = 0.19
    Const PAD_RIGHT As Double = 0.19

    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            On Error Resume Next
            If shp.HasTable Then
                SetTableMargins shp.Table, PAD_TOP, PAD_BOTTOM, PAD_LEFT, PAD_RIGHT
            End If
            On Error GoTo 0
        Next shp
    Next sld

End Sub

Private Sub SetTableMargins(ByVal tbl As Table, ByVal topCm As Double, ByVal bottomCm As Double, ByVal leftCm As Double, ByVal rightCm As Double)

    Dim r As Long
    Dim c As Long

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

Public Sub SelTableAutofit()

    Const MIN_COL_W As Single = 36
    Const H_PAD As Single = 14

    Dim tbl As Table
    Dim shp As Shape
    Dim r As Long
    Dim c As Long
    Dim cellW As Single
    Dim totalW As Single
    Dim widthScale As Single
    Dim newW As Single
    Dim avgCharW As Single
    Dim tr As TextRange
    Dim colWidths() As Single

    Set shp = GetSelectedTableShape()
    If shp Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", vbExclamation, "Autofit Table"
        Exit Sub
    End If

    Set tbl = shp.Table

    ReDim colWidths(1 To tbl.Columns.Count)

    For c = 1 To tbl.Columns.Count
        colWidths(c) = MIN_COL_W
    Next c

    On Error Resume Next

    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Set tr = tbl.Cell(r, c).Shape.TextFrame.TextRange
            If Len(tr.Text) > 0 Then
                avgCharW = tr.Font.Size * 0.55
                cellW = (Len(tr.Text) * avgCharW) + H_PAD
                If cellW > colWidths(c) Then
                    colWidths(c) = cellW
                End If
            End If
        Next c
    Next r

    totalW = 0
    For c = 1 To tbl.Columns.Count
        totalW = totalW + colWidths(c)
    Next c

    If totalW > 0 Then
        widthScale = shp.Width / totalW

        For c = 1 To tbl.Columns.Count
            newW = colWidths(c) * widthScale
            tbl.Columns(c).Width = newW
        Next c
    End If

    For r = 1 To tbl.Rows.Count
        tbl.Rows(r).Height = 0
    Next r

    On Error GoTo 0

End Sub


' ===== RESET TABLE ===========================================================

Public Sub SelTableReset()

    Dim shp As Shape
    Dim tbl As Table
    Dim cel As Cell
    Dim cellTR As TextRange
    Dim r As Long
    Dim c As Long
    Dim p As Long

    Set shp = GetSelectedTableShape()
    If shp Is Nothing Then
        MsgBox "Please select a table or place your cursor inside one.", vbExclamation, "Reset Table"
        Exit Sub
    End If

    Set tbl = shp.Table

    On Error Resume Next

    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count

            Set cel = tbl.Cell(r, c)

            cel.Shape.Fill.Visible = msoFalse

            cel.Borders(ppBorderTop).Visible = msoFalse
            cel.Borders(ppBorderBottom).Visible = msoFalse
            cel.Borders(ppBorderLeft).Visible = msoFalse
            cel.Borders(ppBorderRight).Visible = msoFalse

            With cel.Shape.TextFrame
                .MarginTop = CmToPt(0.13)
                .MarginBottom = CmToPt(0.13)
                .MarginLeft = CmToPt(0.25)
                .MarginRight = CmToPt(0.25)
                .WordWrap = msoTrue
                .AutoSize = ppAutoSizeNone
            End With

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

                For p = 1 To cellTR.Paragraphs.Count
                    With cellTR.Paragraphs(p).ParagraphFormat
                        .Alignment = ppAlignLeft
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                        .SpaceWithin = 1
                        .Bullet.Type = ppBulletNone
                    End With
                    cellTR.Paragraphs(p).IndentLevel = 1
                Next p
            End If

        Next c
    Next r

    On Error GoTo 0

End Sub


' =============================================================================
' HELPERS
' =============================================================================

Private Function GetSelectedTable() As Table

    Dim shp As Shape

    Set shp = GetSelectedTableShape()
    If Not shp Is Nothing Then
        Set GetSelectedTable = shp.Table
    Else
        Set GetSelectedTable = Nothing
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

    Dim sel As Selection
    Dim shp As Shape
    Dim r As Long
    Dim c As Long
    Dim selTR As TextRange
    Dim cellTR As TextRange

    GetSelectedTableAndCell = False
    Set sel = ActiveWindow.Selection

    On Error GoTo Fail

    If sel.Type <> ppSelectionText Then Exit Function

    Set shp = sel.ShapeRange(1)
    If shp Is Nothing Then Exit Function
    If Not shp.HasTable Then Exit Function

    Set tbl = shp.Table
    Set selTR = sel.TextRange

    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Set cellTR = tbl.Cell(r, c).Shape.TextFrame.TextRange
            If selTR.Parent Is cellTR.Parent Then
                outRow = r
                outCol = c
                GetSelectedTableAndCell = True
                Exit Function
            End If
        Next c
    Next r

Fail:
End Function

Private Function CleanNumericText(ByVal s As String) As String

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

Private Function CmToPt(ByVal cm As Double) As Double
    CmToPt = cm * 28.3464567
End Function
