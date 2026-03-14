Attribute VB_Name = "Subs"
Option Explicit

' =============================================================================
' MODULE: Subs — PowerPoint Edition
' =============================================================================
' Ported from Word:
'   - DocFontSizeDecrease / DocFontSizeIncrease
'   - DocSpacingSingle
'   - DocTableMargin
'   - SelTableBorder
'   - SelTableMargin
'   - SelListAlphaRoman  (NOT ported — PPT has no list templates)
'   - ViewSplitVerticalToggle  (NOT ported — PPT has no split view)
'
' New PowerPoint-only features:
'   - SelTableAutofit
'   - SelTableReset
' =============================================================================


' ===== FONT SIZE — ALL SHAPES ON ALL SLIDES ==================================

Sub DocFontSizeDecrease()
    ChangeFontSizeAll -1
End Sub

Sub DocFontSizeIncrease()
    ChangeFontSizeAll 1
End Sub

Private Sub ChangeFontSizeAll(delta As Single)

    Dim sld As Slide
    Dim shp As Shape
    Dim slideNum As Long

    For Each sld In ActivePresentation.Slides
        slideNum = slideNum + 1
        Application.StatusBar = "Resizing fonts — slide " & slideNum & " of " & ActivePresentation.Slides.Count & "..."

        For Each shp In sld.Shapes
            ChangeFontSizeInShape shp, delta
        Next shp
    Next sld

    ' Masters & layouts
    Dim dsgn As Design
    For Each dsgn In ActivePresentation.Designs
        For Each shp In dsgn.SlideMaster.Shapes
            ChangeFontSizeInShape shp, delta
        Next shp
        Dim lay As CustomLayout
        For Each lay In dsgn.SlideMaster.CustomLayouts
            For Each shp In lay.Shapes
                ChangeFontSizeInShape shp, delta
            Next shp
        Next lay
    Next dsgn

    Application.StatusBar = ""

End Sub

Private Sub ChangeFontSizeInShape(shp As Shape, delta As Single)

    On Error Resume Next

    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            ChangeFontSizeInShape subShp, delta
        Next subShp
        Exit Sub
    End If

    If shp.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                ChangeFontSizeInTextFrame shp.Table.Cell(r, c).Shape.TextFrame, delta
            Next c
        Next r
        Exit Sub
    End If

    If shp.HasTextFrame Then
        ChangeFontSizeInTextFrame shp.TextFrame, delta
    End If

    On Error GoTo 0

End Sub

Private Sub ChangeFontSizeInTextFrame(tf As TextFrame, delta As Single)

    On Error Resume Next
    If Not tf.HasText Then Exit Sub

    ' Change per-run so mixed sizes are preserved
    Dim i As Long
    For i = 1 To tf.TextRange.Runs.Count
        Dim sz As Single
        sz = tf.TextRange.Runs(i).Font.Size
        If sz + delta >= 1 Then
            tf.TextRange.Runs(i).Font.Size = sz + delta
        End If
    Next i

    On Error GoTo 0

End Sub


' ===== SINGLE SPACING — ALL SHAPES ON ALL SLIDES ============================

Sub DocSpacingSingle()

    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            ApplySpacingSingle shp
        Next shp
    Next sld

    ' Masters & layouts
    Dim dsgn As Design
    For Each dsgn In ActivePresentation.Designs
        For Each shp In dsgn.SlideMaster.Shapes
            ApplySpacingSingle shp
        Next shp
        Dim lay As CustomLayout
        For Each lay In dsgn.SlideMaster.CustomLayouts
            For Each shp In lay.Shapes
                ApplySpacingSingle shp
            Next shp
        Next lay
    Next dsgn

End Sub

Private Sub ApplySpacingSingle(shp As Shape)

    On Error Resume Next

    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            ApplySpacingSingle subShp
        Next subShp
        Exit Sub
    End If

    If shp.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                SetSpacingSingleTF shp.Table.Cell(r, c).Shape.TextFrame
            Next c
        Next r
        Exit Sub
    End If

    If shp.HasTextFrame Then
        SetSpacingSingleTF shp.TextFrame
    End If

    On Error GoTo 0

End Sub

Private Sub SetSpacingSingleTF(tf As TextFrame)
    On Error Resume Next
    If Not tf.HasText Then Exit Sub

    Dim i As Long
    For i = 1 To tf.TextRange.Paragraphs.Count
        With tf.TextRange.Paragraphs(i).ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .SpaceWithin = 1  ' single line spacing
        End With
    Next i

    On Error GoTo 0
End Sub


' ===== TABLE MARGINS — ALL TABLES IN DOCUMENT ================================

Sub DocTableMargin()

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
