Attribute VB_Name = "Subs"
Option Explicit

' =============================================================================
' MODULE: Subs — PowerPoint Edition
' =============================================================================
'   - DeckFontSizeDecrease / DeckFontSizeIncrease
'   - DeckSpacingSingle
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

Sub DeckFontSizeDecrease()
    ChangeFontSizeAll -1
End Sub

Sub DeckFontSizeIncrease()
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

Sub DeckSpacingSingle()

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

' ===== NUMBER FORMATTING — WORKS ON ANY SELECTED TEXT ========================

Public Sub SelFormatNumDecimal()
    FormatSelectedNumbers "#,##0.00", ""
End Sub

Public Sub SelFormatNumNoDecimal()
    FormatSelectedNumbers "#,##0", ""
End Sub

Public Sub SelFormatNumDollar()
    FormatSelectedNumbers "#,##0.00", "$"
End Sub

Private Sub FormatSelectedNumbers(fmt As String, prefix As String)

    Dim sel As Selection
    Set sel = ActiveWindow.Selection

    ' --- Text is highlighted (works in text box, table cell, title, etc.) ---
    If sel.Type = ppSelectionText Then
        FormatNumInTextRange sel.TextRange, fmt, prefix
        Exit Sub
    End If

    ' --- Whole shape(s) selected — format all text inside ---
    If sel.Type = ppSelectionShapes Then
        Dim i As Long
        For i = 1 To sel.ShapeRange.Count
            FormatNumInShape sel.ShapeRange(i), fmt, prefix
        Next i
        Exit Sub
    End If

    MsgBox "Please select text or a shape containing numbers.", vbExclamation, "Number Format"

End Sub

Private Sub FormatNumInShape(shp As Shape, fmt As String, prefix As String)

    On Error Resume Next

    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            FormatNumInShape subShp, fmt, prefix
        Next subShp
        Exit Sub
    End If

    If shp.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                FormatNumInTextRange shp.Table.Cell(r, c).Shape.TextFrame.TextRange, fmt, prefix
            Next c
        Next r
        Exit Sub
    End If

    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            FormatNumInTextRange shp.TextFrame.TextRange, fmt, prefix
        End If
    End If

    On Error GoTo 0

End Sub

Private Sub FormatNumInTextRange(tr As TextRange, fmt As String, prefix As String)

    On Error Resume Next

    ' Try the whole range as a single number first
    Dim txt As String
    txt = CleanNumericText(tr.Text)

    If IsNumeric(txt) And Len(txt) > 0 Then
        Dim val As Double
        val = CDbl(txt)
        tr.Text = FormatValue(val, fmt, prefix)
        tr.ParagraphFormat.Alignment = ppAlignRight
        Exit Sub
    End If

    ' Otherwise try each paragraph individually
    Dim p As Long
    For p = tr.Paragraphs.Count To 1 Step -1
        Dim pTxt As String
        pTxt = CleanNumericText(tr.Paragraphs(p).Text)
        If IsNumeric(pTxt) And Len(pTxt) > 0 Then
            tr.Paragraphs(p).Text = FormatValue(CDbl(pTxt), fmt, prefix)
            tr.Paragraphs(p).ParagraphFormat.Alignment = ppAlignRight
        End If
    Next p

    On Error GoTo 0

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

' ===== DATE FORMATTING — WORKS ON ANY SELECTED TEXT ==========================

Public Sub SelFormatDateShort()
    FormatSelectedDates "DD-MMM-YY"
End Sub

Public Sub SelFormatDateLong()
    FormatSelectedDates "DD-MMMM-YYYY"
End Sub

Private Sub FormatSelectedDates(fmt As String)

    Dim sel As Selection
    Set sel = ActiveWindow.Selection

    ' --- Text is highlighted ---
    If sel.Type = ppSelectionText Then
        FormatDateInTextRange sel.TextRange, fmt
        Exit Sub
    End If

    ' --- Whole shape(s) selected ---
    If sel.Type = ppSelectionShapes Then
        Dim i As Long
        For i = 1 To sel.ShapeRange.Count
            FormatDateInShape sel.ShapeRange(i), fmt
        Next i
        Exit Sub
    End If

    MsgBox "Please select text or a shape containing dates.", vbExclamation, "Date Format"

End Sub

Private Sub FormatDateInShape(shp As Shape, fmt As String)

    On Error Resume Next

    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            FormatDateInShape subShp, fmt
        Next subShp
        Exit Sub
    End If

    If shp.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                FormatDateInTextRange shp.Table.Cell(r, c).Shape.TextFrame.TextRange, fmt
            Next c
        Next r
        Exit Sub
    End If

    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            FormatDateInTextRange shp.TextFrame.TextRange, fmt
        End If
    End If

    On Error GoTo 0

End Sub

Private Sub FormatDateInTextRange(tr As TextRange, fmt As String)

    On Error Resume Next

    ' Try the whole range as a single date
    Dim txt As String
    txt = Trim$(tr.Text)

    If IsDate(txt) Then
        tr.Text = Format(CDate(txt), fmt)
        Exit Sub
    End If

    ' Otherwise try each paragraph
    Dim p As Long
    For p = tr.Paragraphs.Count To 1 Step -1
        Dim pTxt As String
        pTxt = Trim$(tr.Paragraphs(p).Text)
        If IsDate(pTxt) Then
            tr.Paragraphs(p).Text = Format(CDate(pTxt), fmt)
        End If
    Next p

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
    If InStr(t, "(") > 0 And InStr(t, ")") > 0 Then
        t = Replace(t, "(", "-")
        t = Replace(t, ")", "")
    End If
    CleanNumericText = Trim$(t)
End Function