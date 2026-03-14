Option Explicit

' =============================================================================
' MODULE: Subs — PowerPoint Edition
' Purpose: General formatting operations that work anywhere in the presentation
'          (text boxes, titles, table cells, notes, masters, etc.)
'
' Contents:
'   - DocFontSizeDecrease / DocFontSizeIncrease
'   - DocSpacingSingle
'   - RunPresetFontArial / RunPresetFontEY / RunPresetFontTimes /
'     RunPresetFontCalibri / RunPresetFontRepeat
'   - SelFormatNumDecimal / SelFormatNumNoDecimal / SelFormatNumDollar /
'     SelFormatNumRepeat
'   - SelFormatDateShort / SelFormatDateLong / SelFormatDateRepeat
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

    For Each sld In ActivePresentation.Slides
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
                ChangeFontSizeInTF shp.Table.Cell(r, c).Shape.TextFrame, delta
            Next c
        Next r
        Exit Sub
    End If

    If shp.HasTextFrame Then
        ChangeFontSizeInTF shp.TextFrame, delta
    End If

    On Error GoTo 0

End Sub

Private Sub ChangeFontSizeInTF(tf As TextFrame, delta As Single)

    On Error Resume Next
    If Not tf.HasText Then Exit Sub

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
            .SpaceWithin = 1
        End With
    Next i

    On Error GoTo 0
End Sub


' ===== FONT PRESET ===========================================================

Public Sub RunPresetFontArial()
    SavePref "LastFont", "Arial"
    ApplyFontToPresentation "Arial"
End Sub

Public Sub RunPresetFontEY()
    SavePref "LastFont", "EYInterstate Light"
    ApplyFontToPresentation "EYInterstate Light"
End Sub

Public Sub RunPresetFontTimes()
    SavePref "LastFont", "Times New Roman"
    ApplyFontToPresentation "Times New Roman"
End Sub

Public Sub RunPresetFontCalibri()
    SavePref "LastFont", "Calibri"
    ApplyFontToPresentation "Calibri"
End Sub

Public Sub RunPresetFontRepeat()
    ApplyFontToPresentation GetPref("LastFont", "Arial")
End Sub

Private Sub ApplyFontToPresentation(f As String)

    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            ApplyFontToShape shp, f
        Next shp

        ' Notes page
        If sld.HasNotesPage Then
            Dim nshp As Shape
            For Each nshp In sld.NotesPage.Shapes
                ApplyFontToShape nshp, f
            Next nshp
        End If
    Next sld

    ' Masters & layouts
    Dim dsgn As Design
    For Each dsgn In ActivePresentation.Designs
        For Each shp In dsgn.SlideMaster.Shapes
            ApplyFontToShape shp, f
        Next shp
        Dim lay As CustomLayout
        For Each lay In dsgn.SlideMaster.CustomLayouts
            For Each shp In lay.Shapes
                ApplyFontToShape shp, f
            Next shp
        Next lay
    Next dsgn

    MsgBox "Font applied: " & f, vbInformation, "Font"

End Sub

Private Sub ApplyFontToShape(shp As Shape, f As String)

    On Error Resume Next

    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            ApplyFontToShape subShp, f
        Next subShp
        Exit Sub
    End If

    If shp.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                ApplyFontToTF shp.Table.Cell(r, c).Shape.TextFrame, f
            Next c
        Next r
        Exit Sub
    End If

    If shp.HasSmartArt Then
        Dim nd As SmartArtNode
        For Each nd In shp.SmartArt.AllNodes
            If Not nd.TextFrame2 Is Nothing Then
                nd.TextFrame2.TextRange.Font.Name = f
            End If
        Next nd
        Exit Sub
    End If

    If shp.HasChart Then
        If shp.Chart.HasTitle Then
            shp.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Name = f
        End If
        Exit Sub
    End If

    If shp.HasTextFrame Then
        ApplyFontToTF shp.TextFrame, f
    End If

    On Error GoTo 0

End Sub

Private Sub ApplyFontToTF(tf As TextFrame, f As String)
    On Error Resume Next
    If tf.HasText Then
        tf.TextRange.Font.Name = f
    End If
    On Error GoTo 0
End Sub


' ===== NUMBER FORMATTING — WORKS ON ANY SELECTED TEXT ========================

Public Sub SelFormatNumNoDecimal()
    SavePref "LastNumFmt", "#,##0"
    SavePref "LastNumPrefix", ""
    FormatSelectedNumbers "#,##0", ""
End Sub

Public Sub SelFormatNumDecimal()
    SavePref "LastNumFmt", "#,##0.00"
    SavePref "LastNumPrefix", ""
    FormatSelectedNumbers "#,##0.00", ""
End Sub

Public Sub SelFormatNumDollar()
    SavePref "LastNumFmt", "#,##0.00"
    SavePref "LastNumPrefix", "$"
    FormatSelectedNumbers "#,##0.00", "$"
End Sub

Public Sub SelFormatNumRepeat()
    FormatSelectedNumbers _
        GetPref("LastNumFmt", "#,##0.00"), _
        GetPref("LastNumPrefix", "")
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

    MsgBox "Please select text or a shape containing numbers.", _
           vbExclamation, "Number Format"

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
                FormatNumInTextRange _
                    shp.Table.Cell(r, c).Shape.TextFrame.TextRange, fmt, prefix
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


' ===== DATE FORMATTING — WORKS ON ANY SELECTED TEXT ==========================

Public Sub SelFormatDateShort()
    SavePref "LastDateFmt", "DD-MMM-YY"
    FormatSelectedDates "DD-MMM-YY"
End Sub

Public Sub SelFormatDateLong()
    SavePref "LastDateFmt", "DD-MMMM-YYYY"
    FormatSelectedDates "DD-MMMM-YYYY"
End Sub

Public Sub SelFormatDateRepeat()
    FormatSelectedDates GetPref("LastDateFmt", "DD-MMM-YY")
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

    MsgBox "Please select text or a shape containing dates.", _
           vbExclamation, "Date Format"

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
                FormatDateInTextRange _
                    shp.Table.Cell(r, c).Shape.TextFrame.TextRange, fmt
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


' ===== HELPERS ===============================================================

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


' ===== PREFERENCE STORAGE ====================================================

Private Sub SavePref(key As String, val As String)
    SaveSetting "DeckUI", "Preferences", key, val
End Sub

Private Function GetPref(key As String, defaultVal As String) As String
    GetPref = GetSetting("DeckUI", "Preferences", key, defaultVal)
End Function
