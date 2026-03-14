Option Explicit

' =============================================================================
' MODULE: Subs
' Purpose: General formatting operations for PowerPoint presentations
'          Works across all slides, shapes, and text frames.
'
' Contents:
'   - DeckFontSizeDecrease / DeckFontSizeIncrease
'   - DeckSpacingSingle
'   - RunPresetFontArial / RunPresetFontEY / RunPresetFontTimes /
'     RunPresetFontCalibri / RunPresetFontRepeat
'   - SelFormatNumDecimal / SelFormatNumNoDecimal / SelFormatNumDollar /
'     SelFormatNumRepeat
'   - SelFormatDateShort / SelFormatDateLong / SelFormatDateRepeat
' =============================================================================


' ===== FONT SIZE =============================================================

Sub DeckFontSizeDecrease()

    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame2
    Dim rng As TextRange2
    Dim para As TextRange2
    Dim r As Long

    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                Set tf = shp.TextFrame2
                If tf.HasText Then
                    For r = 1 To tf.TextRange.Runs.Count
                        Set rng = tf.TextRange.Runs(r)
                        If rng.Font.Size > 1 Then
                            rng.Font.Size = rng.Font.Size - 1
                        End If
                    Next r
                End If
            End If
        Next shp
    Next sld
    On Error GoTo 0

End Sub

Sub DeckFontSizeIncrease()

    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame2
    Dim rng As TextRange2
    Dim r As Long

    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                Set tf = shp.TextFrame2
                If tf.HasText Then
                    For r = 1 To tf.TextRange.Runs.Count
                        Set rng = tf.TextRange.Runs(r)
                        rng.Font.Size = rng.Font.Size + 1
                    Next r
                End If
            End If
        Next shp
    Next sld
    On Error GoTo 0

End Sub


' ===== SPACING ===============================================================

Sub DeckSpacingSingle()

    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame2
    Dim para As TextRange2
    Dim p As Long

    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                Set tf = shp.TextFrame2
                If tf.HasText Then
                    For p = 1 To tf.TextRange.Paragraphs.Count
                        Set para = tf.TextRange.Paragraphs(p)
                        With para.ParagraphFormat
                            .SpaceBefore = 0
                            .SpaceAfter = 0
                            .SpaceWithin = 1
                        End With
                    Next p
                End If
            End If
        Next shp
    Next sld
    On Error GoTo 0

End Sub


' ===== FONT PRESET ===========================================================

Public Sub RunPresetFontArial()
    SavePref "LastFont", "Arial"
    ApplyFontToDeck "Arial"
End Sub

Public Sub RunPresetFontEY()
    SavePref "LastFont", "EYInterstate Light"
    ApplyFontToDeck "EYInterstate Light"
End Sub

Public Sub RunPresetFontTimes()
    SavePref "LastFont", "Times New Roman"
    ApplyFontToDeck "Times New Roman"
End Sub

Public Sub RunPresetFontCalibri()
    SavePref "LastFont", "Calibri"
    ApplyFontToDeck "Calibri"
End Sub

Public Sub RunPresetFontRepeat()
    ApplyFontToDeck GetPref("LastFont", "Arial")
End Sub

Private Sub ApplyFontToDeck(f As String)

    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame2

    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                Set tf = shp.TextFrame2
                If tf.HasText Then
                    tf.TextRange.Font.Name = f
                End If
            End If
        Next shp
    Next sld
    On Error GoTo 0

    MsgBox "Font applied: " & f, vbInformation, "Font"

End Sub


' ===== NUMBER FORMATTING — WORKS ON SELECTED TEXT ============================

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
    Dim tr As TextRange
    Dim cellText As String
    Dim val As Double

    Set sel = ActiveWindow.Selection

    ' --- Case 1: Text selected in a shape or table cell ---
    If sel.Type = ppSelectionText Then
        Set tr = sel.TextRange
        cellText = CleanNumericText(tr.Text)
        If IsNumeric(cellText) And Len(cellText) > 0 Then
            val = CDbl(cellText)
            tr.Text = FormatValue(val, fmt, prefix)
        End If
        Exit Sub
    End If

    ' --- Case 2: A table shape is selected — format all cells ---
    If sel.Type = ppSelectionShapes Then
        Dim shp As Shape
        Dim tbl As Table
        Dim r As Long, c As Long

        For Each shp In sel.ShapeRange
            If shp.HasTable Then
                Set tbl = shp.Table
                For r = 1 To tbl.Rows.Count
                    For c = 1 To tbl.Columns.Count
                        Set tr = tbl.Cell(r, c).Shape.TextFrame.TextRange
                        cellText = CleanNumericText(tr.Text)
                        If IsNumeric(cellText) And Len(cellText) > 0 Then
                            val = CDbl(cellText)
                            tr.Text = FormatValue(val, fmt, prefix)
                            tr.ParagraphFormat.Alignment = ppAlignRight
                        End If
                    Next c
                Next r
            End If
        Next shp
    End If

End Sub


' ===== DATE FORMATTING — WORKS ON SELECTED TEXT ==============================

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
    Dim tr As TextRange
    Dim cellText As String

    Set sel = ActiveWindow.Selection

    If sel.Type = ppSelectionText Then
        Set tr = sel.TextRange
        cellText = Trim$(tr.Text)
        If IsDate(cellText) Then
            tr.Text = Format(CDate(cellText), fmt)
        End If
        Exit Sub
    End If

    ' Table shape selected — format all cells
    If sel.Type = ppSelectionShapes Then
        Dim shp As Shape
        Dim tbl As Table
        Dim r As Long, c As Long

        For Each shp In sel.ShapeRange
            If shp.HasTable Then
                Set tbl = shp.Table
                For r = 1 To tbl.Rows.Count
                    For c = 1 To tbl.Columns.Count
                        Set tr = tbl.Cell(r, c).Shape.TextFrame.TextRange
                        cellText = Trim$(tr.Text)
                        If IsDate(cellText) Then
                            tr.Text = Format(CDate(cellText), fmt)
                        End If
                    Next c
                Next r
            End If
        Next shp
    End If

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
    t = Replace(t, Chr(13), "")
    If InStr(t, "(") > 0 And InStr(t, ")") > 0 Then
        t = Replace(t, "(", "-")
        t = Replace(t, ")", "")
    End If
    CleanNumericText = Trim$(t)
End Function


' ===== PREFERENCE STORAGE ====================================================
' Uses Windows registry (HKEY_CURRENT_USER — no admin rights needed)
' so preferences persist across all presentations for the user.

Private Sub SavePref(key As String, val As String)
    SaveSetting "DeckUI", "Preferences", key, val
End Sub

Private Function GetPref(key As String, defaultVal As String) As String
    GetPref = GetSetting("DeckUI", "Preferences", key, defaultVal)
End Function
