Option Explicit

' =============================================================================
' MODULE: Subs
' Purpose: General formatting operations for PowerPoint presentations.
'          Works across ALL text in every slide — regular shapes, tables,
'          grouped shapes, and placeholders.
'
' Uses TextFrame / TextRange (not TextFrame2) for reliability.
' A shared recursive helper (ProcessShapeForEachRun / ProcessShapeFont)
' ensures tables and groups are never skipped.
' =============================================================================


' ===== FONT SIZE =============================================================

Sub DeckFontSizeDecrease()
    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        Dim shp As Shape
        For Each shp In sld.Shapes
            AdjustShapeFontSize shp, -1
        Next shp
    Next sld
End Sub

Sub DeckFontSizeIncrease()
    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        Dim shp As Shape
        For Each shp In sld.Shapes
            AdjustShapeFontSize shp, 1
        Next shp
    Next sld
End Sub

' --- Recursive font-size adjuster ---
Private Sub AdjustShapeFontSize(shp As Shape, delta As Long)
    Dim tr As TextRange
    Dim run As TextRange
    Dim i As Long

    On Error Resume Next

    ' --- Group: recurse into each child ---
    If shp.Type = msoGroup Then
        Dim child As Shape
        For Each child In shp.GroupItems
            AdjustShapeFontSize child, delta
        Next child
        Exit Sub
    End If

    ' --- Table: iterate every cell ---
    If shp.HasTable Then
        Dim tbl As Table
        Dim r As Long, c As Long
        Set tbl = shp.Table
        For r = 1 To tbl.Rows.Count
            For c = 1 To tbl.Columns.Count
                Set tr = tbl.Cell(r, c).Shape.TextFrame.TextRange
                For i = 1 To tr.Runs.Count
                    Set run = tr.Runs(i)
                    If run.Font.Size + delta >= 1 Then
                        run.Font.Size = run.Font.Size + delta
                    End If
                Next i
            Next c
        Next r
        Exit Sub
    End If

    ' --- Regular text frame ---
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            Set tr = shp.TextFrame.TextRange
            For i = 1 To tr.Runs.Count
                Set run = tr.Runs(i)
                If run.Font.Size + delta >= 1 Then
                    run.Font.Size = run.Font.Size + delta
                End If
            Next i
        End If
    End If

    On Error GoTo 0
End Sub


' ===== SPACING ===============================================================

Sub DeckSpacingSingle()
    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        Dim shp As Shape
        For Each shp In sld.Shapes
            ApplyShapeSingleSpacing shp
        Next shp
    Next sld
End Sub

Private Sub ApplyShapeSingleSpacing(shp As Shape)
    Dim tr As TextRange

    On Error Resume Next

    ' --- Group ---
    If shp.Type = msoGroup Then
        Dim child As Shape
        For Each child In shp.GroupItems
            ApplyShapeSingleSpacing child
        Next child
        Exit Sub
    End If

    ' --- Table ---
    If shp.HasTable Then
        Dim tbl As Table
        Dim r As Long, c As Long
        Set tbl = shp.Table
        For r = 1 To tbl.Rows.Count
            For c = 1 To tbl.Columns.Count
                Err.Clear
                Set tr = tbl.Cell(r, c).Shape.TextFrame.TextRange
                With tr.ParagraphFormat
                    .LineRuleBefore = msoFalse
                    .SpaceBefore = 0
                    .LineRuleAfter = msoFalse
                    .SpaceAfter = 0
                    .LineRuleWithin = msoFalse
                    .SpaceWithin = 0
                    .LineRuleWithin = msoTrue
                    .SpaceWithin = 1
                End With
            Next c
        Next r
        Exit Sub
    End If

    ' --- Regular text frame ---
    If shp.HasTextFrame Then
        Err.Clear
        Set tr = shp.TextFrame.TextRange
        With tr.ParagraphFormat
            .LineRuleBefore = msoFalse
            .SpaceBefore = 0
            .LineRuleAfter = msoFalse
            .SpaceAfter = 0
            .LineRuleWithin = msoFalse
            .SpaceWithin = 0
            .LineRuleWithin = msoTrue
            .SpaceWithin = 1
        End With
    End If

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
    For Each sld In ActivePresentation.Slides
        Dim shp As Shape
        For Each shp In sld.Shapes
            ApplyShapeFont shp, f
        Next shp
    Next sld
    MsgBox "Font applied: " & f, vbInformation, "Font"
End Sub

' --- Recursive font-name applier ---
Private Sub ApplyShapeFont(shp As Shape, f As String)
    On Error Resume Next

    ' --- Group ---
    If shp.Type = msoGroup Then
        Dim child As Shape
        For Each child In shp.GroupItems
            ApplyShapeFont child, f
        Next child
        Exit Sub
    End If

    ' --- Table ---
    If shp.HasTable Then
        Dim tbl As Table
        Dim r As Long, c As Long
        Set tbl = shp.Table
        For r = 1 To tbl.Rows.Count
            For c = 1 To tbl.Columns.Count
                tbl.Cell(r, c).Shape.TextFrame.TextRange.Font.Name = f
            Next c
        Next r
        Exit Sub
    End If

    ' --- Regular text frame ---
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            shp.TextFrame.TextRange.Font.Name = f
        End If
    End If

    On Error GoTo 0
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

Private Sub SavePref(key As String, val As String)
    SaveSetting "DeckUI", "Preferences", key, val
End Sub

Private Function GetPref(key As String, defaultVal As String) As String
    GetPref = GetSetting("DeckUI", "Preferences", key, defaultVal)
End Function


