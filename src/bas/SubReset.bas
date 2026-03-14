Option Explicit

' =============================================================================
' MODULE: SubReset — PowerPoint Edition
' Purpose: Presentation-level reset operations.
'          Ported from wordUI SubReset.bas with PPT equivalents.
'
' Contents:
'   - ResetAll
'   - RunResetFormat
'   - RunResetTables
'   - RunResetHyperlinks
' =============================================================================


Public Sub ResetAll()
    ResetFormat
    ResetTables
    ResetHyperlinks
    MsgBox "Reset complete:" & vbCrLf & vbCrLf & _
           "Formatting, Tables, Hyperlinks (All)", _
           vbInformation, "Reset"
End Sub

Public Sub RunResetFormat()
    ResetFormat
    MsgBox "Reset complete: Formatting", vbInformation, "Reset"
End Sub

Public Sub RunResetTables()
    ResetTables
    MsgBox "Reset complete: Tables", vbInformation, "Reset"
End Sub

Public Sub RunResetHyperlinks()
    ResetHyperlinks
    MsgBox "Reset complete: Hyperlinks", vbInformation, "Reset"
End Sub


' ===========================================================================
'  RESET SUBS (Private)
' ===========================================================================

Private Sub ResetFormat()
' Resets all text in the presentation to Calibri 11pt, no bold/italic/underline,
' black text, left-aligned, single-spaced, no bullets.

    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            ResetFormatInShape shp
        Next shp
    Next sld

    ' Masters & layouts
    Dim dsgn As Design
    For Each dsgn In ActivePresentation.Designs
        For Each shp In dsgn.SlideMaster.Shapes
            ResetFormatInShape shp
        Next shp
        Dim lay As CustomLayout
        For Each lay In dsgn.SlideMaster.CustomLayouts
            For Each shp In lay.Shapes
                ResetFormatInShape shp
            Next shp
        Next lay
    Next dsgn

End Sub

Private Sub ResetFormatInShape(shp As Shape)

    On Error Resume Next

    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            ResetFormatInShape subShp
        Next subShp
        Exit Sub
    End If

    If shp.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                ResetFormatInTF shp.Table.Cell(r, c).Shape.TextFrame
            Next c
        Next r
        Exit Sub
    End If

    If shp.HasTextFrame Then
        ResetFormatInTF shp.TextFrame
    End If

    On Error GoTo 0

End Sub

Private Sub ResetFormatInTF(tf As TextFrame)

    On Error Resume Next
    If Not tf.HasText Then Exit Sub

    Dim tr As TextRange
    Set tr = tf.TextRange

    With tr.Font
        .Name = "Calibri"
        .Size = 11
        .Bold = msoFalse
        .Italic = msoFalse
        .Underline = msoFalse
        .Color.RGB = RGB(0, 0, 0)
        .Shadow = msoFalse
    End With

    Dim p As Long
    For p = 1 To tr.Paragraphs.Count
        With tr.Paragraphs(p).ParagraphFormat
            .Alignment = ppAlignLeft
            .SpaceBefore = 0
            .SpaceAfter = 0
            .SpaceWithin = 1
            .Bullet.Type = ppBulletNone
        End With
        tr.Paragraphs(p).IndentLevel = 1
    Next p

    On Error GoTo 0

End Sub


Private Sub ResetTables()
' Resets all tables in the presentation to plain formatting:
' no fill, thin borders, default margins, Calibri 11pt.

    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim cel As Cell
    Dim r As Long, c As Long, p As Long
    Dim cellTR As TextRange

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            On Error Resume Next
            If shp.HasTable Then
                Set tbl = shp.Table

                For r = 1 To tbl.Rows.Count
                    For c = 1 To tbl.Columns.Count
                        Set cel = tbl.Cell(r, c)

                        ' Clear fill
                        cel.Shape.Fill.Visible = msoFalse

                        ' Reset borders to thin black
                        With cel.Borders(ppBorderTop)
                            .ForeColor.RGB = RGB(0, 0, 0)
                            .Weight = 0.25
                            .DashStyle = msoLineSolid
                            .Visible = msoTrue
                        End With
                        With cel.Borders(ppBorderBottom)
                            .ForeColor.RGB = RGB(0, 0, 0)
                            .Weight = 0.25
                            .DashStyle = msoLineSolid
                            .Visible = msoTrue
                        End With
                        With cel.Borders(ppBorderLeft)
                            .ForeColor.RGB = RGB(0, 0, 0)
                            .Weight = 0.25
                            .DashStyle = msoLineSolid
                            .Visible = msoTrue
                        End With
                        With cel.Borders(ppBorderRight)
                            .ForeColor.RGB = RGB(0, 0, 0)
                            .Weight = 0.25
                            .DashStyle = msoLineSolid
                            .Visible = msoTrue
                        End With

                        ' Reset margins
                        With cel.Shape.TextFrame
                            .MarginTop = CmToPt(0.13)
                            .MarginBottom = CmToPt(0.13)
                            .MarginLeft = CmToPt(0.25)
                            .MarginRight = CmToPt(0.25)
                            .WordWrap = msoTrue
                            .AutoSize = ppAutoSizeNone
                        End With

                        ' Reset text formatting
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
            End If
            On Error GoTo 0
        Next shp
    Next sld

End Sub


Private Sub ResetHyperlinks()
' Removes all hyperlinks from shapes across all slides.
' PowerPoint hyperlinks are stored as ActionSettings on text runs and shapes.

    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            RemoveHyperlinksInShape shp
        Next shp
    Next sld

End Sub

Private Sub RemoveHyperlinksInShape(shp As Shape)

    On Error Resume Next

    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            RemoveHyperlinksInShape subShp
        Next subShp
        Exit Sub
    End If

    ' Shape-level hyperlink
    If shp.ActionSettings(ppMouseClick).Hyperlink.Address <> "" Then
        shp.ActionSettings(ppMouseClick).Hyperlink.Address = ""
        shp.ActionSettings(ppMouseClick).Hyperlink.SubAddress = ""
        shp.ActionSettings(ppMouseClick).Action = ppActionNone
    End If

    ' Text-level hyperlinks
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            Dim hl As Hyperlink
            Dim hlCount As Long
            hlCount = shp.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink.Address <> ""

            ' Walk runs in reverse to clear hyperlinks
            Dim i As Long
            For i = shp.TextFrame.TextRange.Runs.Count To 1 Step -1
                With shp.TextFrame.TextRange.Runs(i).ActionSettings(ppMouseClick)
                    If .Hyperlink.Address <> "" Or .Hyperlink.SubAddress <> "" Then
                        .Hyperlink.Address = ""
                        .Hyperlink.SubAddress = ""
                        .Action = ppActionNone
                    End If
                End With
            Next i
        End If
    End If

    ' Table cell hyperlinks
    If shp.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                Dim cellShp As Shape
                Set cellShp = shp.Table.Cell(r, c).Shape
                If cellShp.HasTextFrame Then
                    If cellShp.TextFrame.HasText Then
                        Dim j As Long
                        For j = cellShp.TextFrame.TextRange.Runs.Count To 1 Step -1
                            With cellShp.TextFrame.TextRange.Runs(j).ActionSettings(ppMouseClick)
                                If .Hyperlink.Address <> "" Or .Hyperlink.SubAddress <> "" Then
                                    .Hyperlink.Address = ""
                                    .Hyperlink.SubAddress = ""
                                    .Action = ppActionNone
                                End If
                            End With
                        Next j
                    End If
                End If
            Next c
        Next r
    End If

    On Error GoTo 0

End Sub


' =============================================================================
' HELPERS
' =============================================================================

Private Function CmToPt(ByVal cm As Double) As Double
    CmToPt = cm * 28.3464567
End Function
