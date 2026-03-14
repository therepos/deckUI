Option Explicit

' =============================================================================
' MODULE: SubReset
' Purpose: Reset operations for PowerPoint presentations.
' Each Public Sub is called directly from a ribbon menu button.
' =============================================================================

Public Sub ResetAll()
    ResetFormat
    ResetObject
    ResetTables
    ResetHyperlinks
    MsgBox "Reset complete:" & vbCrLf & vbCrLf & _
           "Formatting, Objects, Tables, Hyperlinks (All)", _
           vbInformation, "Reset"
End Sub

Public Sub RunResetFormat()
    ResetFormat
    MsgBox "Reset complete: Formatting", vbInformation, "Reset"
End Sub

Public Sub RunResetObject()
    ResetObject
    MsgBox "Reset complete: Objects", vbInformation, "Reset"
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
' Clears all direct character formatting from every text frame
' in every shape on every slide, reverting to the slide-master style.

    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame2
    Dim r As Long

    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                Set tf = shp.TextFrame2
                If tf.HasText Then
                    ' Reset each run to inherit from theme/master
                    For r = 1 To tf.TextRange.Runs.Count
                        With tf.TextRange.Runs(r).Font
                            .Bold = msoFalse
                            .Italic = msoFalse
                            .Underline = msoFalse
                            .Strikethrough = msoNoStrike
                            .Subscript = msoFalse
                            .Superscript = msoFalse
                            .Shadow = msoFalse
                        End With
                    Next r
                End If
            End If
        Next shp
    Next sld
    On Error GoTo 0

End Sub


Private Sub ResetObject()
' Resets all pictures / inline shapes to their original size.

    Dim sld As Slide
    Dim shp As Shape

    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
                shp.ScaleHeight 1, msoTrue
                shp.ScaleWidth 1, msoTrue
            End If
        Next shp
    Next sld
    On Error GoTo 0

End Sub


Private Sub ResetTables()
' Resets table cell padding and applies thin borders to all tables.

    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim r As Long, c As Long
    Dim cel As Cell

    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTable Then
                Set tbl = shp.Table
                For r = 1 To tbl.Rows.Count
                    For c = 1 To tbl.Columns.Count
                        Set cel = tbl.Cell(r, c)
                        With cel.Shape.TextFrame2
                            .MarginTop = Application.CentimetersToPoints(0.05)
                            .MarginBottom = Application.CentimetersToPoints(0.05)
                            .MarginLeft = Application.CentimetersToPoints(0.19)
                            .MarginRight = Application.CentimetersToPoints(0.19)
                        End With
                        ' Borders
                        Dim bdr As Long
                        For bdr = ppBorderTop To ppBorderDiagonalUp
                            With cel.Borders(bdr)
                                If bdr <= ppBorderRight Then
                                    .ForeColor.RGB = RGB(0, 0, 0)
                                    .Weight = 0.5
                                    .Visible = msoTrue
                                    .DashStyle = msoLineSolid
                                Else
                                    .Visible = msoFalse
                                End If
                            End With
                        Next bdr
                    Next c
                Next r
            End If
        Next shp
    Next sld
    On Error GoTo 0

End Sub


Private Sub ResetHyperlinks()
' Removes all hyperlinks from the active presentation.

    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame
    Dim tr As TextRange
    Dim hl As Hyperlink
    Dim i As Long

    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        ' Shape-level hyperlinks (click actions)
        For Each shp In sld.Shapes
            If shp.ActionSettings(ppMouseClick).Hyperlink.Address <> "" Then
                shp.ActionSettings(ppMouseClick).Hyperlink.Address = ""
                shp.ActionSettings(ppMouseClick).Hyperlink.SubAddress = ""
            End If

            ' Text-level hyperlinks
            If shp.HasTextFrame Then
                Set tf = shp.TextFrame
                If tf.HasText Then
                    Set tr = tf.TextRange
                    ' Walk hyperlinks in reverse to avoid index issues
                    For i = tr.ActionSettings.Count To 1 Step -1
                        tr.ActionSettings(i).Hyperlink.Address = ""
                        tr.ActionSettings(i).Hyperlink.SubAddress = ""
                    Next i
                End If
            End If
        Next shp
    Next sld
    On Error GoTo 0

End Sub
