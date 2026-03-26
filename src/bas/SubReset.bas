Attribute VB_Name = "SubReset"
Option Explicit

' 1 cm = 28.3464567 points (PPT has no built-in CentimetersToPoints)
Private Function CmToPt(cm As Double) As Double
    CmToPt = cm * 28.3464567
End Function

' =============================================================================
' MODULE: SubReset
' Purpose: Reset operations for PowerPoint presentations.
' Each Public Sub is called directly from a ribbon menu button.
' =============================================================================

' -----------------------------------------------------------------------------
' DECK-WIDE (all slides)
' -----------------------------------------------------------------------------

Public Sub ResetAll()
    ResetFormat ActivePresentation.slides
    ResetObject ActivePresentation.slides
    ResetTables ActivePresentation.slides
    ResetHyperlinks ActivePresentation.slides
    MsgBox "Reset complete:" & vbCrLf & vbCrLf & _
           "Formatting, Objects, Tables, Hyperlinks (All)", _
           vbInformation, "Reset"
End Sub

Public Sub RunResetFormat()
    ResetFormat ActivePresentation.slides
    MsgBox "Reset complete: Formatting (All Slides)", vbInformation, "Reset"
End Sub

Public Sub RunResetObject()
    ResetObject ActivePresentation.slides
    MsgBox "Reset complete: Objects (All Slides)", vbInformation, "Reset"
End Sub

Public Sub RunResetTables()
    ResetTables ActivePresentation.slides
    MsgBox "Reset complete: Tables (All Slides)", vbInformation, "Reset"
End Sub

Public Sub RunResetHyperlinks()
    ResetHyperlinks ActivePresentation.slides
    MsgBox "Reset complete: Hyperlinks (All Slides)", vbInformation, "Reset"
End Sub

' -----------------------------------------------------------------------------
' SELECTION-SCOPED (selected slides only)
' -----------------------------------------------------------------------------

Public Sub SelResetFormat()
    Dim sel As SlideRange
    If Not GetSelectedSlides(sel) Then Exit Sub
    ResetFormat sel
    MsgBox "Reset complete: Formatting (" & sel.Count & " slide(s))", vbInformation, "Reset"
End Sub

Public Sub SelResetObject()
    Dim sel As SlideRange
    If Not GetSelectedSlides(sel) Then Exit Sub
    ResetObject sel
    MsgBox "Reset complete: Objects (" & sel.Count & " slide(s))", vbInformation, "Reset"
End Sub

Public Sub SelResetTables()
    Dim sel As SlideRange
    If Not GetSelectedSlides(sel) Then Exit Sub
    ResetTables sel
    MsgBox "Reset complete: Tables (" & sel.Count & " slide(s))", vbInformation, "Reset"
End Sub

Public Sub SelResetHyperlinks()
    Dim sel As SlideRange
    If Not GetSelectedSlides(sel) Then Exit Sub
    ResetHyperlinks sel
    MsgBox "Reset complete: Hyperlinks (" & sel.Count & " slide(s))", vbInformation, "Reset"
End Sub

' -----------------------------------------------------------------------------
' HELPER: Get selected slides (returns False if none selected)
' -----------------------------------------------------------------------------

Private Function GetSelectedSlides(ByRef sel As SlideRange) As Boolean
    On Error GoTo NoSelection
    Set sel = ActiveWindow.Selection.SlideRange
    If sel.Count = 0 Then GoTo NoSelection
    GetSelectedSlides = True
    Exit Function
NoSelection:
    MsgBox "Please select one or more slides first.", vbExclamation, "Reset"
    GetSelectedSlides = False
End Function

' ===========================================================================
'  RESET SUBS (Private) - accept any slide collection
' ===========================================================================

Private Sub ResetFormat(slides As Object)
    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame2
    Dim r As Long

    On Error Resume Next
    For Each sld In slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                Set tf = shp.TextFrame2
                If tf.HasText Then
                    For r = 1 To tf.TextRange.Runs.Count
                        With tf.TextRange.Runs(r).Font
                            .Bold = msoFalse
                            .Italic = msoFalse
                            .UnderlineStyle = msoNoUnderline
                            .Strike = msoNoStrike
                            .Subscript = msoFalse
                            .Superscript = msoFalse
                        End With
                    Next r
                End If
            End If
        Next shp
    Next sld
    On Error GoTo 0
End Sub


Private Sub ResetObject(slides As Object)
    Dim sld As Slide
    Dim shp As Shape

    On Error Resume Next
    For Each sld In slides
        For Each shp In sld.Shapes
            If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
                shp.ScaleHeight 1, msoTrue
                shp.ScaleWidth 1, msoTrue
            End If
        Next shp
    Next sld
    On Error GoTo 0
End Sub


Private Sub ResetTables(slides As Object)
    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim r As Long, c As Long
    Dim cel As Cell

    On Error Resume Next
    For Each sld In slides
        For Each shp In sld.Shapes
            If shp.HasTable Then
                Set tbl = shp.Table
                For r = 1 To tbl.Rows.Count
                    For c = 1 To tbl.Columns.Count
                        Set cel = tbl.Cell(r, c)
                        With cel.Shape.TextFrame2
                            .MarginTop = CmToPt(0.05)
                            .MarginBottom = CmToPt(0.05)
                            .MarginLeft = CmToPt(0.19)
                            .MarginRight = CmToPt(0.19)
                        End With
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


Private Sub ResetHyperlinks(slides As Object)
    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame
    Dim tr As TextRange
    Dim i As Long

    On Error Resume Next
    For Each sld In slides
        For Each shp In sld.Shapes
            If shp.ActionSettings(ppMouseClick).Hyperlink.Address <> "" Then
                shp.ActionSettings(ppMouseClick).Hyperlink.Address = ""
                shp.ActionSettings(ppMouseClick).Hyperlink.SubAddress = ""
            End If

            If shp.HasTextFrame Then
                Set tf = shp.TextFrame
                If tf.HasText Then
                    Set tr = tf.TextRange
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

