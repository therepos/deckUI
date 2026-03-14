Attribute VB_Name = "SubPreset"
Option Explicit

' =============================================================================
' Font Presets — PowerPoint Edition
' =============================================================================
' Differences from Word version:
'   - No StoryRanges; walks Slides > Shapes > TextFrames instead
'   - Also covers slide masters, layouts, and notes pages
'   - Uses Presentation.CustomDocumentProperties (same COM interface)
'   - Status bar progress so users see it working
' =============================================================================

Public Sub RunPresetFontArial()
    ApplyPreset "Arial"
End Sub

Public Sub RunPresetFontEY()
    ApplyPreset "EYInterstate Light"
End Sub

Public Sub RunPresetFontTimes()
    ApplyPreset "Times New Roman"
End Sub

Public Sub RunPresetFontCalibri()
    ApplyPreset "Calibri"
End Sub


' ===========================================================================
' INTERNAL
' ===========================================================================

Private Sub ApplyPreset(f As String)

    SetPresetFont f
    ApplyFontPreset
    Application.StatusBar = ""
    MsgBox "Font applied: " & f, vbInformation, "Font"

End Sub


Private Sub ApplyFontPreset()

    Dim f As String
    Dim sld As Slide
    Dim shp As Shape
    Dim slideNum As Long

    On Error Resume Next
    f = ActivePresentation.CustomDocumentProperties("PresetFont").Value
    If f = "" Then f = "Arial"
    On Error GoTo 0

    ' --- Slides ---
    For Each sld In ActivePresentation.Slides
        slideNum = slideNum + 1
        Application.StatusBar = "Applying font — slide " & slideNum & " of " & ActivePresentation.Slides.Count & "..."

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

    ' --- Slide masters & layouts ---
    Application.StatusBar = "Applying font to masters and layouts..."
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

End Sub


Private Sub ApplyFontToShape(shp As Shape, f As String)

    On Error Resume Next

    ' Grouped shapes — recurse
    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            ApplyFontToShape subShp, f
        Next subShp
        Exit Sub
    End If

    ' Tables — each cell
    If shp.HasTable Then
        Dim row As Long, col As Long
        For row = 1 To shp.Table.Rows.Count
            For col = 1 To shp.Table.Columns.Count
                ApplyFontToTextFrame shp.Table.Cell(row, col).Shape.TextFrame, f
            Next col
        Next row
        Exit Sub
    End If

    ' SmartArt nodes
    If shp.HasSmartArt Then
        Dim nd As SmartArtNode
        For Each nd In shp.SmartArt.AllNodes
            If Not nd.TextFrame2 Is Nothing Then
                nd.TextFrame2.TextRange.Font.Name = f
            End If
        Next nd
        Exit Sub
    End If

    ' Charts — title
    If shp.HasChart Then
        If shp.Chart.HasTitle Then
            shp.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Name = f
        End If
        Exit Sub
    End If

    ' Regular text frame
    If shp.HasTextFrame Then
        ApplyFontToTextFrame shp.TextFrame, f
    End If

    On Error GoTo 0

End Sub


Private Sub ApplyFontToTextFrame(tf As TextFrame, f As String)
    On Error Resume Next
    If tf.HasText Then
        tf.TextRange.Font.Name = f
    End If
    On Error GoTo 0
End Sub


Private Sub SetPresetFont(f As String)

    On Error Resume Next
    If ActivePresentation.CustomDocumentProperties("PresetFont").Name = "" Then
        ActivePresentation.CustomDocumentProperties.Add _
            Name:="PresetFont", _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            Value:=f
    Else
        ActivePresentation.CustomDocumentProperties("PresetFont").Value = f
    End If
    On Error GoTo 0

End Sub
