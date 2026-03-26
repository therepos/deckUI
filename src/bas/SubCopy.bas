Attribute VB_Name = "SubCopy"
'=============================================================================
' FreezeColors - PowerPoint VBA Module
' Converts all theme-referenced colors to hard-coded RGB values.
' Run before copying slides to another deck to preserve colors.
'=============================================================================

Option Explicit

'-------------------------------------------------------------
' PUBLIC ENTRY POINTS
'-------------------------------------------------------------

' Freeze colors on ALL slides in the active presentation
Public Sub FreezeAllSlides()
    Dim sld As Slide
    Dim total As Long
    total = ActivePresentation.slides.Count
    
    If total = 0 Then
        MsgBox "No slides found.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To total
        Set sld = ActivePresentation.slides(i)
        FreezeSlide sld
        DoEvents
    Next i
    
    MsgBox "Done! Colors frozen on all " & total & " slides." & vbCrLf & _
           "You can now safely copy them to any deck.", vbInformation
End Sub

' Freeze colors on currently selected slides only
Public Sub FreezeSelectedSlides()
    On Error GoTo NoSelection
    
    Dim sel As SlideRange
    Set sel = ActiveWindow.Selection.SlideRange
    
    If sel.Count = 0 Then
        MsgBox "No slides selected.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To sel.Count
        FreezeSlide sel(i)
        DoEvents
    Next i
    
    MsgBox "Done! Colors frozen on " & sel.Count & " selected slide(s).", vbInformation
    Exit Sub
    
NoSelection:
    MsgBox "Please select one or more slides in the slide panel first.", vbExclamation
End Sub

'-------------------------------------------------------------
' CORE: Process a single slide
'-------------------------------------------------------------
Private Sub FreezeSlide(sld As Slide)
    ' Freeze slide background
    FreezeBackground sld
    
    ' Freeze all shapes
    Dim shp As Shape
    For Each shp In sld.Shapes
        FreezeShape shp
    Next shp
End Sub

'-------------------------------------------------------------
' BACKGROUND
'-------------------------------------------------------------
Private Sub FreezeBackground(sld As Slide)
    On Error Resume Next
    With sld.Background.fill
        If .Type = msoFillSolid Then
            FreezeColorFormat .ForeColor
        ElseIf .Type = msoFillGradient Then
            FreezeGradient .ForeColor, .BackColor
            FreezeGradientStops sld.Background.fill
        End If
    End With
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
' SHAPE DISPATCHER (handles groups, tables, charts, etc.)
'-------------------------------------------------------------
Private Sub FreezeShape(shp As Shape)
    On Error Resume Next
    
    ' Handle grouped shapes recursively
    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            FreezeShape subShp
        Next subShp
        Exit Sub
    End If
    
    ' Handle tables
    If shp.HasTable Then
        FreezeTable shp.Table
        Exit Sub
    End If
    
    ' Handle charts
    If shp.HasChart Then
        FreezeChart shp.Chart
        Exit Sub
    End If
    
    ' Standard shape: fill, line, text
    FreezeFill shp
    FreezeLine shp
    
    If shp.HasTextFrame Then
        FreezeTextFrame shp.TextFrame
    End If
    
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
' FILL
'-------------------------------------------------------------
Private Sub FreezeFill(shp As Shape)
    On Error Resume Next
    With shp.fill
        Select Case .Type
            Case msoFillSolid
                FreezeColorFormat .ForeColor
            Case msoFillGradient
                FreezeGradient .ForeColor, .BackColor
                FreezeGradientStops shp.fill
            Case msoFillPatterned
                FreezeColorFormat .ForeColor
                FreezeColorFormat .BackColor
        End Select
    End With
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
' GRADIENT STOPS
'-------------------------------------------------------------
Private Sub FreezeGradientStops(fill As FillFormat)
    On Error Resume Next
    Dim gs As GradientStop
    Dim i As Long
    For i = 1 To fill.GradientStops.Count
        Set gs = fill.GradientStops(i)
        FreezeColorFormat gs.Color
    Next i
    On Error GoTo 0
End Sub

Private Sub FreezeGradient(fc As ColorFormat, bc As ColorFormat)
    On Error Resume Next
    FreezeColorFormat fc
    FreezeColorFormat bc
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
' LINE / BORDER
'-------------------------------------------------------------
Private Sub FreezeLine(shp As Shape)
    On Error Resume Next
    If shp.Line.Visible = msoTrue Then
        FreezeColorFormat shp.Line.ForeColor
        FreezeColorFormat shp.Line.BackColor
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
' TEXT
'-------------------------------------------------------------
Private Sub FreezeTextFrame(tf As TextFrame)
    On Error Resume Next
    If Not tf.HasText Then Exit Sub
    
    Dim tr As TextRange
    Set tr = tf.TextRange
    
    ' Freeze each run individually to preserve per-run formatting
    Dim i As Long
    For i = 1 To tr.Runs.Count
        FreezeColorFormat tr.Runs(i).Font.Color
    Next i
    
    ' Also freeze any paragraph-level bullet colors
    Dim para As TextRange
    For i = 1 To tr.Paragraphs.Count
        Set para = tr.Paragraphs(i)
        ' Bullet color
        If para.ParagraphFormat.Bullet.Type <> ppBulletNone Then
            FreezeColorFormat para.ParagraphFormat.Bullet.Font.Color
        End If
    Next i
    
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
' TABLE
'-------------------------------------------------------------
Private Sub FreezeTable(tbl As Table)
    On Error Resume Next
    Dim r As Long, c As Long
    
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            With tbl.Cell(r, c)
                ' Cell fill
                If .Shape.fill.Type = msoFillSolid Then
                    FreezeColorFormat .Shape.fill.ForeColor
                End If
                
                ' Cell borders
                FreezeBorder .Borders(ppBorderTop)
                FreezeBorder .Borders(ppBorderBottom)
                FreezeBorder .Borders(ppBorderLeft)
                FreezeBorder .Borders(ppBorderRight)
                
                ' Cell text
                If .Shape.HasTextFrame Then
                    FreezeTextFrame .Shape.TextFrame
                End If
            End With
        Next c
    Next r
    On Error GoTo 0
End Sub

Private Sub FreezeBorder(brd As LineFormat)
    On Error Resume Next
    If brd.Visible = msoTrue Then
        FreezeColorFormat brd.ForeColor
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
' CHART (basic series/axes colors)
'-------------------------------------------------------------
Private Sub FreezeChart(cht As Chart)
    On Error Resume Next
    Dim i As Long
    
    ' Series fills and lines
    For i = 1 To cht.SeriesCollection.Count
        With cht.SeriesCollection(i)
            If .Format.fill.Type = msoFillSolid Then
                FreezeColorFormat .Format.fill.ForeColor
            End If
            If .Format.Line.Visible = msoTrue Then
                FreezeColorFormat .Format.Line.ForeColor
            End If
        End With
    Next i
    
    ' Chart title
    If cht.HasTitle Then
        FreezeColorFormat cht.ChartTitle.Format.TextFrame2.TextRange.Font.fill.ForeColor
    End If
    
    On Error GoTo 0
End Sub

'-------------------------------------------------------------
' THE KEY FUNCTION: Convert a single ColorFormat to hard RGB
'-------------------------------------------------------------
Private Sub FreezeColorFormat(cf As ColorFormat)
    On Error Resume Next
    
    ' Only convert if it's a theme/scheme color
    If cf.Type = msoColorTypeScheme Or cf.ObjectThemeColor <> msoNotThemeColor Then
        Dim rgbVal As Long
        rgbVal = cf.RGB  ' Read the currently rendered color
        cf.RGB = rgbVal  ' Write it back as a hard-coded value
    End If
    
    On Error GoTo 0
End Sub

