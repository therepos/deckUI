Attribute VB_Name = "Subs"
Sub FontArial()
    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim row As Long, col As Long
    
    ' Loop through all slides
    For Each sld In ActivePresentation.Slides
        ' Loop through all shapes in each slide
        For Each shp In sld.Shapes
            ' Check if shape has text
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Font.Name = "Arial"
                End If
            End If
            
            ' Check if shape is a table
            If shp.HasTable Then
                Set tbl = shp.Table
                For row = 1 To tbl.Rows.Count
                    For col = 1 To tbl.Columns.Count
                        tbl.Cell(row, col).Shape.TextFrame.TextRange.Font.Name = "Arial"
                    Next col
                Next row
            End If
        Next shp
    Next sld
End Sub

Sub FontEY()
    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim row As Long, col As Long
    
    ' Loop through all slides
    For Each sld In ActivePresentation.Slides
        ' Loop through all shapes in each slide
        For Each shp In sld.Shapes
            ' Check if shape has text
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Font.Name = "EYInterstate Light"
                End If
            End If
            
            ' Check if shape is a table
            If shp.HasTable Then
                Set tbl = shp.Table
                For row = 1 To tbl.Rows.Count
                    For col = 1 To tbl.Columns.Count
                        tbl.Cell(row, col).Shape.TextFrame.TextRange.Font.Name = "EYInterstate Light"
                    Next col
                Next row
            End If
        Next shp
    Next sld
End Sub

Sub FontSize12()
    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim row As Long, col As Long
    
    ' Loop through all slides
    For Each sld In ActivePresentation.Slides
        ' Loop through all shapes in each slide
        For Each shp In sld.Shapes
            ' Check if shape has text
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    With shp.TextFrame.TextRange.Font
                        .Size = 12
                    End With
                End If
            End If
            
            ' Check if shape is a table
            If shp.HasTable Then
                Set tbl = shp.Table
                For row = 1 To tbl.Rows.Count
                    For col = 1 To tbl.Columns.Count
                        With tbl.Cell(row, col).Shape.TextFrame.TextRange.Font
                            .Size = 12
                        End With
                    Next col
                Next row
            End If
        Next shp
    Next sld
End Sub

Sub FontSizeUp()
    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim r As Long, c As Long
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Font.Size = shp.TextFrame.TextRange.Font.Size + 1
                End If
            End If
            
            If shp.HasTable Then
                Set tbl = shp.Table
                For r = 1 To tbl.Rows.Count
                    For c = 1 To tbl.Columns.Count
                        tbl.Cell(r, c).Shape.TextFrame.TextRange.Font.Size = _
                            tbl.Cell(r, c).Shape.TextFrame.TextRange.Font.Size + 1
                    Next c
                Next r
            End If
        Next shp
    Next sld
End Sub

Sub FontSizeDown()
    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim r As Long, c As Long
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Font.Size = shp.TextFrame.TextRange.Font.Size - 1
                End If
            End If
            
            If shp.HasTable Then
                Set tbl = shp.Table
                For r = 1 To tbl.Rows.Count
                    For c = 1 To tbl.Columns.Count
                        tbl.Cell(r, c).Shape.TextFrame.TextRange.Font.Size = _
                            tbl.Cell(r, c).Shape.TextFrame.TextRange.Font.Size - 1
                    Next c
                Next r
            End If
        Next shp
    Next sld
End Sub

Sub SelectedTableBorders()
    Dim shp As Shape
    Dim tbl As Table
    Dim r As Long, c As Long
    
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set shp = ActiveWindow.Selection.ShapeRange(1)
        If shp.HasTable Then
            Set tbl = shp.Table
            For r = 1 To tbl.Rows.Count
                For c = 1 To tbl.Columns.Count
                    With tbl.Cell(r, c).Borders(ppBorderTop)
                        .Weight = 1
                        .ForeColor.RGB = RGB(0, 0, 0)
                    End With
                    With tbl.Cell(r, c).Borders(ppBorderBottom)
                        .Weight = 1
                        .ForeColor.RGB = RGB(0, 0, 0)
                    End With
                    With tbl.Cell(r, c).Borders(ppBorderLeft)
                        .Weight = 1
                        .ForeColor.RGB = RGB(0, 0, 0)
                    End With
                    With tbl.Cell(r, c).Borders(ppBorderRight)
                        .Weight = 1
                        .ForeColor.RGB = RGB(0, 0, 0)
                    End With
                Next c
            Next r
        Else
            MsgBox "Selected shape is not a table."
        End If
    Else
        MsgBox "Please select a table first."
    End If
End Sub

Sub SelectedTableShade()
    Dim shp As Shape
    Dim tbl As Table
    Dim r As Long, c As Long
    Const LIGHT_GREY As Long = &HF2F2F2   ' RGB(242,242,242)

    ' Get the table from the current selection
    On Error Resume Next
    If ActiveWindow.Selection.Type = ppSelectionShapes _
       And ActiveWindow.Selection.ShapeRange(1).HasTable Then
        Set tbl = ActiveWindow.Selection.ShapeRange(1).Table
    ElseIf ActiveWindow.Selection.Type = ppSelectionText _
       And ActiveWindow.Selection.TextRange.Parent.HasTable Then
        Set tbl = ActiveWindow.Selection.TextRange.Parent.Table
    End If
    On Error GoTo 0

    If tbl Is Nothing Then
        MsgBox "Please select a table first.", vbExclamation
        Exit Sub
    End If

    ' Row 1 is header – skip it
    For r = 2 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            With tbl.Cell(r, c).Shape.Fill
                If r Mod 2 = 0 Then
                    .Visible = msoTrue
                    .ForeColor.RGB = LIGHT_GREY
                    .Solid
                Else
                    .Visible = msoFalse   ' no color
                End If
            End With
        Next c
    Next r
End Sub

Sub TableNormalMargin()
    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim row As Long, col As Long
    
    ' Loop through all slides
    For Each sld In ActivePresentation.Slides
        ' Loop through all shapes in each slide
        For Each shp In sld.Shapes
            ' Check if shape is a table
            If shp.HasTable Then
                Set tbl = shp.Table
                ' Loop through all cells in the table
                For row = 1 To tbl.Rows.Count
                    For col = 1 To tbl.Columns.Count
                        With tbl.Cell(row, col).Shape.TextFrame
                            ' Apply Normal margins (default in PowerPoint)
                            .MarginTop = 3    ' Normal top margin
                            .MarginBottom = 3 ' Normal bottom margin
                            .MarginLeft = 3   ' Normal left margin
                            .MarginRight = 3  ' Normal right margin
                        End With
                    Next col
                Next row
            End If
        Next shp
    Next sld
End Sub
