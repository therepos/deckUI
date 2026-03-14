# VBA modules

_This file is generated automatically from `.bas` files in `src/bas`._

## Module `Ribbon`

### `RibbonOnLoad`

```vbnet
Public Sub RibbonOnLoad(r As IRibbonUI)
    Set Ribbon = r
End Sub
```

### `RunByName`

```vbnet
Public Sub RunByName(control As IRibbonControl)
    Dim macro As String
    macro = control.Tag
    If Len(macro) = 0 Then macro = control.Id
    On Error GoTo errh
    Application.Run macro
    Exit Sub
errh:
    MsgBox "Macro not found: " & macro, vbExclamation
End Sub
```

## Module `SubConvert`

### `ConvertUStoUK`

```vbnet
Public Sub ConvertUStoUK()

    If ActivePresentation.Slides.Count = 0 Then
        MsgBox "Presentation has no slides.", vbInformation, "US to UK English"
        Exit Sub
    End If

    BuildDictionary
    m_totalReplaced = 0

    Dim sld As Slide
    Dim shp As Shape
    Dim slideNum As Long

    For Each sld In ActivePresentation.Slides
        slideNum = slideNum + 1
        Application.StatusBar = "Converting slide " & slideNum & " of " & ActivePresentation.Slides.Count & "..."

        For Each shp In sld.Shapes
            ProcessShape shp
        Next shp
    Next sld

    ' Also process slide masters and layouts
    Dim dsgn As Design
    For Each dsgn In ActivePresentation.Designs
        For Each shp In dsgn.SlideMaster.Shapes
            ProcessShape shp
        Next shp

        Dim lay As CustomLayout
        For Each lay In dsgn.SlideMaster.CustomLayouts
            For Each shp In lay.Shapes
                ProcessShape shp
            Next shp
        Next lay
    Next dsgn

    ' Notes pages
    slideNum = 0
    For Each sld In ActivePresentation.Slides
        slideNum = slideNum + 1
        If sld.HasNotesPage Then
            Dim nshp As Shape
            For Each nshp In sld.NotesPage.Shapes
                If nshp.HasTextFrame Then
                    If nshp.TextFrame.HasText Then
                        ProcessTextFrame nshp.TextFrame
                    End If
                End If
            Next nshp
        End If
    Next sld

    Application.StatusBar = ""
    Set m_dict = Nothing

    If m_totalReplaced > 0 Then
        MsgBox m_totalReplaced & " word(s) converted across " & _
               ActivePresentation.Slides.Count & " slide(s)." & vbCrLf & _
               "Use Ctrl+Z to undo.", vbInformation, "US to UK English"
    Else
        MsgBox "No US English words found.", vbInformation, "US to UK English"
    End If

End Sub
```

### `ProcessShape`

```vbnet
Private Sub ProcessShape(shp As Shape)

    On Error Resume Next

    ' Grouped shapes
    If shp.Type = msoGroup Then
        Dim subShp As Shape
        For Each subShp In shp.GroupItems
            ProcessShape subShp
        Next subShp
        Exit Sub
    End If

    ' Tables
    If shp.HasTable Then
        Dim row As Long, col As Long
        For row = 1 To shp.Table.Rows.Count
            For col = 1 To shp.Table.Columns.Count
                ProcessTextFrame shp.Table.Cell(row, col).Shape.TextFrame
            Next col
        Next row
        Exit Sub
    End If

    ' SmartArt
    If shp.HasSmartArt Then
        Dim nd As SmartArtNode
        For Each nd In shp.SmartArt.AllNodes
            If Not nd.TextFrame2 Is Nothing Then
                ProcessTextFrame2 nd.TextFrame2
            End If
        Next nd
        Exit Sub
    End If

    ' Charts (title and axis labels)
    If shp.HasChart Then
        Dim cht As Chart
        Set cht = shp.Chart
        If cht.HasTitle Then
            ProcessTextFrame2 cht.ChartTitle.Format.TextFrame2
        End If
        Exit Sub
    End If

    ' Regular text frame
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            ProcessTextFrame shp.TextFrame
        End If
    End If

    On Error GoTo 0

End Sub
```

### `ProcessTextFrame`

```vbnet
Private Sub ProcessTextFrame(tf As TextFrame)
    On Error Resume Next
    If tf.HasText Then
        ReplaceInTextRange tf.TextRange
    End If
    On Error GoTo 0
End Sub
```

### `ProcessTextFrame2`

```vbnet
Private Sub ProcessTextFrame2(tf2 As TextFrame2)
    On Error Resume Next
    If tf2.HasText Then
        ReplaceInTextRange tf2.TextRange
    End If
    On Error GoTo 0
End Sub
```

### `ReplaceInTextRange`

```vbnet
Private Sub ReplaceInTextRange(tr As TextRange)
    ' Walk words in reverse so that index positions stay valid
    ' after a replacement changes string length.

    Dim i As Long
    Dim w As TextRange
    Dim wordText As String
    Dim lookupKey As String
    Dim replacement As String
    Dim isCapitalised As Boolean
    Dim isAllCaps As Boolean

    On Error Resume Next

    For i = tr.Words.Count To 1 Step -1
        Set w = tr.Words(i)
        wordText = Trim$(w.Text)

        ' Strip trailing punctuation for lookup
        Dim clean As String
        clean = StripPunctuation(wordText)

        If Len(clean) = 0 Then GoTo NextWord

        ' Determine case pattern
        lookupKey = LCase$(clean)

        If Not m_dict.Exists(lookupKey) Then GoTo NextWord

        replacement = m_dict(lookupKey)

        ' Preserve original case
        isAllCaps = (clean = UCase$(clean) And Len(clean) > 1)
        isCapitalised = (Left$(clean, 1) = UCase$(Left$(clean, 1))) And Not isAllCaps

        If isAllCaps Then
            replacement = UCase$(replacement)
        ElseIf isCapitalised Then
            replacement = UCase$(Left$(replacement, 1)) & Mid$(replacement, 2)
        End If

        ' Replace only the clean portion, preserving trailing punctuation
        Dim suffix As String
        suffix = Mid$(wordText, Len(clean) + 1)
        w.Text = replacement & suffix

        m_totalReplaced = m_totalReplaced + 1

NextWord:
    Next i

    On Error GoTo 0

End Sub
```

### `StripPunctuation`

```vbnet
Private Function StripPunctuation(s As String) As String
    ' Remove trailing punctuation (.,;:!? etc.) for dictionary lookup
    Dim i As Long
    For i = Len(s) To 1 Step -1
        Select Case Mid$(s, i, 1)
            Case ".", ",", ";", ":", "!", "?", """", "'", ")", "]", "}", " ", vbCr, vbLf
                ' continue stripping
            Case Else
                StripPunctuation = Left$(s, i)
                Exit Function
        End Select
    Next i
    StripPunctuation = ""
End Function
```

### `BuildDictionary`

```vbnet
Private Sub BuildDictionary()

    Set m_dict = CreateObject("Scripting.Dictionary")
    m_dict.CompareMode = vbTextCompare  ' Case-insensitive keys

    ' ----- -ize / -ise (with inflections) -----
    Dim izeRoots As Variant
    izeRoots = Array( _
        "recogn", "organ", "real", "minim", "maxim", "optim", "util", "author", _
        "categor", "character", "custom", "emphas", "final", "global", "harmon", _
        "initial", "legal", "memor", "modern", "neutral", "normal", "prior", _
        "special", "standard", "summar", "symbol", "synchron", "apolog", _
        "capital", "central", "critic", "digit", "dramat", "familiar", _
        "fertil", "general", "hospital", "hypothes", "ideal", "immun", _
        "item", "jeopard", "liberal", "local", "marginal", "material", _
        "mechan", "mobil", "monopol", "national", "penal", "polar", _
        "privat", "revolution", "scrutin", "sensit", "social", "stabil", _
        "steril", "subsid", "terror", "traumat", "trivial", "vandal", _
        "vapor", "visual")

    Dim stem As Variant
    For Each stem In izeRoots
        AddPair CStr(stem) & "ize", CStr(stem) & "ise"
        AddPair CStr(stem) & "izes", CStr(stem) & "ises"
        AddPair CStr(stem) & "ized", CStr(stem) & "ised"
        AddPair CStr(stem) & "izing", CStr(stem) & "ising"
        AddPair CStr(stem) & "izer", CStr(stem) & "iser"
        AddPair CStr(stem) & "izers", CStr(stem) & "isers"
        AddPair CStr(stem) & "ization", CStr(stem) & "isation"
        AddPair CStr(stem) & "izations", CStr(stem) & "isations"
    Next stem

    ' ----- -or / -our (with inflections) -----
    AddGroup "color", "colour", True
    AddGroup "favor", "favour", True
    AddPair "favorable", "favourable"
    AddPair "favorite", "favourite"
    AddPair "favorites", "favourites"
    AddGroup "honor", "honour", True
    AddPair "honorable", "honourable"
    AddGroup "humor", "humour", True
    AddPair "humorous", "humourous"
    AddGroup "labor", "labour", True
    AddGroup "neighbor", "neighbour", True
    AddPair "neighborhood", "neighbourhood"
    AddPair "neighborhoods", "neighbourhoods"
    AddGroup "behavior", "behaviour", False
    AddPair "behavioral", "behavioural"
    AddGroup "flavor", "flavour", True
    AddGroup "harbor", "harbour", True
    AddGroup "rumor", "rumour", True
    AddGroup "tumor", "tumour", False
    AddPair "tumors", "tumours"
    AddPair "valor", "valour"
    AddPair "vigor", "vigour"
    AddPair "vigorous", "vigourous"

    ' ----- -er / -re (with inflections) -----
    AddPair "center", "centre"
    AddPair "centers", "centres"
    AddPair "centered", "centred"
    AddPair "centering", "centring"
    AddPair "fiber", "fibre"
    AddPair "fibers", "fibres"
    AddPair "liter", "litre"
    AddPair "liters", "litres"
    AddPair "meter", "metre"
    AddPair "meters", "metres"
    AddPair "theater", "theatre"
    AddPair "theaters", "theatres"

    ' ----- exact words -----
    AddPair "aging", "ageing"
    AddPair "airplane", "aeroplane"
    AddPair "airplanes", "aeroplanes"
    AddPair "aluminum", "aluminium"
    AddPair "cozy", "cosy"
    AddPair "gray", "grey"
    AddPair "grays", "greys"
    AddPair "judgment", "judgement"
    AddPair "judgments", "judgements"
    AddPair "math", "maths"
    AddPair "program", "programme"
    AddPair "programs", "programmes"
    AddPair "check", "cheque"
    AddPair "checks", "cheques"
    AddPair "curb", "kerb"
    AddPair "curbs", "kerbs"
    AddPair "jewelry", "jewellery"
    AddPair "skillful", "skilful"
    AddPair "skillfully", "skilfully"

End Sub
```

### `AddPair`

```vbnet
Private Sub AddPair(us As String, uk As String)
    If Not m_dict.Exists(LCase$(us)) Then
        m_dict.Add LCase$(us), LCase$(uk)
    End If
End Sub
```

### `AddGroup`

```vbnet
Private Sub AddGroup(usBase As String, ukBase As String, addIng As Boolean)
    ' Adds base + s + ed (+ ing if flag set)
    AddPair usBase, ukBase
    AddPair usBase & "s", ukBase & "s"
    AddPair usBase & "ed", ukBase & "ed"
    If addIng Then AddPair usBase & "ing", ukBase & "ing"
End Sub
```

## Module `Subs`

### `FontArial`

```vbnet
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
```

### `FontEY`

```vbnet
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
```

### `FontSize12`

```vbnet
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
```

### `FontSizeUp`

```vbnet
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
```

### `FontSizeDown`

```vbnet
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
```

### `SelectedTableBorders`

```vbnet
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
```

### `SelectedTableShade`

```vbnet
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

    ' Row 1 is header  skip it
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
```

### `TableNormalMargin`

```vbnet
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
```

### `SelectedTableFormatReset`

```vbnet
Sub SelectedTableFormatReset()
    Dim shp As Shape
    Dim tbl As Table
    Dim r As Long, c As Long
    
    ' Check if selection is a table
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set shp = ActiveWindow.Selection.ShapeRange(1)
        
        If shp.HasTable Then
            Set tbl = shp.Table
            
            ' Loop through all cells
            For r = 1 To tbl.Rows.Count
                For c = 1 To tbl.Columns.Count
                    With tbl.Cell(r, c).Shape.TextFrame.TextRange
                        .Font.Name = "Arial"
                        .Font.Size = 18
                        .Font.Bold = msoFalse
                        .Font.Italic = msoFalse
                        .Font.Underline = msoFalse
                        .ParagraphFormat.Alignment = ppAlignLeft
                    End With
                    
                    ' Remove cell fill and borders
                    tbl.Cell(r, c).Shape.Fill.Visible = msoFalse
                    tbl.Cell(r, c).Borders(ppBorderTop).Visible = msoFalse
                    tbl.Cell(r, c).Borders(ppBorderBottom).Visible = msoFalse
                    tbl.Cell(r, c).Borders(ppBorderLeft).Visible = msoFalse
                    tbl.Cell(r, c).Borders(ppBorderRight).Visible = msoFalse
                Next c
            Next r
            
            MsgBox "Table formatting has been reset.", vbInformation
        Else
            MsgBox "Selected shape is not a table.", vbExclamation
        End If
    Else
        MsgBox "Please select a table first.", vbExclamation
    End If
End Sub
```
