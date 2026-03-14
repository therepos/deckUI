Option Explicit

' ============================================================
' US to UK English Converter — PowerPoint Edition
' ============================================================
' PowerPoint has no Find/Replace API, so we walk every text
' range word-by-word and swap via a Dictionary (O(1) lookup).
' ============================================================

Private m_dict As Object  ' Scripting.Dictionary
Private m_totalReplaced As Long

Public Sub ConvertUStoUK()

    If ActivePresentation.Slides.Count = 0 Then
        MsgBox "Presentation has no slides.", vbInformation, "US to UK English"
        Exit Sub
    End If

    BuildDictionary
    m_totalReplaced = 0

    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            ProcessShape shp
        Next shp
    Next sld

    ' Slide masters and layouts
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
    For Each sld In ActivePresentation.Slides
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

    Set m_dict = Nothing

    If m_totalReplaced > 0 Then
        MsgBox m_totalReplaced & " word(s) converted across " & _
               ActivePresentation.Slides.Count & " slide(s)." & vbCrLf & _
               "Use Ctrl+Z to undo.", vbInformation, "US to UK English"
    Else
        MsgBox "No US English words found.", vbInformation, "US to UK English"
    End If

End Sub

' ============================================================
' SHAPE WALKER — recurse into groups, tables, charts
' ============================================================

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

    ' Charts (title)
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

' ============================================================
' TEXT REPLACEMENT ENGINE — word-by-word via TextRange
' ============================================================

Private Sub ProcessTextFrame(tf As TextFrame)
    On Error Resume Next
    If tf.HasText Then
        ReplaceInTextRange tf.TextRange
    End If
    On Error GoTo 0
End Sub

Private Sub ProcessTextFrame2(tf2 As TextFrame2)
    On Error Resume Next
    If tf2.HasText Then
        ReplaceInTextRange tf2.TextRange
    End If
    On Error GoTo 0
End Sub

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

' ============================================================
' DICTIONARY BUILDER — all pairs loaded once into a hash map
' ============================================================

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

Private Sub AddPair(us As String, uk As String)
    If Not m_dict.Exists(LCase$(us)) Then
        m_dict.Add LCase$(us), LCase$(uk)
    End If
End Sub

Private Sub AddGroup(usBase As String, ukBase As String, addIng As Boolean)
    ' Adds base + s + ed (+ ing if flag set)
    AddPair usBase, ukBase
    AddPair usBase & "s", ukBase & "s"
    AddPair usBase & "ed", ukBase & "ed"
    If addIng Then AddPair usBase & "ing", ukBase & "ing"
End Sub
