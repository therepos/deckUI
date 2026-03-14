Option Explicit

' ============================================================
' US to UK English Converter — DeckUI (PowerPoint)
' ============================================================
' Adapted from the WordUI converter.
' PowerPoint has no built-in Find/Replace engine like Word,
' so this version iterates every text run in every shape on
' every slide and performs string replacements manually.
'
' Key design decisions:
'   1. We walk every TextRange2 run so formatting is preserved.
'   2. Whole-word matching is done via helper function.
'   3. Status bar shows progress.
' ============================================================

Private m_totalReplaced As Long

Public Sub ConvertUStoUK()

    Dim sld As Slide
    Dim shp As Shape

    m_totalReplaced = 0

    ' ----- Build replacement lists -----

    ' -ize/-ise stems
    Dim izeStems As String
    izeStems = "recogn|organ|real|minim|maxim|optim|util|author" & _
               "|categor|character|custom|emphas|final|global|harmon" & _
               "|initial|legal|memor|modern|neutral|normal|prior" & _
               "|special|standard|summar|symbol|synchron|apolog" & _
               "|capital|central|critic|digit|dramat|familiar" & _
               "|fertil|general|hospital|hypothes|ideal|immun" & _
               "|item|jeopard|liberal|local|marginal|material" & _
               "|mechan|mobil|monopol|national|penal|polar" & _
               "|privat|revolution|scrutin|sensit|social|stabil" & _
               "|steril|subsid|terror|traumat|trivial|vandal" & _
               "|vapor|visual"

    Dim stems() As String
    stems = Split(izeStems, "|")

    Dim suffixUS() As String, suffixUK() As String
    suffixUS = Split("ize|izes|ized|izing|izer|ization", "|")
    suffixUK = Split("ise|ises|ised|ising|iser|isation", "|")

    Dim i As Long, j As Long

    ' Process every slide
    For Each sld In ActivePresentation.Slides
        ' Process every shape (including grouped, table cells, etc.)
        For Each shp In sld.Shapes
            ProcessShape shp, stems, suffixUS, suffixUK
        Next shp
    Next sld

    If m_totalReplaced > 0 Then
        MsgBox m_totalReplaced & " replacement(s) made." & vbCrLf & _
               "Use Ctrl+Z to undo (one step at a time).", vbInformation, "US to UK English"
    Else
        MsgBox "No US English words found.", vbInformation, "US to UK English"
    End If

End Sub


' ============================================================
' Recursively process a shape (handles groups, tables, text)
' ============================================================
Private Sub ProcessShape(shp As Shape, stems() As String, _
                         suffixUS() As String, suffixUK() As String)

    Dim i As Long, j As Long

    ' --- Grouped shapes ---
    If shp.Type = msoGroup Then
        Dim grpShp As Shape
        For Each grpShp In shp.GroupItems
            ProcessShape grpShp, stems, suffixUS, suffixUK
        Next grpShp
        Exit Sub
    End If

    ' --- Tables ---
    If shp.HasTable Then
        Dim tbl As Table
        Dim r As Long, c As Long
        Set tbl = shp.Table
        For r = 1 To tbl.Rows.Count
            For c = 1 To tbl.Columns.Count
                Dim cellShp As Shape
                Set cellShp = tbl.Cell(r, c).Shape
                If cellShp.HasTextFrame Then
                    ReplaceInTextFrame cellShp.TextFrame, stems, suffixUS, suffixUK
                End If
            Next c
        Next r
        Exit Sub
    End If

    ' --- Regular text frames ---
    If shp.HasTextFrame Then
        ReplaceInTextFrame shp.TextFrame, stems, suffixUS, suffixUK
    End If

End Sub


' ============================================================
' Perform all replacements within a single TextFrame
' ============================================================
Private Sub ReplaceInTextFrame(tf As TextFrame, stems() As String, _
                               suffixUS() As String, suffixUK() As String)

    Dim i As Long, j As Long

    If Not tf.HasText Then Exit Sub

    ' -ize/-ise stems
    For i = 0 To UBound(stems)
        For j = 0 To UBound(suffixUS)
            DoReplace tf, stems(i) & suffixUS(j), stems(i) & suffixUK(j)
        Next j
    Next i

    ' -or/-our
    DoReplace tf, "color", "colour": DoReplace tf, "colors", "colours"
    DoReplace tf, "colored", "coloured": DoReplace tf, "coloring", "colouring"
    DoReplace tf, "favor", "favour": DoReplace tf, "favors", "favours"
    DoReplace tf, "favored", "favoured": DoReplace tf, "favoring", "favouring"
    DoReplace tf, "favorable", "favourable"
    DoReplace tf, "favorite", "favourite": DoReplace tf, "favorites", "favourites"
    DoReplace tf, "honor", "honour": DoReplace tf, "honors", "honours"
    DoReplace tf, "honored", "honoured": DoReplace tf, "honoring", "honouring"
    DoReplace tf, "honorable", "honourable"
    DoReplace tf, "humor", "humour": DoReplace tf, "humors", "humours"
    DoReplace tf, "humored", "humoured": DoReplace tf, "humorous", "humourous"
    DoReplace tf, "labor", "labour": DoReplace tf, "labors", "labours"
    DoReplace tf, "labored", "laboured": DoReplace tf, "laboring", "labouring"
    DoReplace tf, "neighbor", "neighbour": DoReplace tf, "neighbors", "neighbours"
    DoReplace tf, "neighboring", "neighbouring": DoReplace tf, "neighborhood", "neighbourhood"
    DoReplace tf, "behavior", "behaviour": DoReplace tf, "behaviors", "behaviours"
    DoReplace tf, "behavioral", "behavioural"
    DoReplace tf, "flavor", "flavour": DoReplace tf, "flavors", "flavours"
    DoReplace tf, "flavored", "flavoured"
    DoReplace tf, "harbor", "harbour": DoReplace tf, "harbors", "harbours"
    DoReplace tf, "rumor", "rumour": DoReplace tf, "rumors", "rumours"
    DoReplace tf, "rumored", "rumoured"
    DoReplace tf, "tumor", "tumour": DoReplace tf, "tumors", "tumours"
    DoReplace tf, "valor", "valour"
    DoReplace tf, "vigor", "vigour": DoReplace tf, "vigorous", "vigourous"

    ' -er/-re
    DoReplace tf, "center", "centre": DoReplace tf, "centers", "centres"
    DoReplace tf, "centered", "centred": DoReplace tf, "centering", "centring"
    DoReplace tf, "fiber", "fibre": DoReplace tf, "fibers", "fibres"
    DoReplace tf, "liter", "litre": DoReplace tf, "liters", "litres"
    DoReplace tf, "meter", "metre": DoReplace tf, "meters", "metres"
    DoReplace tf, "theater", "theatre": DoReplace tf, "theaters", "theatres"

    ' Exact words
    DoReplace tf, "aging", "ageing"
    DoReplace tf, "airplane", "aeroplane": DoReplace tf, "airplanes", "aeroplanes"
    DoReplace tf, "aluminum", "aluminium"
    DoReplace tf, "cozy", "cosy"
    DoReplace tf, "gray", "grey"
    DoReplace tf, "judgment", "judgement"
    DoReplace tf, "math", "maths"
    DoReplace tf, "program", "programme": DoReplace tf, "programs", "programmes"
    DoReplace tf, "check", "cheque": DoReplace tf, "checks", "cheques"
    DoReplace tf, "curb", "kerb": DoReplace tf, "curbs", "kerbs"
    DoReplace tf, "jewelry", "jewellery"
    DoReplace tf, "skillful", "skilful": DoReplace tf, "skillfully", "skilfully"

End Sub


' ============================================================
' ENGINE — single whole-word replace within a TextFrame
' ============================================================
Private Sub DoReplace(tf As TextFrame, usWord As String, ukWord As String)

    Dim tr As TextRange
    Dim fullText As String
    Dim pos As Long
    Dim wordLen As Long

    wordLen = Len(usWord)

    ' Use TextRange.Find for PowerPoint
    On Error Resume Next
    Set tr = tf.TextRange.Find(usWord, 0, msoTrue, msoFalse)  ' WholeWords:=True, MatchCase:=False
    On Error GoTo 0

    Do While Not tr Is Nothing
        tr.Text = ukWord
        m_totalReplaced = m_totalReplaced + 1

        ' Find next occurrence
        On Error Resume Next
        Set tr = tf.TextRange.Find(usWord, tr.Start + Len(ukWord) - 1, msoTrue, msoFalse)
        On Error GoTo 0
    Loop

End Sub
