Option Explicit

' ============================================================
' Image Inserter — PowerPoint Edition
' ============================================================
' Differences from Word version:
'   - Uses ActiveWindow.Selection + Shapes.AddPicture
'   - Positions image at centre of slide (or at selected shape
'     location if a shape is selected)
'   - Uses Application.CentimetersToPoints equivalent
'     (PowerPoint uses points natively)
' ============================================================

Public Sub InsertLogoEYWhite()
    InsertBase64Img GetLogoEYWhite(), 2.93
End Sub

Public Sub InsertLogoEYOffBlack()
    InsertBase64Img GetLogoEYOffBlack(), 2.93
End Sub

Public Sub InsertLogoEYSTFWCWhite()
    InsertBase64Img GetLogoEYSTFWCWhite(), 2.93
End Sub

Public Sub InsertLogoEYSTFWCOffBlack()
    InsertBase64Img GetLogoEYSTFWCOffBlack(), 2.93
End Sub

Public Sub InsertSignatureEYAPL()
    InsertBase64Img GetSignatureEYAPL()
End Sub

Public Sub InsertSealEYAPL()
    InsertBase64Img GetSealEYAPL()
End Sub

Public Sub InsertSealEYAPLRound()
    InsertBase64Img GetSealEYAPLRound(), 2.93
End Sub

' ===== CORE INSERT LOGIC =====

Private Sub InsertBase64Img(base64String As String, Optional widthCm As Double = 0)

    Dim tempPath As String
    Dim fileNum  As Integer
    Dim fileData() As Byte
    Dim xml      As Object
    Dim node     As Object
    Dim pic      As Shape
    Dim ratio    As Double
    Dim sld      As Slide
    Dim posLeft  As Single
    Dim posTop   As Single

    If Len(base64String) = 0 Then
        MsgBox "Image data is empty.", vbCritical
        Exit Sub
    End If

    ' --- Decode base64 to temp file ---
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = base64String
    fileData = node.nodeTypedValue

    tempPath = Environ("TEMP") & "\ey_temp_img.png"
    fileNum = FreeFile
    Open tempPath For Binary As #fileNum
    Put #fileNum, , fileData
    Close #fileNum

    ' --- Determine target slide ---
    On Error Resume Next
    Set sld = ActiveWindow.View.Slide
    On Error GoTo 0

    If sld Is Nothing Then
        MsgBox "Please select a slide first.", vbExclamation, "Insert Image"
        Kill tempPath
        Exit Sub
    End If

    ' --- Default position: centre of slide ---
    posLeft = ActivePresentation.PageSetup.SlideWidth / 2
    posTop = ActivePresentation.PageSetup.SlideHeight / 2

    ' If a shape is currently selected, use its position
    On Error Resume Next
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        posLeft = ActiveWindow.Selection.ShapeRange(1).Left
        posTop = ActiveWindow.Selection.ShapeRange(1).Top
    End If
    On Error GoTo 0

    ' --- Insert the picture ---
    Set pic = sld.Shapes.AddPicture( _
        FileName:=tempPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=0, _
        Top:=0)

    ' --- Resize if widthCm specified ---
    If widthCm > 0 Then
        ratio = pic.Height / pic.Width
        pic.LockAspectRatio = msoTrue
        pic.Width = CmToPoints(widthCm)
        pic.Height = pic.Width * ratio
    End If

    ' --- Centre on target position ---
    pic.Left = posLeft - (pic.Width / 2)
    pic.Top = posTop - (pic.Height / 2)

    ' Select the newly inserted picture
    pic.Select

    ' --- Cleanup ---
    Kill tempPath

End Sub

' ===== UNIT CONVERSION =====
' PowerPoint VBA doesn't have CentimetersToPoints,
' so we provide our own.

Private Function CmToPoints(cm As Double) As Double
    CmToPoints = cm * 28.3464567
End Function

' ===== HELPER: Run this once per image to get Base64 =====

Public Sub ConvertImageToBase64()

    Dim fd       As FileDialog
    Dim filePath As String
    Dim fileNum  As Integer
    Dim fileData() As Byte
    Dim xml      As Object
    Dim node     As Object
    Dim outPath  As String
    Dim base64   As String
    Dim vbaCode  As String
    Dim i        As Long
    Dim chunkSize As Long

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select an image file"
    fd.Filters.Add "Images", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"

    If fd.Show = -1 Then
        filePath = fd.SelectedItems(1)

        fileNum = FreeFile
        Open filePath For Binary As #fileNum
        ReDim fileData(LOF(fileNum) - 1)
        Get #fileNum, , fileData
        Close #fileNum

        Set xml = CreateObject("MSXML2.DOMDocument")
        Set node = xml.createElement("b64")
        node.DataType = "bin.base64"
        node.nodeTypedValue = fileData

        ' Strip all whitespace
        base64 = node.Text
        base64 = Replace(base64, vbCrLf, "")
        base64 = Replace(base64, vbCr, "")
        base64 = Replace(base64, vbLf, "")
        base64 = Replace(base64, " ", "")

        ' Build VBA-ready code in chunks
        chunkSize = 70
        vbaCode = "    Dim s As String" & vbCrLf
        vbaCode = vbaCode & "    s = """"" & vbCrLf

        For i = 1 To Len(base64) Step chunkSize
            vbaCode = vbaCode & "    s = s & """ & Mid(base64, i, chunkSize) & """" & vbCrLf
        Next i

        outPath = Environ("TEMP") & "\img_base64_vba.txt"
        fileNum = FreeFile
        Open outPath For Output As #fileNum
        Print #fileNum, vbaCode
        Close #fileNum

        MsgBox "VBA-ready code saved to: " & outPath
        Shell "notepad.exe " & outPath, vbNormalFocus
    End If

End Sub
