Option Explicit

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
    Dim fileNum As Integer
    Dim fileData() As Byte
    Dim xml As Object
    Dim node As Object
    Dim pic As Shape
    Dim sld As Slide
    Dim ratio As Double

    If Len(base64String) = 0 Then
        MsgBox "Image data is empty.", vbCritical
        Exit Sub
    End If

    ' Decode base64
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.text = base64String
    fileData = node.nodeTypedValue

    ' Write to temp file
    tempPath = Environ("TEMP") & "\deckui_temp_img.png"
    fileNum = FreeFile
    Open tempPath For Binary As #fileNum
    Put #fileNum, , fileData
    Close #fileNum

    ' Get active slide
    On Error GoTo errNoSlide
    Set sld = ActiveWindow.View.Slide
    On Error GoTo 0

    ' Insert picture onto the active slide
    Set pic = sld.Shapes.AddPicture( _
        FileName:=tempPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=100, _
        Top:=100)

    ' Resize only if widthCm is specified
    If widthCm > 0 Then
        ratio = pic.Height / pic.Width
        pic.LockAspectRatio = msoTrue
        pic.Width = Application.CentimetersToPoints(widthCm)
        pic.Height = pic.Width * ratio
    End If

    ' Select the new picture so user can move it
    pic.Select

    Kill tempPath
    Exit Sub

errNoSlide:
    MsgBox "Please select a slide before inserting an image.", vbExclamation
    If Dir(tempPath) <> "" Then Kill tempPath

End Sub

' ===== HELPER: Run this once per image to get Base64 =====

Public Sub ConvertImageToBase64()

    Dim fd As FileDialog
    Dim filePath As String
    Dim fileNum As Integer
    Dim fileData() As Byte
    Dim xml As Object
    Dim node As Object
    Dim outPath As String
    Dim base64 As String
    Dim vbaCode As String
    Dim i As Long
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
        base64 = node.text
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
