Public Ribbon As IRibbonUI

Public Sub RibbonOnLoad(r As IRibbonUI)
    Set Ribbon = r
End Sub

Public Sub RunByName(control As IRibbonControl)
    Dim macro As String
    macro = control.Tag
    If Len(macro) = 0 Then macro = control.Id
    On Error GoTo errh
    Application.run macro
    Exit Sub
errh:
    MsgBox "Macro not found: " & macro, vbExclamation
End Sub

