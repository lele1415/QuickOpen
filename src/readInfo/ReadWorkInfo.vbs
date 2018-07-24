Const ID_DIV_SHORTCUT = "div_shortcut"
Const ID_WORK_NAME = "work_name"

Dim pWorkInfoText : pWorkInfoText = oWs.CurrentDirectory & "\shortcutsInfo.ini"

Dim vaWorksInfo : Set vaWorksInfo = New VariableArray

Call readWorksInfoText()

Sub readWorksInfoText()
    If Not oFso.FileExists(pWorkInfoText) Then Exit Sub
    
    Dim oText, sReadLine
    Set oText = oFso.OpenTextFile(pWorkInfoText, FOR_READING, False, True)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        If InStr(sReadLine, "########") > 0 Then
            Call handleForWorksInfo(oText, sReadLine)
        End If
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub handleForWorksInfo(oText, sReadLine)
    Dim oWork
    Set oWork = New WorkInfo
    oWork.WorkName = Trim(oText.ReadLine)
    oWork.WorkCode = Trim(oText.ReadLine)
    oWork.WorkPrj = Trim(oText.ReadLine)
    oWork.WorkOpt = Trim(oText.ReadLine)

    vaWorksInfo.Append(oWork)
End Sub
