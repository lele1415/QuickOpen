Dim pProjectText : pProjectText = oWs.CurrentDirectory & "\res\project.ini"

Dim vaWorksInfo : Set vaWorksInfo = New VariableArray

Call readWorksInfoText()

Sub readWorksInfoText()
    If Not oFso.FileExists(pProjectText) Then Exit Sub
    
    Dim oText, sReadLine
    Set oText = oFso.OpenTextFile(pProjectText, FOR_READING, False, True)

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
    Dim oInfos
    Set oInfos = New ProjectInfos
    oInfos.Work = Trim(oText.ReadLine)
    oInfos.Sdk = Trim(oText.ReadLine)
    oInfos.Product = Trim(oText.ReadLine)
    oInfos.Project = Trim(oText.ReadLine)

    vaWorksInfo.Append(oInfos)
End Sub
