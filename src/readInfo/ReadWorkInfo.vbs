Option Explicit

Dim pProjectText : pProjectText = oWs.CurrentDirectory & "\res\project.ini"

Dim vaWorksInfo : Set vaWorksInfo = New VariableArray

Sub readWorksInfoText()
    If Not isFileExists(pProjectText) Then Exit Sub
    
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

Sub applyLastWorkInfo()
    Dim oInfos
    If vaWorksInfo.Bound > -1 Then
        Set oInfos = vaWorksInfo.V(vaWorksInfo.Bound)
        if Not checkProjectExist(oInfos.Sdk, oInfos.Product, oInfos.Project) Then
            Call mIp.clearSdkInfos()
            Call mIp.clearWorkInfos()
            Exit Sub
        End If
        mIp.Work = oInfos.Work
        mIp.Sdk = oInfos.Sdk
        mIp.Product = oInfos.Product
        mIp.Project = oInfos.Project
        mIp.Firmware = oInfos.Firmware
        mIp.Requirements = oInfos.Requirements
        mIp.Zentao = oInfos.Zentao
        Call updateProductList()
    End If
End Sub

Sub handleForWorksInfo(oText, sReadLine)
    Dim i, sLine, oInfos
    i = 0
    sLine = oText.ReadLine
    Set oInfos = New ProjectInfos

    Do Until (Trim(sLine) = "" Or i > 7)
        i = i + 1
        Select Case i
            Case 1 : oInfos.Work = Trim(sLine)
            Case 2 : oInfos.Sdk = Trim(sLine)
            Case 3 : oInfos.Product = Trim(sLine)
            Case 4 : oInfos.Project = Trim(sLine)
            Case 5 : oInfos.Firmware = Trim(sLine)
            Case 6 : oInfos.Requirements = Trim(sLine)
            Case 7 : oInfos.Zentao = Trim(sLine)
        End Select
        sLine = oText.ReadLine
    Loop

    vaWorksInfo.Append(oInfos)
End Sub
