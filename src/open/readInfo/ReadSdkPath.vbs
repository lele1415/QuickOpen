Option Explicit

Dim vaAndroidVer : Set vaAndroidVer = New VariableArray



Sub onSdkPathInputClick()
    Call mSdkPathList.toggleList()
End Sub

Sub addSdkPathList()
    Call readSdkPathText()
    Call mSdkPathList.addList(vaAndroidVer)
End Sub

Sub readSdkPathText()
    If Not isFileExists(pSdkPathText) Then Exit Sub
    
    Dim oText, sReadLine
    Set oText = oFso.OpenTextFile(pSdkPathText, FOR_READING)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine

        If InStr(sReadLine, "{") > 0 Then
            Call getAllSdkPath(oText, sReadLine)
        End If
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub getAllSdkPath(oText, sReadLine)
    Dim verStr : verStr = Trim(Replace(sReadLine, "{", ""))
    if verStr <> "" Then
        Dim vaSdkPath : Set vaSdkPath = New VariableArray
        vaSdkPath.Name = verStr

        sReadLine = oText.ReadLine
        Do until InStr(sReadLine, "}") > 0
            vaSdkPath.Append(Trim(sReadLine))
            sReadLine = oText.ReadLine
        Loop

        vaAndroidVer.Append(vaSdkPath)
    End If
End Sub

Sub readTextAndDoSomething(path, strFun)
    If Not isFileExists(path) Then Exit Sub
    
    Dim oText, sReadLine, exitFlag
    Set oText = oFso.OpenTextFile(path, FOR_READING)
    exitFlag = False

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        Execute strFun
        If exitFlag Then Exit Do
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub openSdkText()
    Call runPath(pSdkPathText)
End Sub
