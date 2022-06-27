Option Explicit

Dim vaAndroidVer : Set vaAndroidVer = New VariableArray



Sub onSdkPathInputClick()
    Call mSdkPathList.toggleList()
End Sub

Sub readSdkPathText(DictPath)
    If Not oFso.FileExists(DictPath) Then Exit Sub
    
    Dim oText, sReadLine
    Set oText = oFso.OpenTextFile(DictPath, FOR_READING)

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
    If Not oFso.FileExists(path) Then Exit Sub
    
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


