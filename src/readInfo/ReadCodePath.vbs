

Const VALUE_SELECT_SDK_PATH_SHOW = "选择代码"
Const VALUE_SELECT_SDK_PATH_HIDE = "收起"

Dim pConfigText : pConfigText = oWs.CurrentDirectory & "\config.ini"
Dim pCodeText : pCodeText = oWs.CurrentDirectory & "\code.ini"

Dim mTextEditorPath


Dim vaAndroidVer : Set vaAndroidVer = New VariableArray

Call readConfigText(pConfigText)
Call readCodeText(pCodeText)
Call setSdkPathIds()
Call addCodeList()



Sub setSdkPathIds()
    Call setListParentAndInputIds(getParentSdkPathId(), getSdkPathInputId())
    Call setListDirectoryDivId(getSdkPathDirectoryDivId())
End Sub

Sub setSdkPathDirectoryIds()
    Call setListDivIds(getSdkPathDirectoryDivId(), getSdkPathDirectoryULId())
End Sub

Sub setSdkPathListIds(category)
    Call setListDivIds(getSdkPathDivId() & category, getSdkPathULId() & category)
End Sub

Sub readConfigText(DictPath)
    If Not oFso.FileExists(DictPath) Then Exit Sub
    
    Dim oText
    Set oText = oFso.OpenTextFile(DictPath, FOR_READING)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        If InStr(sReadLine, "TextEditorPath") > 0 Then getTextEditor(sReadLine)
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub getTextEditor(sReadLine)
    mTextEditorPath = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
End Sub

Sub readCodeText(DictPath)
    If Not oFso.FileExists(DictPath) Then Exit Sub
    
    Dim oText
    Set oText = oFso.OpenTextFile(DictPath, FOR_READING)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine

        If InStr(sReadLine, "{") > 0 Then
            Call getAllCode(oText, sReadLine)
        End If
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub getAllCode(oText, sReadLine)
    Dim verStr : verStr = Trim(Replace(sReadLine, "{", ""))
    if verStr <> "" Then
        Dim vaCode : Set vaCode = New VariableArray
        vaCode.Name = verStr

        sReadLine = oText.ReadLine
        Do until InStr(sReadLine, "}") > 0
            vaCode.Append(Trim(sReadLine))
            sReadLine = oText.ReadLine
        Loop

        vaAndroidVer.Append(vaCode)
    End If
End Sub

Sub addCodeList()
    If vaAndroidVer.Bound <> -1 Then
        Call setSdkPathDirectoryIds()
        Call addListUL()
        Dim i, j, category

        For i = 0 To vaAndroidVer.Bound
            category = vaAndroidVer.V(i).Name
            Call setSdkPathDirectoryIds()
            Call addListDirectoryLi(category, getSdkPathDivId() & LCase(category))

            Call setSdkPathListIds(category)
            Call addListUL()

            if vaAndroidVer.V(i).Bound <> -1 Then
                For j = 0 To vaAndroidVer.V(i).Bound
                    Call addListLi(vaAndroidVer.V(i).V(j))
                Next
            End If
        Next
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


