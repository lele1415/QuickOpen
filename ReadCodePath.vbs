Const ID_SELECT_CODE_PATH = "select_code_path"
Const ID_INPUT_CODE_PATH = "input_code_path"
Const ID_LIST_CODE_PATH = "list_code_path"
Const ID_UL_CODE_PATH = "ul_code_path"

Const VALUE_SELECT_CODE_PATH_SHOW = "选择代码"
Const VALUE_SELECT_CODE_PATH_HIDE = "收起"

Dim pCodePathTxt : pCodePathTxt = oWs.CurrentDirectory & "\codePath.txt"

Call readCodePath(pCodePathTxt, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH, ID_UL_CODE_PATH)

Sub selectCodePathOnClick()
    Dim value
    value = getElementValue(ID_SELECT_CODE_PATH)

    If value = VALUE_SELECT_CODE_PATH_SHOW Then
        Call setElementValue(ID_SELECT_CODE_PATH, VALUE_SELECT_CODE_PATH_HIDE)
        Call showOrHideCodePathList(ID_LIST_CODE_PATH, "show")
    ElseIf value = VALUE_SELECT_CODE_PATH_HIDE Then
        Call setElementValue(ID_SELECT_CODE_PATH, VALUE_SELECT_CODE_PATH_SHOW)
        Call HideCodePathList()
    End If
End Sub

Sub readCodePath(DictPath, inputId, listId, ulId)
    Call readTextAndDoSomething(DictPath, _
            "If Len(Trim(sReadLine)) > 0 Then" &_
            " Call addAfterLiForCodePath(sReadLine, """&inputId&""", """&listId&""", """&ulId&""")")
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

Sub setListValueForCodePath(inputId, listId, value)
    Call showOrHideCodePathList(listId, "hide")

    Call setElementValue(inputId, value)
    Call setElementValue(ID_SELECT_CODE_PATH, VALUE_SELECT_CODE_PATH_SHOW)
End Sub
