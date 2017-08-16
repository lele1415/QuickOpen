Const ID_INPUT_CODE_PATH = "input_code_path"
Const ID_LIST_CODE_PATH = "list_code_path"
Const ID_UL_CODE_PATH = "ul_code_path"

Dim pCodePathTxt : pCodePathTxt = oWs.CurrentDirectory & "\codePath.txt"

Call readCodePath(pCodePathTxt, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH, ID_UL_CODE_PATH)

Sub readCodePath(DictPath, inputId, listId, ulId)
    Call readTextAndDoSomething(DictPath, _
            "If Len(Trim(sReadLine)) > 0 Then" &_
            " Call addAfterLi(sReadLine, """&inputId&""", """&listId&""", """&ulId&""")")
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

Sub setListValue(inputId, listId, value)
    Call showAndHide(listId, "hide")
    Call setElementValue(inputId, value)
End Sub
