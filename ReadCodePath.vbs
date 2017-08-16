Const ID_INPUT_CODE_PATH = "input_code_path"
Const ID_LIST_CODE_PATH = "list_code_path"
Const ID_UL_CODE_PATH = "ul_code_path"

Const ID_LIST_CODE_PATH_M0 = "list_code_path_m0"
Const ID_UL_CODE_PATH_M0 = "ul_code_path_m0"

Const ID_LIST_CODE_PATH_N0 = "list_code_path_n0"
Const ID_UL_CODE_PATH_N0 = "ul_code_path_n0"

Dim pCodePathTxt : pCodePathTxt = oWs.CurrentDirectory & "\codePath.txt"
Dim pCodePathTxt_M0 : pCodePathTxt_M0 = oWs.CurrentDirectory & "\codePath_M0.txt"
Dim pCodePathTxt_N0 : pCodePathTxt_N0 = oWs.CurrentDirectory & "\codePath_N0.txt"

Call readCodePath(pCodePathTxt, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH, ID_UL_CODE_PATH)
Call readCodePath(pCodePathTxt_M0, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_M0, ID_UL_CODE_PATH_M0)
Call readCodePath(pCodePathTxt_N0, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_N0, ID_UL_CODE_PATH_N0)

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

    Select Case inputId
        Case ID_INPUT_CODE_PATH
            Call onloadPrjAndOpt()
        Case ID_INPUT_PROJECT
            Call onloadOpt()
    End Select
End Sub
