Const ID_INPUT_CODE_PATH_KK = "input_code_path_kk"
Const ID_LIST_CODE_PATH_KK = "list_code_path_kk"
Const ID_UL_CODE_PATH_KK = "ul_code_path_kk"
Const ID_INPUT_CODE_PATH_L1 = "input_code_path_l1"
Const ID_LIST_CODE_PATH_L1 = "list_code_path_l1"
Const ID_UL_CODE_PATH_L1 = "ul_code_path_l1"

Dim pCodePathTxt_KK : pCodePathTxt_KK = oWs.CurrentDirectory & "\codePath_KK.txt"
Dim pCodePathTxt_L1 : pCodePathTxt_L1 = oWs.CurrentDirectory & "\codePath_L1.txt"
Dim pID_LIST_CODE_PATH_KK : pID_LIST_CODE_PATH_KK = oWs.CurrentDirectory & "\codePath_KK" 
Dim pID_LIST_CODE_PATH_L1 : pID_LIST_CODE_PATH_L1 = oWs.CurrentDirectory & "\codePath_L1"

Call readCodePath(pCodePathTxt_KK, ID_INPUT_CODE_PATH_KK, ID_LIST_CODE_PATH_KK, ID_UL_CODE_PATH_KK)
Call readCodePath(pCodePathTxt_L1, ID_INPUT_CODE_PATH_L1, ID_LIST_CODE_PATH_L1, ID_UL_CODE_PATH_L1)

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
        Case ID_INPUT_CODE_PATH_L1
            Call onloadPrjAndOpt()
        Case ID_INPUT_PROJECT_L1
            Call onloadOpt()
    End Select
End Sub

Function getElementValue(elementId)
    getElementValue = document.getElementById(elementId).value
End Function

Sub setElementValue(elementId, str)
    document.getElementById(elementId).value = str
End Sub