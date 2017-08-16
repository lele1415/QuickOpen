Const ID_INPUT_CODE_PATH = "input_code_path"

Const ID_LIST_CODE_PATH_SELECT_VER = "list_code_path_select_ver"
Const ID_UL_CODE_PATH_SELECT_VER = "ul_code_path_select_ver"

Const ID_LIST_CODE_PATH_L1 = "list_code_path_l1"
Const ID_UL_CODE_PATH_L1 = "ul_code_path_l1"

Const ID_LIST_CODE_PATH_M0 = "list_code_path_m0"
Const ID_UL_CODE_PATH_M0 = "ul_code_path_m0"

Const ID_LIST_CODE_PATH_N0 = "list_code_path_n0"
Const ID_UL_CODE_PATH_N0 = "ul_code_path_n0"

Dim pCodePathTxt_L1 : pCodePathTxt_L1 = oWs.CurrentDirectory & "\codePath_L1.txt"
Dim pCodePathTxt_M0 : pCodePathTxt_M0 = oWs.CurrentDirectory & "\codePath_M0.txt"
Dim pCodePathTxt_N0 : pCodePathTxt_N0 = oWs.CurrentDirectory & "\codePath_N0.txt"

Call addVerForSelect()
Call readCodePath(pCodePathTxt_L1, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_L1, ID_UL_CODE_PATH_L1)
Call readCodePath(pCodePathTxt_M0, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_M0, ID_UL_CODE_PATH_M0)
Call readCodePath(pCodePathTxt_N0, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_N0, ID_UL_CODE_PATH_N0)

Sub addVerForSelect()
    Call addAfterLi("N0", ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
    Call addAfterLi("M0", ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
    Call addAfterLi("L1", ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
End Sub

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

    If listId = ID_LIST_CODE_PATH_SELECT_VER Then
        Call showAndHide(Eval("ID_LIST_CODE_PATH_" & value), "show")
    Else
        Call setElementValue(inputId, value)
        
        Select Case inputId
            Case ID_INPUT_CODE_PATH
                Call onloadPrjAndOpt()
            Case ID_INPUT_PROJECT
                Call onloadOpt()
        End Select
    End If

End Sub
