Const ID_INPUT_CODE_PATH = "input_code_path"

Const ID_SELECT_CODE_PATH = "select_code_path"

Const ID_LIST_CODE_PATH_SELECT_VER = "list_code_path_select_ver"
Const ID_UL_CODE_PATH_SELECT_VER = "ul_code_path_select_ver"

Const ID_LIST_CODE_PATH_L1 = "list_code_path_l1"
Const ID_UL_CODE_PATH_L1 = "ul_code_path_l1"

Const ID_LIST_CODE_PATH_M0 = "list_code_path_m0"
Const ID_UL_CODE_PATH_M0 = "ul_code_path_m0"

Const ID_LIST_CODE_PATH_N0 = "list_code_path_n0"
Const ID_UL_CODE_PATH_N0 = "ul_code_path_n0"

Const ANDROID_VERSION_N0 = "N0"
Const ANDROID_VERSION_M0 = "M0"
Const ANDROID_VERSION_L1 = "L1"

Const VALUE_SELECT_CODE_PATH_SHOW = "选择代码"
Const VALUE_SELECT_CODE_PATH_HIDE = "收起"

Dim pConfigText : pConfigText = oWs.CurrentDirectory & "\config.ini"

Dim mTextEditorPath

Dim vaCodePath_N0 : Set vaCodePath_N0 = New VariableArray
Dim vaCodePath_M0 : Set vaCodePath_M0 = New VariableArray
Dim vaCodePath_L1 : Set vaCodePath_L1 = New VariableArray

Call addVerForSelect()
Call readConfigText(pConfigText)
call addLiOfCodePath(vaCodePath_N0, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_N0, ID_UL_CODE_PATH_N0)
call addLiOfCodePath(vaCodePath_M0, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_M0, ID_UL_CODE_PATH_M0)
call addLiOfCodePath(vaCodePath_L1, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_L1, ID_UL_CODE_PATH_L1)

Sub addVerForSelect()
    Call addAfterLiForCodePath("N0", ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
    Call addAfterLiForCodePath("M0", ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
    Call addAfterLiForCodePath("L1", ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
End Sub

Sub selectCodePathOnClick()
    Dim value
    value = getElementValue(ID_SELECT_CODE_PATH)

    If value = VALUE_SELECT_CODE_PATH_SHOW Then
        Call setElementValue(ID_SELECT_CODE_PATH, VALUE_SELECT_CODE_PATH_HIDE)
        Call showOrHideCodePathList(ID_LIST_CODE_PATH_SELECT_VER, "show")
    ElseIf value = VALUE_SELECT_CODE_PATH_HIDE Then
        Call setElementValue(ID_SELECT_CODE_PATH, VALUE_SELECT_CODE_PATH_SHOW)
        Call HideCodePathList()
    End If
End Sub

Sub readConfigText(DictPath)
    If Not oFso.FileExists(DictPath) Then Exit Sub
    
    Dim oText, sReadLine, sAndroidVer
    sAndroidVer = ""
    Set oText = oFso.OpenTextFile(DictPath, FOR_READING)

    Do Until oText.AtEndOfStream
        Call handleReadLine(oText, sReadLine, sAndroidVer)
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub handleReadLine(oText, sReadLine, sAndroidVer)
    sReadLine = oText.ReadLine

    If InStr(sReadLine, "TextEditorPath") > 0 Then
        mTextEditorPath = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
        Exit Sub
    End If

    If sAndroidVer = "" Then
        Select Case Trim(sReadLine)
            Case "N0 {"
                sAndroidVer = ANDROID_VERSION_N0
                sReadLine = oText.ReadLine
            Case "M0 {"
                sAndroidVer = ANDROID_VERSION_M0
                sReadLine = oText.ReadLine
            Case "L1 {"
                sAndroidVer = ANDROID_VERSION_L1
                sReadLine = oText.ReadLine
        End Select
    End If

    If sReadLine = "}" Then sAndroidVer = ""

    If sAndroidVer <> "" Then
        Select Case sAndroidVer
            Case ANDROID_VERSION_N0
                Call vaCodePath_N0.Append(sReadLine)
            Case ANDROID_VERSION_M0
                Call vaCodePath_M0.Append(sReadLine)
            Case ANDROID_VERSION_L1
                Call vaCodePath_L1.Append(sReadLine)
        End Select
    End If
End Sub

Sub addLiOfCodePath(vaObj, inputId, listId, ulId)
    If vaObj.Bound <> -1 Then
        Dim i
        For i = 0 To vaObj.Bound
            Call addAfterLiForCodePath(vaObj.V(i), inputId, listId, ulId)
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

Sub setListValueForCodePath(inputId, listId, value)
    Call showOrHideCodePathList(listId, "hide")

    If listId = ID_LIST_CODE_PATH_SELECT_VER Then
        Call showOrHideCodePathList(Eval("ID_LIST_CODE_PATH_" & value), "show")
    Else
        Call setElementValue(inputId, value)
        Call setElementValue(ID_SELECT_CODE_PATH, VALUE_SELECT_CODE_PATH_SHOW)
        Call onloadPrjAndOpt()
    End If
End Sub
