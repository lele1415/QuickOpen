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

Const ID_LIST_CODE_PATH_O1 = "list_code_path_o1"
Const ID_UL_CODE_PATH_O1 = "ul_code_path_o1"

Const ANDROID_VERSION_O1 = "O1"
Const ANDROID_VERSION_N0 = "N0"
Const ANDROID_VERSION_M0 = "M0"
Const ANDROID_VERSION_L1 = "L1"

Const VALUE_SELECT_CODE_PATH_SHOW = "选择代码"
Const VALUE_SELECT_CODE_PATH_HIDE = "收起"

Dim pConfigText : pConfigText = oWs.CurrentDirectory & "\config.ini"

Dim mTextEditorPath

Dim mCodeExist_O1 : mCodeExist_O1 = False
Dim mCodeExist_N0 : mCodeExist_N0 = False
Dim mCodeExist_M0 : mCodeExist_M0 = False
Dim mCodeExist_L1 : mCodeExist_L1 = False

Dim vaCodePath_O1 : Set vaCodePath_O1 = New VariableArray
Dim vaCodePath_N0 : Set vaCodePath_N0 = New VariableArray
Dim vaCodePath_M0 : Set vaCodePath_M0 = New VariableArray
Dim vaCodePath_L1 : Set vaCodePath_L1 = New VariableArray

Call readConfigText(pConfigText)
Call addVerForSelect()
call addLiOfCodePath(vaCodePath_O1, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_O1, ID_UL_CODE_PATH_O1)
call addLiOfCodePath(vaCodePath_N0, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_N0, ID_UL_CODE_PATH_N0)
call addLiOfCodePath(vaCodePath_M0, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_M0, ID_UL_CODE_PATH_M0)
call addLiOfCodePath(vaCodePath_L1, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_L1, ID_UL_CODE_PATH_L1)

Sub addVerForSelect()
    If mCodeExist_O1 Then _
        Call addAfterLiForCodePath(ANDROID_VERSION_O1, _
                ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
    If mCodeExist_N0 Then _
        Call addAfterLiForCodePath(ANDROID_VERSION_N0, _
                ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
    If mCodeExist_M0 Then _
        Call addAfterLiForCodePath(ANDROID_VERSION_M0, _
                ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
    If mCodeExist_L1 Then _
        Call addAfterLiForCodePath(ANDROID_VERSION_L1, _
                ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
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
    
    Dim oText, sAndroidVer
    sAndroidVer = ""
    Set oText = oFso.OpenTextFile(DictPath, FOR_READING)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        Call handleForConfig(oText, sReadLine, sAndroidVer)
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub handleForConfig(oText, sReadLine, sAndroidVer)
    If InStr(sReadLine, "TextEditorPath") > 0 Then
        mTextEditorPath = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
        Exit Sub
    End If

    If sAndroidVer = "" Then
        Select Case Trim(sReadLine)
            Case "O1 {"
                mCodeExist_O1 = True
                sAndroidVer = ANDROID_VERSION_O1
                sReadLine = oText.ReadLine
            Case "N0 {"
                mCodeExist_N0 = True
                sAndroidVer = ANDROID_VERSION_N0
                sReadLine = oText.ReadLine
            Case "M0 {"
                mCodeExist_M0 = True
                sAndroidVer = ANDROID_VERSION_M0
                sReadLine = oText.ReadLine
            Case "L1 {"
                mCodeExist_L1 = True
                sAndroidVer = ANDROID_VERSION_L1
                sReadLine = oText.ReadLine
        End Select
    End If

    If sReadLine = "}" Then sAndroidVer = ""

    If sAndroidVer <> "" Then
        Select Case sAndroidVer
            Case ANDROID_VERSION_O1
                Call vaCodePath_O1.Append(sReadLine)
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
