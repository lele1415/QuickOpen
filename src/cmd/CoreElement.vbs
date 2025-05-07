Option Explicit

Const ID_INPUT_CMD = "input_cmd"

Const ID_PARENT_OPEN_PATH = "parent_open_path"
Const ID_INPUT_OPEN_PATH = "input_open_path"

Const ID_DIV_OPEN_PATH_DIRECTORY = "div_open_path_directory"
Const ID_UL_OPEN_PATH_DIRECTORY = "ul_open_path_directory"

Const ID_DIV_OPEN_PATH_ = "div_open_path_"
Const ID_UL_OPEN_PATH_ = "ul_open_path_"

Const ID_PARENT_OUT_BUTTON = "parent_out_button"

Const ID_PARENT_OPEN_BUTTON = "parent_open_button"

Const ID_PARENT_FILE_BUTTON = "parent_file_button"

Const ID_PARENT_FIND_PROJECT_BUTTON = "parent_find_project_button"

Dim mCmdInput : Set mCmdInput = (New InputText)(getCmdInputId())
Dim mOpenPathInput : Set mOpenPathInput = (New InputText)(getOpenPathInputId())
Dim mOpenButtonList : Set mOpenButtonList = (New ButtonWithOneLayerList)(getOpenButtonParentId(), "openbutton")
Dim mFileButtonList : Set mFileButtonList = (New ButtonWithOneLayerList)(getFileButtonParentId(), "filebutton")
Dim mOverlayPathDict : Set mOverlayPathDict = CreateObject("Scripting.Dictionary")



'Open path
Function getCmdInputId()
    getCmdInputId = ID_INPUT_CMD
End Function

Function getOpenPathParentId()
    getOpenPathParentId = ID_PARENT_OPEN_PATH
End Function

Function getOpenPathInputId()
    getOpenPathInputId = ID_INPUT_OPEN_PATH
End Function

Function getOpenPathDirectoryDivId()
    getOpenPathDirectoryDivId = ID_DIV_OPEN_PATH_DIRECTORY
End Function

Function getOpenPathDirectoryULId()
    getOpenPathDirectoryULId = ID_UL_OPEN_PATH_DIRECTORY
End Function

Function getOpenPathDivId()
    getOpenPathDivId = ID_DIV_OPEN_PATH_
End Function

Function getOpenPathULId()
    getOpenPathULId = ID_UL_OPEN_PATH_
End Function

Function getOutButtonParentId()
    getOutButtonParentId = ID_PARENT_OUT_BUTTON
End Function

Function getOpenButtonParentId()
    getOpenButtonParentId = ID_PARENT_OPEN_BUTTON
End Function

Function getFileButtonParentId()
    getFileButtonParentId = ID_PARENT_FILE_BUTTON
End Function

Function getFindProjectButtonParentId()
    getFindProjectButtonParentId = ID_PARENT_FIND_PROJECT_BUTTON
End Function

Function getElementValue(elementId)
    getElementValue = document.getElementById(elementId).value
End Function

Sub setElementValue(elementId, str)
    document.getElementById(elementId).value = str
End Sub

Sub disableElement(inputId)
    document.getElementById(inputId).disabled = "disabled"
End Sub

Sub enableElement(inputId)
    document.getElementById(inputId).disabled = ""
End Sub

Sub hideElement(elementId)
    document.getElementById(elementId).style.display = "none"
End Sub

Sub showElement(elementId)
    document.getElementById(elementId).style.display = "block"
End Sub

Sub freezeAllInput()
    Call disableElement(getWorkInputId())
    Call disableElement(getSdkPathInputId())
    Call disableElement(getProductInputId())
    Call disableElement(getProjectInputId())
    Call disableElement(ID_CREATE_SHORTCUTS)
    Call disableElement(ID_SHOW_SHORTCUTS)
    Call disableElement(ID_HIDE_SHORTCUTS)
End Sub

Sub unfreezeAllInput()
    Call enableElement(getWorkInputId())
    Call enableElement(getSdkPathInputId())
    Call enableElement(getProductInputId())
    Call enableElement(getProjectInputId())
    Call enableElement(ID_CREATE_SHORTCUTS)
    Call enableElement(ID_SHOW_SHORTCUTS)
    Call enableElement(ID_HIDE_SHORTCUTS)
End Sub

Function getCmdText()
    getCmdText = mCmdInput.text
End Function

Sub setCmdText(path)
    mCmdInput.setText(path)
End Sub

Function getOpenPath()
    getOpenPath = mOpenPathInput.text
End Function

Sub setOpenPath(path)
    mOpenPathInput.setText(path)
End Sub

Function getOpenButtonListPath(where)
    getOpenButtonListPath = mOverlayPathDict.Item(where)
End Function

Sub onInputListClick(divId, str)
    Dim path
    If InStr(divId, "openbutton") > 0 Then
        path = getOpenButtonListPath(str)
        If path <> "" Then runPath(path)

    ElseIf InStr(divId, "filebutton") > 0 Then
        Call mFileButtonList.removeList()
        Call setOpenPath(str)
    End If
End Sub

Function changeListFocus(keyCode)
    If mFileButtonList.changeFocus(keyCode) Then changeListFocus = True : Exit Function
    If mOpenButtonList.changeFocus(keyCode) Then changeListFocus = True : Exit Function
    changeListFocus = False
End Function

Sub onOpenButtonClick()
    If mCmdInput.text <> "" Then
        Dim cmd : cmd = mCmdInput.text
        Call handleCmdInput()
        If mCmdInput.text = "" Then Call saveHistoryCmd(cmd)
        Exit Sub
    End If
    Call removeOpenButtonList()
    Call makeOpenButton()
    If mOpenButtonList.VaArray.Bound = -1 Then
        Call runPath(getOpenPath())
    Else
        Call mOpenButtonList.toggleButtonList()
    End If
End Sub

Sub removeOpenButtonList()
    Call mOverlayPathDict.RemoveAll()
    Call mOpenButtonList.removeList()
End Sub

Sub makeOpenButton()
    Dim inputPath
    inputPath = getOpenPath()
    If Trim(inputPath) = "" Or _
            InStr(inputPath, ":\") > 0 Or _
            InStr(inputPath, "\\") = 1 Or _
            InStr(inputPath, "weibu/") = 1 Or _
            InStr(inputPath, "out/") = 1 Then
        Exit Sub
    End If

    Call findOverlayPath()

    If mOpenButtonList.VaArray.Bound > -1 Then
        Call mOverlayPathDict.Add("Origin", inputPath)
        Call mOpenButtonList.VaArray.append("Origin")
        Call mOpenButtonList.addList()
    End If
End Sub

Sub findOverlayPath()
    Dim inputPath, isFile, projectName, newName, wholePath, configFilePath
    inputPath = getOpenPath()
    If isFileExists(inputPath) Then
        isFile = True
    ElseIf isFolderExists(inputPath) Then
        isFile = False
    Else
        Exit Sub
    End If

    projectName = mBuild.Project
    wholePath = mBuild.Infos.getOverlayPath(inputPath)
    If isFile Then configFilePath = mBuild.Infos.ProjectPath & "/config/" & getFileNameFromPath(inputPath)

    Do
        If isFile And (Not isFileExists(wholePath)) Then
            If isFileExists(configFilePath) Then
                wholePath = configFilePath
            End If
        End If

        If isFile Then
            If isFileExists(wholePath) Then
                Call mOverlayPathDict.Add(projectName, wholePath)
                Call mOpenButtonList.VaArray.append(projectName)
            ElseIf isFileExists(configFilePath) Then
                Call mOverlayPathDict.Add(projectName, configFilePath)
                Call mOpenButtonList.VaArray.append(projectName)
            End If
                
        ElseIf isFolderExists(wholePath) Then
            Call mOverlayPathDict.Add(projectName, wholePath)
            Call mOpenButtonList.VaArray.append(projectName)
        End If

        If InStr(projectName, "-") > 0 Then
            newName = Left(projectName, InStrRev(projectName, "-") - 1)
            wholePath = Replace(wholePath, projectName, newName)
            If isFile Then configFilePath = Replace(configFilePath, projectName, newName)
            projectName = newName
        Else
            Exit Do
        End If
    Loop
End Sub

Sub onOpenPathChange()
    Call replaceOpenPath()
End Sub

Sub replaceOpenPath()
    Dim path : path = getOpenPath()
    If InStr(path, ":\") > 0 Or InStr(path, "\\192.168") > 0 Then Exit Sub

    path = replaceProjectInfoStr(path)
    Call setOpenPath(relpaceSlashInPath(path))

    'Call cutSdkPath()
End Sub

Sub pasteAndOpenPath()
    Call setElementValue(ID_INPUT_OPEN_PATH, "")
    Call focusElement(ID_INPUT_OPEN_PATH)
    oWs.SendKeys "^v"
    oWs.SendKeys "{ENTER}"
End Sub

Const KEYCODE_ENTER = 13
Const KEYCODE_SPACE = 32
Const KEYCODE_TAB = 9
Const KEYCODE_UP = 38
Const KEYCODE_DOWN = 40
Const KEYCODE_ESC = 27
Function onKeyDown(keyCode)
    If keyCode = KEYCODE_ENTER Then
        If mFileButtonList.isShowing() Then
            Call mFileButtonList.clickFocusedLi()
        ElseIf mOpenButtonList.isShowing() Then
            Call mOpenButtonList.clickFocusedLi()
        Else
            Call onOpenButtonClick()
        End If

    ElseIf keyCode = KEYCODE_ESC Then
        Call mFileButtonList.hideListIfShowing()
        Call mOpenButtonList.hideListIfShowing()

    ElseIf keyCode = KEYCODE_SPACE Then
        Call pasteAndOpenPath()
    
    ElseIf keyCode = KEYCODE_TAB Then
        Call setOpenPath(getOpenPath() & getTabStr())
    
    ElseIf keyCode = KEYCODE_UP Or keyCode = KEYCODE_DOWN Then
        If changeListFocus(keyCode) Then onKeyDown = False : Exit Function
        If document.activeElement.id = ID_INPUT_OPEN_PATH Then
            Call showHistoryPath(keyCode)
        ElseIf document.activeElement.id = ID_INPUT_CMD Then
            Call showHistoryCmd(keyCode)
        End If

    Else
        onKeyDown = True : Exit Function
    End If
    
    onKeyDown = False
End Function

Sub updateTitle(titleStr)
    document.title = titleStr
End Sub
