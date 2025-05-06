Option Explicit

Const ID_PARENT_SDK_PATH = "parent_sdk_path"
Const ID_INPUT_SDK_PATH = "input_sdk_path"

Const ID_DIV_SDK_PATH_DIRECTORY = "div_sdk_path_directory"
Const ID_UL_SDK_PATH_DIRECTORY = "ul_sdk_path_directory"
Const ID_DIV_SDK_PATH_ = "div_sdk_path_"
Const ID_UL_SDK_PATH_ = "ul_sdk_path_"

Const ID_PARENT_PRODUCT = "parent_product"
Const ID_INPUT_PRODUCT = "input_product"
Const ID_DIV_PRODUCT = "list_target_product"
Const ID_UL_PRODUCT = "ul_target_product"

Const ID_PARENT_PROJECT = "parent_project"
Const ID_INPUT_PROJECT = "input_project"
Const ID_DIV_PROJECT = "list_custom_project"
Const ID_UL_PROJECT = "ul_custom_project"

Const ID_WORK_NAME = "work_name"
Const ID_DIV_SHORTCUT = "div_shortcut"
Const ID_INPUT_FIRMWARE = "input_firmware"
Const ID_INPUT_REQUIREMENTS = "input_requirements"
Const ID_INPUT_ZENTAO = "input_zentao"

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

Const ID_SELECT_FOR_COMPARE = "select_for_compare"
Const ID_COMPARE_TO = "compare_to"

Dim mIsCmdMode
mIsCmdMode = True



Function getWorkInputId()
    getWorkInputId = ID_WORK_NAME
End Function

Function getFirmwareInputId()
    getFirmwareInputId = ID_INPUT_FIRMWARE
End Function

Function getRequirementsInputId()
    getRequirementsInputId = ID_INPUT_REQUIREMENTS
End Function

Function getZentaoInputId()
    getZentaoInputId = ID_INPUT_ZENTAO
End Function

Function getSdkPathParentId()
    getSdkPathParentId = ID_PARENT_SDK_PATH
End Function

Function getSdkPathInputId()
    getSdkPathInputId = ID_INPUT_SDK_PATH
End Function

Function getSdkPathDirectoryDivId()
    getSdkPathDirectoryDivId = ID_DIV_SDK_PATH_DIRECTORY
End Function

Function getSdkPathDirectoryULId()
    getSdkPathDirectoryULId = ID_UL_SDK_PATH_DIRECTORY
End Function

Function getSdkPathDivId()
    getSdkPathDivId = ID_DIV_SDK_PATH_
End Function

Function getSdkPathULId()
    getSdkPathULId = ID_UL_SDK_PATH_
End Function

'Product
Function getProductParentId()
    getProductParentId = ID_PARENT_PRODUCT
End Function

Function getProductInputId()
    getProductInputId = ID_INPUT_PRODUCT
End Function

Function getProductDivId()
    getProductDivId = ID_DIV_PRODUCT
End Function

Function getProductULId()
    getProductULId = ID_UL_PRODUCT
End Function

'Project
Function getProjectParentId()
    getProjectParentId = ID_PARENT_PROJECT
End Function

Function getProjectInputId()
    getProjectInputId = ID_INPUT_PROJECT
End Function

Function getProjectDivId()
    getProjectDivId = ID_DIV_PROJECT
End Function

Function getProjectULId()
    getProjectULId = ID_UL_PROJECT
End Function

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
    If (Not mIsCmdMode) Or (mIsCmdMode And (elementId = getOpenPathInputId() Or elementId = getCmdInputId()))Then 
        document.getElementById(elementId).value = str
    End If
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

Sub startCmdMode()
    Call hideElement("work_fieldset")
    Call hideElement("project_fieldset")
    Call hideElement("openpath_fieldset")
    Call hideElement("openpath2_fieldset")
    Call hideElement("explorer_fieldset")
    Call hideElement("commands_fieldset")
    Call hideElement("br1")
    Call hideElement("br2")
    Call hideElement("br3")
    Call hideElement("br4")
    Call setCmdSmallWindow()
    Call setCmdTextClass(getCmdInputId(), getOpenPathInputId(), "cmd_text")
    mIsCmdMode = True
End Sub

Sub exitCmdMode()
    Call showElement("work_fieldset")
    Call showElement("project_fieldset")
    Call showElement("openpath_fieldset")
    Call showElement("openpath2_fieldset")
    Call showElement("explorer_fieldset")
    Call showElement("commands_fieldset")
    Call showElement("br1")
    Call showElement("br2")
    Call showElement("br3")
    Call showElement("br4")
    Call setDefaultWindow()
    Call setCmdTextClass(getCmdInputId(), getOpenPathInputId(), "textarea_text")
    mIsCmdMode = False
End Sub

Sub onInputListClick(divId, str)
    Dim path
    If InStr(divId, "outfile") > 0 Then
        path = getOutListPath(str)
        If path <> "" Then runPath(path)

    ElseIf InStr(divId, "openbutton") > 0 Then
        path = getOpenButtonListPath(str)
        If path <> "" Then runPath(path)

    ElseIf InStr(divId, "filebutton") > 0 Then
        Call mFileButtonList.removeList()
        Call setOpenPath(str)

    ElseIf InStr(divId, "findprojectbutton") > 0 Then
        Call mFindProjectButtonList.removeList()
        mIp.Project = str
    End If
End Sub

Function changeListFocus(keyCode)
    If mFileButtonList.changeFocus(keyCode) Then changeListFocus = True : Exit Function
    If mOpenButtonList.changeFocus(keyCode) Then changeListFocus = True : Exit Function
    If mFindProjectButtonList.changeFocus(keyCode) Then changeListFocus = True : Exit Function
    changeListFocus = False
End Function

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
        Call tabOpenPath()
    
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

Sub updateTitle()
    document.title = mIp.Infos.Work & "\weibu\" & mIp.Infos.Product & " " & mDrive
End Sub
