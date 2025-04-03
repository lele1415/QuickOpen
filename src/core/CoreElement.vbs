Option Explicit

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
    document.title = mIp.Work & "\weibu\" & mIp.Infos.Product & " " & mDrive
End Sub
