Const SHORTCUT_STATE_HIDE = 0
Const SHORTCUT_STATE_SHOW = 1
Const SHORTCUT_TEXT_HIDE = "Hide"
Const SHORTCUT_TEXT_SHOW = "Show"
Const ID_CREATE_SHORTCUTS = "create_shortcuts"
Const ID_SHOW_OR_HIDE_SHORTCUTS = "show_or_hide_shortcuts"

Dim mShortcutState
mShortcutState = SHORTCUT_STATE_HIDE

Sub creatShortcut()
	Dim sWorkName, sWorkCode, sWorkTargetProduct, sWorkCustomProject
	sWorkName = getElementValue(ID_WORK_NAME)
	sWorkCode = getSdkPath()
	sWorkTargetProduct = getProduct()
	sWorkCustomProject = getProject()

	If Trim(sWorkName) = "" _
			Or Trim(sWorkCode) = "" _
			Or Trim(sWorkTargetProduct) = "" _
			Or Trim(sWorkCustomProject) = "" Then
		MsgBox("work info is not complete!")
		Exit Sub
	End If

	Call saveWorkToArray(sWorkName, sWorkCode, sWorkTargetProduct, sWorkCustomProject)
	Call updateWorkInfoTxt()
	If mShortcutState = SHORTCUT_STATE_SHOW Then Call updateShortcutBtn()
End Sub

Sub showAllShortcuts()
	If parentNode_getChildNodesLength(ID_DIV_SHORTCUT) = 0 Then
		Call AddShortcut()
	End If
	mShortcutState = SHORTCUT_STATE_SHOW
	Call setElementValue(ID_SHOW_OR_HIDE_SHORTCUTS, SHORTCUT_TEXT_HIDE)
End Sub

Sub hideAllShortcuts()
	If parentNode_getChildNodesLength(ID_DIV_SHORTCUT) > 0 Then
		Call parentNode_removeAllChilds(ID_DIV_SHORTCUT)
	End If
	mShortcutState = SHORTCUT_STATE_HIDE
	Call setElementValue(ID_SHOW_OR_HIDE_SHORTCUTS, SHORTCUT_TEXT_SHOW)
End Sub

Sub showOrHideAllShortcuts()
	If mShortcutState = SHORTCUT_STATE_HIDE Then
		Call showAllShortcuts()
	ElseIf mShortcutState = SHORTCUT_STATE_SHOW Then
		Call hideAllShortcuts()
	End If
End Sub

Sub AddShortcut()
    Dim i, obj
    For i = vaWorksInfo.Bound To 0 Step -1
        Set obj = vaWorksInfo.V(i)
        Call addShortcutButton(obj.WorkName, obj.WorkCode, obj.WorkPrj, obj.WorkOpt, ID_DIV_SHORTCUT)
    Next
End Sub

Sub removeShortcut(shortcutId)
	Dim i, obj, value
    For i = 0 To vaWorksInfo.Bound
        Set obj = vaWorksInfo.V(i)
        value = obj.WorkName + "/" + obj.WorkCode + "/" + obj.WorkPrj + "/" + obj.WorkOpt + "_shortcut"
        If value = shortcutId Then
        	Call vaWorksInfo.PopBySeq(i)
        	Exit For
        End If
    Next
    Call updateWorkInfoTxt()
End Sub

Sub applyShortcut(sWorkName, sWorkCode, sWorkTargetProduct, sWorkCustomProject)
	Call hideAllShortcuts()
	
	Call setElementValue(ID_WORK_NAME, sWorkName)

	Call setSdkPath(sWorkCode)
	Call setProduct(sWorkTargetProduct)
	Call setProject(sWorkCustomProject)
End Sub

Sub saveWorkToArray(sWorkName, sWorkCode, sWorkTargetProduct, sWorkCustomProject)
	Dim oWork
	Set oWork = New WorkInfo
	oWork.WorkName = sWorkName
	oWork.WorkCode = sWorkCode
	oWork.WorkPrj = sWorkTargetProduct
	oWork.WorkOpt = sWorkCustomProject

	vaWorksInfo.Append(oWork)
End Sub

Sub updateWorkInfoTxt()
    initTxtFile(pWorkInfoText)
    Dim oTxt, i, obj
    Set oTxt = oFso.OpenTextFile(pWorkInfoText, FOR_APPENDING, False, True)

    For i = 0 To vaWorksInfo.Bound
        Set obj = vaWorksInfo.V(i)
        oTxt.WriteLine("########")
        oTxt.WriteLine(obj.WorkName)
        oTxt.WriteLine(obj.WorkCode)
        oTxt.WriteLine(obj.WorkPrj)
        oTxt.WriteLine(obj.WorkOpt)
        oTxt.WriteLine()
    Next

    oTxt.Close
    Set oTxt = Nothing
End Sub

Sub updateShortcutBtn()
	Set obj = vaWorksInfo.V(vaWorksInfo.Bound)
	Call addShortcutButton(obj.WorkName, obj.WorkCode, obj.WorkPrj, obj.WorkOpt, ID_DIV_SHORTCUT)
End Sub
