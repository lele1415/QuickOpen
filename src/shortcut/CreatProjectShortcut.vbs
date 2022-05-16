Const SHORTCUT_STATE_HIDE = 0
Const SHORTCUT_STATE_SHOW = 1
Const SHORTCUT_TEXT_HIDE = "    Hide    "
Const SHORTCUT_TEXT_SHOW = "    Select    "
Const ID_CREATE_SHORTCUTS = "create_shortcuts"
Const ID_SHOW_OR_HIDE_SHORTCUTS = "show_or_hide_shortcuts"

Dim mShortcutState
mShortcutState = SHORTCUT_STATE_HIDE

Sub creatShortcut()
	If Trim(mIp.Infos.Work) = "" _
			Or Trim(mIp.Infos.Sdk) = "" _
			Or Trim(mIp.Infos.Product) = "" _
			Or Trim(mIp.Infos.Project) = "" Then
		MsgBox("work info is not complete!")
		Exit Sub
	End If

	Call saveWorkToArray()
	If mShortcutState = SHORTCUT_STATE_SHOW Then Call updateAllShortcuts()
	Call updateWorkInfoTxt()
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

Sub updateAllShortcuts()
	Call parentNode_removeAllChilds(ID_DIV_SHORTCUT)
	Call AddShortcut()
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
        Call addShortcutButton(obj.Work, obj.Sdk, obj.Product, obj.Project, ID_DIV_SHORTCUT)
    Next
End Sub

Sub removeShortcut(shortcutId)
	Dim i, obj, value
    For i = 0 To vaWorksInfo.Bound
        Set obj = vaWorksInfo.V(i)
        value = obj.Work + "/" + obj.Sdk + "/" + obj.Product + "/" + obj.Project + "_shortcut"
        If value = shortcutId Then
        	Call vaWorksInfo.PopBySeq(i)
        	Exit For
        End If
    Next
    Call updateWorkInfoTxt()
End Sub

Sub applyShortcut(work, sdk, product, project)
	Call hideAllShortcuts()
	
	mIp.Work = work
	mIp.Sdk = sdk
	mIp.Product = product
	mIp.Project = project
	Call updateProductList()
End Sub

Sub saveWorkToArray()
	Dim i, oInfos
	For i = vaWorksInfo.Bound To 0 Step -1
		Set oInfos = vaWorksInfo.V(i)
		If oInfos.Work = mIp.Infos.Work Then
			oInfos.Sdk = mIp.Infos.Sdk
			oInfos.Product = mIp.Infos.Product
			oInfos.Project = mIp.Infos.Project
			Exit Sub
		ElseIf oInfos.Sdk = mIp.Infos.Sdk And _
	    	    oInfos.Product = mIp.Infos.Product And _
	    	    oInfos.Project = mIp.Infos.Project Then
	    	oInfos.Work = mIp.Infos.Work
	    	Exit Sub
		End If
	Next

	Set oInfos = New ProjectInfos
	oInfos.Work = mIp.Infos.Work
	oInfos.Sdk = mIp.Infos.Sdk
	oInfos.Product = mIp.Infos.Product
	oInfos.Project = mIp.Infos.Project

    vaWorksInfo.Append(oInfos)
End Sub

Sub updateWorkInfoTxt()
    initTxtFile(pProjectText)
    Dim oTxt, i, obj
    Set oTxt = oFso.OpenTextFile(pProjectText, FOR_APPENDING, False, True)

    For i = 0 To vaWorksInfo.Bound
        Set obj = vaWorksInfo.V(i)
        oTxt.WriteLine("########")
        oTxt.WriteLine(obj.Work)
        oTxt.WriteLine(obj.Sdk)
        oTxt.WriteLine(obj.Product)
        oTxt.WriteLine(obj.Project)
        oTxt.WriteLine()
    Next

    oTxt.Close
    Set oTxt = Nothing
End Sub

Sub updateNewShortcutBtn()
	Set obj = vaWorksInfo.V(vaWorksInfo.Bound)
	Call addShortcutButton(obj.Work, obj.Sdk, obj.Product, obj.Project, ID_DIV_SHORTCUT)
End Sub

Sub upShortcut(sName)
	Dim i, oInfos
	For i = 0 To vaWorksInfo.Bound

		Set oInfos = vaWorksInfo.V(i)
		If oInfos.Work = sName Then
			vaWorksInfo.MoveToEnd(i)
			Set oInfos = Nothing
			Exit For
		End If
		Set oInfos = Nothing
	Next
	Call updateAllShortcuts()
	Call updateWorkInfoTxt()
End Sub
