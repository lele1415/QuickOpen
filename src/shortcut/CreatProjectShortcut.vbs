Option Explicit

Const SHORTCUT_STATE_HIDE = 0
Const SHORTCUT_STATE_SHOW = 1
Const SHORTCUT_TEXT_HIDE = "     Hide     "
Const SHORTCUT_TEXT_SHOW = "    Select    "
Const ID_CREATE_SHORTCUTS = "create_shortcuts"
Const ID_SHOW_SHORTCUTS = "show_shortcuts"
Const ID_HIDE_SHORTCUTS = "hide_shortcuts"

Dim mShortcutState
mShortcutState = SHORTCUT_STATE_HIDE

Sub creatShortcut()
	If Not isWorkInfosValid(mIp.Infos) Then
		MsgBox("work info is not complete!")
		Exit Sub
	End If

	Call saveWorkToArray()
	If mShortcutState = SHORTCUT_STATE_SHOW Then Call updateAllShortcuts()
	Call updateWorkInfoTxt()
End Sub

Function isWorkInfosValid(infos)
    If Trim(infos.Work) = "" _
			Or Trim(infos.Sdk) = "" _
			Or Trim(infos.Product) = "" _
			Or (InStr(infos.Sdk, "_t0") > 0 And Trim(infos.SysSdk) = "") _
			Or (InStr(infos.Sdk, "_t0") > 0 And Trim(infos.SysProject) = "") Then
		isWorkInfosValid = False
	Else
	    isWorkInfosValid = True
	End If
End Function

Sub showAllShortcuts()
	If parentNode_getChildNodesLength(ID_DIV_SHORTCUT) = 0 Then
		Call AddShortcut()
	End If
	mShortcutState = SHORTCUT_STATE_SHOW
	Call hideElement(ID_SHOW_SHORTCUTS)
	Call showElement(ID_HIDE_SHORTCUTS)
End Sub

Sub hideAllShortcuts()
	If parentNode_getChildNodesLength(ID_DIV_SHORTCUT) > 0 Then
		Call parentNode_removeAllChilds(ID_DIV_SHORTCUT)
	End If
	mShortcutState = SHORTCUT_STATE_HIDE
	Call hideElement(ID_HIDE_SHORTCUTS)
	Call showElement(ID_SHOW_SHORTCUTS)
End Sub

Sub updateAllShortcuts()
	If parentNode_getChildNodesLength(ID_DIV_SHORTCUT) > 0 Then
		Call parentNode_removeAllChilds(ID_DIV_SHORTCUT)
		Call AddShortcut()
	End If
End Sub

Sub AddShortcut()
    Dim i, obj
    For i = vaWorksInfo.Bound To 0 Step -1
        Set obj = vaWorksInfo.V(i)
        Call addShortcutButton(obj.Work, obj.Sdk, obj.Product, obj.Project, obj.Firmware, obj.Requirements, obj.Zentao, ID_DIV_SHORTCUT)
    Next
End Sub

Sub removeShortcut(shortcutId)
	Dim i, obj, value, work
    For i = 0 To vaWorksInfo.Bound
        Set obj = vaWorksInfo.V(i)
        value = obj.Work + "/" + obj.Sdk + "/" + obj.Product + "/" + obj.Project + "_shortcut"
        If value = shortcutId Then
        	work = obj.Work
        	Call vaWorksInfo.PopBySeq(i)
        	Exit For
        End If
    Next
    If mIp.Work = work Then Call applyLastWorkInfo()
    Call updateWorkInfoTxt()
End Sub

Sub onShortcutButtonClick(work, sdk, product, project, firmware, requirements, zentao)
    Dim oInfos : Set oInfos = New ProjectInfos
	Call oInfos.setProjectAllInfos(work, sdk, product, project, firmware, requirements, zentao)
	Call hideAllShortcuts()
	Call applyShortcutInfos(oInfos)
End Sub

Sub applyShortcutInfos(infos)
	If Not checkProjectExist(infos.Sdk, infos.Product, infos.Project) Then Exit Sub
	Call mIp.setProjectInputs(infos)
	Call moveShortcutToTop(infos.Work)
End Sub

Sub saveWorkToArray()
	Dim i, oInfos
	For i = vaWorksInfo.Bound To 0 Step -1
		Set oInfos = vaWorksInfo.V(i)
		If oInfos.Work = mIp.Infos.Work Or oInfos.isSameProject(mIp.Infos) Then
		    Call oInfos.setProjectInfos(mIp.Infos)
			Exit Sub
		End If
	Next

	Set oInfos = New ProjectInfos
	Call oInfos.setProjectInfos(mIp.Infos)

    vaWorksInfo.Append(oInfos)
End Sub

Sub getProjectInfosFromOpenPath()
    Dim oInfos, inputArray, fullName, workInfosStr
	Set oInfos = New ProjectInfos
	inputArray = Split(getOpenPath(), VbLf)

	oInfos.Sdk = Split(inputArray(0), "/weibu/")(0)
	oInfos.Product = Split(Split(inputArray(0), "/weibu/")(1), "/")(0)
	oInfos.Project = Split(Split(inputArray(0), "/weibu/")(1), "/")(1)
	fullName = trimStr(Right(oInfos.Project, Len(oInfos.Project) - InStr(oInfos.Project, "_")))
	oInfos.Work = fullName & " " & oInfos.Sdk
	oInfos.Firmware = "\\192.168.0.248\安卓固件文件\"
	oInfos.Requirements = "\\192.168.0.24\wbshare\客户需求\"
	oInfos.Zentao = "http://192.168.0.29:3000/zentao/task-view-" & getTaskNum(oInfos.Project) & ".html"

	workInfosStr = oInfos.Work & VbLf &_
	               oInfos.Sdk & VbLf &_
	               oInfos.Product & VbLf &_
	               oInfos.Project & VbLf

	If UBound(inputArray) > 0 Then
	    oInfos.SysSdk = Split(inputArray(1), "/weibu/")(0)
        oInfos.SysProject = Split(Split(inputArray(1), "/weibu/")(1), "/")(1)
		workInfosStr = workInfosStr & oInfos.SysSdk & VbLf &_
	                                  oInfos.SysProject & VbLf
	End If

	workInfosStr = workInfosStr & oInfos.Firmware & VbLf &_
	                              oInfos.Requirements & VbLf &_
	                              oInfos.Zentao
	
	Call setOpenPath(workInfosStr)
End Sub

Sub saveWorkInfosFromOpenPath()
    Dim oInfos, inputArray, i
    Set oInfos = New ProjectInfos
	inputArray = Split(getOpenPath(), VbLf)
	i = 0

    oInfos.Work = trimStr(inputArray(i)) : i = i + 1
    oInfos.Sdk = trimStr(inputArray(i)) : i = i + 1
    oInfos.Product = trimStr(inputArray(i)) : i = i + 1
    oInfos.Project = trimStr(inputArray(i)) : i = i + 1
    If InStr(oInfos.Sdk, "_t0") > 0 Then
        oInfos.SysSdk = trimStr(inputArray(i)) : i = i + 1
        oInfos.SysProject = trimStr(inputArray(i)) : i = i + 1
    End If
    oInfos.Firmware = trimStr(inputArray(i)) : i = i + 1
    oInfos.Requirements = trimStr(inputArray(i)) : i = i + 1
    oInfos.Zentao = trimStr(inputArray(i)) : i = i + 1

	If Not isWorkInfosValid(mIp.Infos) Then
	    MsgBox("Invalid work infos!")
		Exit Sub
	End If

    Dim obj
	For i = vaWorksInfo.Bound To 0 Step -1
		Set obj = vaWorksInfo.V(i)
		If obj.Work = oInfos.Work Or obj.isSameProject(oInfos) Then
		    Call vaWorksInfo.PopBySeq(i)
			Exit For
		End If
	Next

    Call vaWorksInfo.Append(oInfos)
	Call updateWorkInfoTxt()
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
		If InStr(obj.Sdk, "_t0") > 0 Then oTxt.WriteLine(obj.SysSdk)
		If InStr(obj.Sdk, "_t0") > 0 Then oTxt.WriteLine(obj.SysProject)
        oTxt.WriteLine(obj.Firmware)
        oTxt.WriteLine(obj.Requirements)
        oTxt.WriteLine(obj.Zentao)
        oTxt.WriteLine()
    Next

    oTxt.Close
    Set oTxt = Nothing
End Sub

Sub updateNewShortcutBtn()
	Set obj = vaWorksInfo.V(vaWorksInfo.Bound)
	Call addShortcutButton(obj.Work, obj.Sdk, obj.Product, obj.Project, obj.Firmware, obj.Requirements, obj.Zentao, ID_DIV_SHORTCUT)
End Sub

Sub moveShortcutToTop(sName)
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

Sub openFirmwareFolder()
    Call runPath(mIp.Infos.Firmware)
End Sub

Sub openRequirementsFolder()
    Call runPath(mIp.Infos.Requirements)
End Sub

Sub openZentao()
    Call runWebsite(mIp.Infos.Zentao)
End Sub
