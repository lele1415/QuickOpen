Const ID_INPUT_PROJECT = "input_project"
Const ID_LIST_PROJECT = "list_project"
Const ID_UL_PROJECT = "ul_project"
Const ID_INPUT_OPTION = "input_option"
Const ID_LIST_OPTION = "list_option"
Const ID_UL_OPTION = "ul_option"

Const STR_JOYA_O1 = "mediateksample"
Const STR_JOYA_MID = "mid"
Const STR_JOYA_JOYA = "joya_sz"
Const STR_ROCO_MID = "mid"
Const STR_ROCO_ROCO = "ROCO"

Dim vaProject, vaOption
Dim mCodePath, pPrjRoot, pOptRoot
Dim mRocoStr, mJoyaStr
Dim mPrjName, mPrjIndex
Dim mLoadOptInAllPrj

Function checkJoyaStr()
	Dim result, path
	result = True
	path = getElementValue(ID_INPUT_CODE_PATH) & "\device\"

	Select Case True
		Case oFso.FolderExists(path & STR_JOYA_O1)
		    mJoyaStr = STR_JOYA_O1
		Case oFso.FolderExists(path & STR_JOYA_MID)
		    mJoyaStr = STR_JOYA_MID
		Case oFso.FolderExists(path & STR_JOYA_JOYA)
		    mJoyaStr = STR_JOYA_JOYA
		Case Else
		    MsgBox("Joya path is not exist") : result = False
	End Select

	If result Then pPrjRoot = path & mJoyaStr

	checkJoyaStr = result
End Function

Sub onloadPrj(prj, opt)
	If Not checkJoyaStr() Then Exit Sub
	Call freezeAllInput()
	idTimer = window.setTimeout( _
			"Call getAllProject(""" & prj & """, """ & opt & """)", _
			0, "VBScript")
End Sub

Sub addNewPrjLi()
	Call removeAllChild(ID_UL_PROJECT)

	Dim i
	For i = 0 To vaProject.Bound
		Call addAfterLiForOnloadPrj(vaProject.V(i), ID_INPUT_PROJECT, ID_LIST_PROJECT, ID_UL_PROJECT)
	Next
End Sub

Sub findValidPrj(opt)
	mPrjName = vaProject.V(mPrjIndex)
	Call setElementValue(ID_INPUT_PROJECT, mPrjName)
	mLoadOptInAllPrj = True
	Call onloadOpt(opt)
End Sub

Sub setPrjElement(prj, opt)
	Call setElementValue(ID_INPUT_PROJECT, "")

	If prj <> "" Then
		Call setElementValue(ID_INPUT_PROJECT, prj)
		mPrjName = prj
		Call onloadOpt(opt)
	Else
		mPrjIndex = vaProject.Bound
		Call findValidPrj(opt)
	End If
End Sub

Sub getAllProject(prj, opt)
	window.clearTimeout(idTimer)
	Select Case mJoyaStr
		Case STR_JOYA_O1
			Set vaProject = searchFolder(pPrjRoot, "tb", _
				    SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
		Case STR_JOYA_MID
			Set vaProject = searchFolder(pPrjRoot, "mt", _
				    SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
		Case STR_JOYA_JOYA
			Set vaProject = searchFolder(pPrjRoot, "joyasz", _
				    SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
			If vaProject.Bound = -1 Then
				Set vaProject = searchFolder(pPrjRoot, "jasz", _
					    SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
			End If
	End Select

	If vaProject.Bound = -1 Then MsgBox("No joya folder found!") : Exit Sub

	Call vaProject.SortArray()
	'MsgBox(vaProject.ToString())

	Call addNewPrjLi()
	Call setPrjElement(prj, opt)

	
End Sub

Function checkRocoStr()
	Dim result, path
	result = True
	path = pPrjRoot & "\" & mPrjName & "\"

	Select Case True
		Case oFso.FolderExists(path & STR_ROCO_MID)
		    mRocoStr = STR_ROCO_MID
		Case oFso.FolderExists(path & STR_ROCO_ROCO)
		    mRocoStr = STR_ROCO_ROCO
		Case Else
		    result = False
	End Select

	If result Then pOptRoot = path & mRocoStr

	checkRocoStr = result
End Function

Sub setOptElement(opt)
	Call setElementValue(ID_INPUT_OPTION, "")
	Call removeAllChild(ID_UL_OPTION)

	Dim i
	For i = 0 To vaOption.Bound
		Call addAfterLiForOnloadPrj(vaOption.V(i), ID_INPUT_OPTION, ID_LIST_OPTION, ID_UL_OPTION)
	Next

	If opt <> "" Then
		Call setElementValue(ID_INPUT_OPTION, opt)
	Else
		Call setElementValue(ID_INPUT_OPTION, vaOption.V(0))
	End If
End Sub

Sub onloadOpt(opt)
	If Not checkRocoStr() Then
		Call findOptInNextPrj(opt)
		Exit Sub
	End If
	idTimer = window.setTimeout( _
			"getAllOption(""" & opt & """)", _
			0, "VBScript")
End Sub

Sub onloadPrjAndOpt()
	Call onloadPrj("", "")
End Sub

Sub findOptInNextPrj(opt)
	If mLoadOptInAllPrj And mPrjIndex > 0 Then
		mPrjIndex = mPrjIndex - 1
		Call findValidPrj(opt)
	Else
		MsgBox("No roco folder found!")
	End If
End Sub

Sub getAllOption(opt)
	window.clearTimeout(idTimer)
	Set vaOption = searchFolder(pOptRoot, "", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	If vaOption.Bound = -1 Then
		Call findOptInNextPrj(opt)
		Exit Sub
	End If

	Call vaOption.SortArray()
	'MsgBox(vaOption.ToString())

	If opt <> "" Then
		Call setOptElement(opt)
	Else
		Call setOptElement(vaOption.V(0))
	End If

	Call unfreezeAllInput()
End Sub

Sub setListValueForOnloadPrj(inputId, listId, value)
    Call showOrHidePrjList(listId, "hide")
    Call setElementValue(inputId, value)

    If inputId = ID_INPUT_PROJECT Then
    	mPrjName = value
    	mLoadOptInAllPrj = False
        Call onloadOpt("")
    End If
End Sub
