Const ID_INPUT_PROJECT = "input_project"
Const ID_LIST_PROJECT = "list_project"
Const ID_UL_PROJECT = "ul_project"
Const ID_INPUT_OPTION = "input_option"
Const ID_LIST_OPTION = "list_option"
Const ID_UL_OPTION = "ul_option"
Dim vaProject, vaOption
Dim mCodePath, pPrjRoot, pOptRoot
Dim mRocoStr : mRocoStr = "roco"
Dim mJoyaStr : mJoyaStr = "joya_sz"

Sub onloadPrj()
	Call setElementValue(ID_INPUT_PROJECT, "")
	Call removeAllChild(ID_UL_PROJECT)

	mJoyaStr = "joya_sz"
	pPrjRoot = getElementValue(ID_INPUT_CODE_PATH) & "\device\" & mJoyaStr
	If Not oFso.FolderExists(pPrjRoot) Then
		mJoyaStr = "mid"
		pPrjRoot = Replace(pPrjRoot, "joya_sz", mJoyaStr)
		If Not oFso.FolderExists(pPrjRoot) Then
			mJoyaStr = "mediateksample"
			pPrjRoot = Replace(pPrjRoot, "mid", mJoyaStr)
			If Not oFso.FolderExists(pPrjRoot) Then MsgBox("Path is not exist:" & Vblf & pPrjRoot) : Exit Sub
		End If
	End If

	Call getAllProject(pPrjRoot)
End Sub

Sub onloadOpt()
	Call setElementValue(ID_INPUT_OPTION, "")
	Call removeAllChild(ID_UL_OPTION)

	mRocoStr = "roco"
	pOptRoot = pPrjRoot & "\" & getElementValue(ID_INPUT_PROJECT) & "\" & mRocoStr
	If Not oFso.FolderExists(pOptRoot) Then
		mRocoStr = "mid"
		pOptRoot = Replace(pOptRoot, "roco", mRocoStr)
		If Not oFso.FolderExists(pOptRoot) Then Exit Sub
	End If

	Call getAllOption(pOptRoot)
End Sub

Sub onloadPrjAndOpt()
	Call onloadPrj()
	Call onloadOpt()
End Sub

Sub getAllProject(pPrjRoot)
	If mJoyaStr = "joya_sz" Then
		Set vaProject = searchFolder(pPrjRoot, "joyasz", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
		If vaProject.Bound = -1 Then
			Set vaProject = searchFolder(pPrjRoot, "jasz", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
			If vaProject.Bound = -1 Then Exit Sub
		End If
	ElseIf mJoyaStr = "mid" Then
		Set vaProject = searchFolder(pPrjRoot, "mt", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	Else
		Set vaProject = searchFolder(pPrjRoot, "tb", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	End If

	Call vaProject.SortArray()

	Dim i
	For i = 0 To vaProject.Bound
		Call addAfterLiForOnloadPrj(vaProject.V(i), ID_INPUT_PROJECT, ID_LIST_PROJECT, ID_UL_PROJECT)
	Next
	'MsgBox(vaProject.ToString())

	Call setElementValue(ID_INPUT_PROJECT, vaProject.V(vaProject.Bound))
End Sub

Sub getAllOption(pOptRoot)
	Set vaOption = searchFolder(pOptRoot, "", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	If vaOption.Bound = -1 Then Exit Sub

	Call vaOption.SortArray()

	Dim i
	For i = 0 To vaOption.Bound
		Call addAfterLiForOnloadPrj(vaOption.V(i), ID_INPUT_OPTION, ID_LIST_OPTION, ID_UL_OPTION)
	Next
	'MsgBox(vaOption.ToString())

	Call setElementValue(ID_INPUT_OPTION, vaOption.V(0))
End Sub

Sub setListValueForOnloadPrj(inputId, listId, value)
    Call showOrHidePrjList(listId, "hide")
    Call setElementValue(inputId, value)

    If inputId = ID_INPUT_PROJECT Then
        Call onloadOpt()
    End If
End Sub
