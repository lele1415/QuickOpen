Const ID_INPUT_PROJECT = "input_project"
Const ID_LIST_PROJECT = "list_project"
Const ID_UL_PROJECT = "ul_project"
Const ID_INPUT_OPTION = "input_option"
Const ID_LIST_OPTION = "list_option"
Const ID_UL_OPTION = "ul_option"
Dim vaProject, vaOption
Dim pPrjRoot, pOptRoot

Sub onloadPrj()
	Call setElementValue(ID_INPUT_PROJECT, "")
	Call removeAllChild(ID_UL_PROJECT)

	pPrjRoot = getElementValue(ID_INPUT_CODE_PATH) & "\device\joya_sz"
	If Not oFso.FolderExists(pPrjRoot) Then MsgBox("Path is not exist:" & Vblf & pPrjRoot) : Exit Sub

	Call getAllProject(pPrjRoot)
End Sub

Sub onloadOpt()
	Call setElementValue(ID_INPUT_OPTION, "")
	Call removeAllChild(ID_UL_OPTION)

	pOptRoot = pPrjRoot & "\" & getElementValue(ID_INPUT_PROJECT) & "\roco"
	If Not oFso.FolderExists(pOptRoot) Then Exit Sub

	Call getAllOption(pOptRoot)
End Sub

Sub onloadPrjAndOpt()
	Call onloadPrj()
	Call onloadOpt()
End Sub

Sub getAllProject(pPrjRoot)
	Set vaProject = searchFolder(pPrjRoot, "joyasz", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	If vaProject.Bound = -1 Then Exit Sub

	Call vaProject.SortArray()

	Dim i
	For i = 0 To vaProject.Bound
		Call addAfterLi(vaProject.V(i), ID_INPUT_PROJECT, ID_LIST_PROJECT, ID_UL_PROJECT)
	Next
	'MsgBox(vaProject.ToString())

	Call setElementValue(ID_INPUT_PROJECT, vaProject.V(0))
End Sub

Sub getAllOption(pOptRoot)
	Set vaOption = searchFolder(pOptRoot, "", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	If vaOption.Bound = -1 Then Exit Sub

	Call vaOption.SortArray()

	Dim i
	For i = 0 To vaOption.Bound
		Call addAfterLi(vaOption.V(i), ID_INPUT_OPTION, ID_LIST_OPTION, ID_UL_OPTION)
	Next
	'MsgBox(vaOption.ToString())

	Call setElementValue(ID_INPUT_OPTION, vaOption.V(0))
End Sub
