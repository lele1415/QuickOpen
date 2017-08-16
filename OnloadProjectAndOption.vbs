Const ID_INPUT_PROJECT_L1 = "input_project_l1"
Const ID_LIST_PROJECT_L1 = "list_project_l1"
Const ID_UL_PROJECT_L1 = "ul_project_l1"
Const ID_INPUT_OPTION_L1 = "input_option_l1"
Const ID_LIST_OPTION_L1 = "list_option_l1"
Const ID_UL_OPTION_L1 = "ul_option_l1"
Dim vaProject, vaOption
Dim pPrjRoot, pOptRoot

Sub onloadPrj()
	Call setElementValue(ID_INPUT_PROJECT_L1, "")
	Call removeAllChild(ID_UL_PROJECT_L1)

	pPrjRoot = getElementValue(ID_INPUT_CODE_PATH_L1) & "\device\joya_sz"
	If Not oFso.FolderExists(pPrjRoot) Then MsgBox("Path is not exist:" & Vblf & pPrjRoot) : Exit Sub

	Call getAllProject(pPrjRoot)
End Sub

Sub onloadOpt()
	Call setElementValue(ID_INPUT_OPTION_L1, "")
	Call removeAllChild(ID_UL_OPTION_L1)

	pOptRoot = pPrjRoot & "\" & getElementValue(ID_INPUT_PROJECT_L1) & "\roco"
	If Not oFso.FolderExists(pOptRoot) Then Exit Sub

	Call getAllOption(pOptRoot)
End Sub

Sub onloadPrjAndOpt()
	Call onloadPrj()
	Call onloadOpt()
End Sub

Sub getAllProject(pPrjRoot)
	Set vaProject = searchFolder(pPrjRoot, "joyasz", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	If vaProject.Length = -1 Then Exit Sub

	Call vaProject.SortArray()

	Dim i
	For i = 0 To vaProject.Length
		Call addAfterLi(vaProject.Value(i), ID_INPUT_PROJECT_L1, ID_LIST_PROJECT_L1, ID_UL_PROJECT_L1)
	Next
	'MsgBox(vaProject.ToString())

	Call setElementValue(ID_INPUT_PROJECT_L1, vaProject.Value(0))
End Sub

Sub getAllOption(pOptRoot)
	Set vaOption = searchFolder(pOptRoot, "", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	If vaOption.Length = -1 Then Exit Sub

	Call vaOption.SortArray()

	Dim i
	For i = 0 To vaOption.Length
		Call addAfterLi(vaOption.Value(i), ID_INPUT_OPTION_L1, ID_LIST_OPTION_L1, ID_UL_OPTION_L1)
	Next
	'MsgBox(vaOption.ToString())

	Call setElementValue(ID_INPUT_OPTION_L1, vaOption.Value(0))
End Sub
