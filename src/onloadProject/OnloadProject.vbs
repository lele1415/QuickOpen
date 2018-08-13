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
	idTimer = window.setTimeout( _
			"Call getAllProject(""" & prj & """, """ & opt & """)", _
			0, "VBScript")
End Sub

Sub setPrjElement(prj)
	Call setElementValue(ID_INPUT_PROJECT, "")
	Call removeAllChild(ID_UL_PROJECT)

	Dim i
	For i = 0 To vaProject.Bound
		Call addAfterLiForOnloadPrj(vaProject.V(i), ID_INPUT_PROJECT, ID_LIST_PROJECT, ID_UL_PROJECT)
	Next

	If prj <> "" Then
		Call setElementValue(ID_INPUT_PROJECT, prj)
	Else
		Call setElementValue(ID_INPUT_PROJECT, vaProject.V(vaProject.Bound))
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

	Call setPrjElement(prj)

	Call onloadOpt(opt)
End Sub

Function checkRocoStr()
	Dim result, path
	result = True
	path = pPrjRoot & "\" & getElementValue(ID_INPUT_PROJECT) & "\"

	Select Case True
		Case oFso.FolderExists(path & STR_ROCO_MID)
		    mRocoStr = STR_ROCO_MID
		Case oFso.FolderExists(path & STR_ROCO_ROCO)
		    mRocoStr = STR_ROCO_ROCO
		Case Else
		    MsgBox("Roco path is not exist") : result = False
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
	If Not checkRocoStr() Then Exit Sub
	idTimer = window.setTimeout( _
			"getAllOption(""" & opt & """)", _
			0, "VBScript")
End Sub

Sub onloadPrjAndOpt()
	Call onloadPrj("", "")
End Sub

Sub getAllOption(opt)
	window.clearTimeout(idTimer)
	Set vaOption = searchFolder(pOptRoot, "", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	If vaOption.Bound = -1 Then MsgBox("No roco folder found!") : Exit Sub

	Call vaOption.SortArray()
	'MsgBox(vaOption.ToString())

	Call setOptElement(opt)
End Sub

Sub setListValueForOnloadPrj(inputId, listId, value)
    Call showOrHidePrjList(listId, "hide")
    Call setElementValue(inputId, value)

    If inputId = ID_INPUT_PROJECT Then
        Call onloadOpt()
    End If
End Sub
