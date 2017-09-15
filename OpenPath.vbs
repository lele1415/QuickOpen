Const PATH_TEXT_EDITOR = "F:\tools\Sublime_Text_3\sublime_text.exe"

Const ID_SELECT_OPEN_PATH = "select_open_path"

Const ID_INPUT_OPEN_PATH = "input_open_path"

Const ID_LIST_OPEN_PATH_SELECT_CLASS = "list_open_path_select_class"
Const ID_UL_OPEN_PATH_SELECT_CLASS = "ul_open_path_select_class"

Const ID_LIST_OPEN_PATH_FILE = "list_open_path_file"
Const ID_UL_OPEN_PATH_FILE = "ul_open_path_file"

Const ID_LIST_OPEN_PATH_FOLDER = "list_open_path_folder"
Const ID_UL_OPEN_PATH_FOLDER = "ul_open_path_folder"

Const ID_INPUT_PROJECT = "input_project"

Const PATH_FILE_PROJECTCONFIG_MK = "..\ProjectConfig.mk"
Const PATH_FILE_SYSTEM_PROP = "..\system.prop"
Const PATH_FILE_PRODUCT_ROCO_MK = "..\product_roco.mk"
Const PATH_FILE_CUSTOM_CONF = "..\custom.conf"
Const PATH_FILE_GMS_MK = "..\gms.mk"

Const PATH_FOLDER_BINARY = "..\binary"
Const PATH_FOLDER_CONFIG = "..\config"
Const PATH_FOLDER_CUSTOM = "..\custom"
Const PATH_FOLDER_OVERLAY = "..\overlay\.."
Const PATH_FOLDER_OUT = "out\..\"

Const VALUE_SELECT_OPEN_PATH_SHOW = "选择路径"
Const VALUE_SELECT_OPEN_PATH_HIDE = "收起"

Dim aFUPath_File : aFUPath_File = Array( _
	    PATH_FILE_PROJECTCONFIG_MK, _
	    PATH_FILE_SYSTEM_PROP, _
	    PATH_FILE_PRODUCT_ROCO_MK, _
	    PATH_FILE_CUSTOM_CONF, _
	    PATH_FILE_GMS_MK)

Dim aFUPath_Folder : aFUPath_Folder = Array( _
	    PATH_FOLDER_BINARY, _
	    PATH_FOLDER_CONFIG, _
	    PATH_FOLDER_CUSTOM, _
	    PATH_FOLDER_OVERLAY, _
	    PATH_FOLDER_OUT)

Call addClassForSelect()
Call onloadFUPath(aFUPath_File, ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH_FILE, ID_UL_OPEN_PATH_FILE)
Call onloadFUPath(aFUPath_Folder, ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH_FOLDER, ID_UL_OPEN_PATH_FOLDER)

Sub addClassForSelect()
    Call addAfterLiForOpenPath("file", ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH_SELECT_CLASS, ID_UL_OPEN_PATH_SELECT_CLASS)
    Call addAfterLiForOpenPath("folder", ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH_SELECT_CLASS, ID_UL_OPEN_PATH_SELECT_CLASS)
End Sub

Sub onloadFUPath(aFUPath, inputId, listId, ulId)
	Dim i
	For i = 0 To UBound(aFUPath)
		Call addAfterLiForOpenPath(aFUPath(i), inputId, listId, ulId)
	Next
End Sub

Sub selectOpenPathOnClick()
	Dim value
	value = getElementValue(ID_SELECT_OPEN_PATH)

    If value = VALUE_SELECT_OPEN_PATH_SHOW Then
    	Call setElementValue(ID_SELECT_OPEN_PATH, VALUE_SELECT_OPEN_PATH_HIDE)
        Call showOrHideOpenPathList(ID_LIST_OPEN_PATH_SELECT_CLASS, "show")
    ElseIf value = VALUE_SELECT_OPEN_PATH_HIDE Then
    	Call setElementValue(ID_SELECT_OPEN_PATH, VALUE_SELECT_OPEN_PATH_SHOW)
        Call HideOpenPathList()
    End If
End Sub

Sub setListValueForOpenPath(inputId, listId, value)
    Call showOrHideOpenPathList(listId, "hide")

    If listId = ID_LIST_OPEN_PATH_SELECT_CLASS Then
        Call showOrHideOpenPathList(Eval("ID_LIST_OPEN_PATH_" & UCase(value)), "show")
    Else
        Call setElementValue(inputId, value)
        Call setElementValue(ID_SELECT_OPEN_PATH, VALUE_SELECT_OPEN_PATH_SHOW)
    End If
End Sub

Const DO_OPEN_PATH = 0
Const DO_RETURN_PATH = 1
Function handlePath(doWhat)
	Dim code : code = getElementValue(ID_INPUT_CODE_PATH)
	Dim path : path = getElementValue(ID_INPUT_OPEN_PATH)

	If Trim(code) = "" Then Exit Function

	If InStr(path, "..\") = 0 Then
		path = code & "\" & path
		path = Replace(path, "/", "\")
		Select Case doWhat
			Case DO_OPEN_PATH : Call runOpenPath(path)
			Case DO_RETURN_PATH : handlePath = path
		End Select
		Exit Function
	End If

	Dim optionName : optionName = getElementValue(ID_INPUT_PROJECT)
	Dim projectName : projectName = getProjectName(optionName)

	Select Case path
		Case PATH_FILE_PROJECTCONFIG_MK
			path = code & "\mediatek\config\" & optionName & "\ProjectConfig.mk"
		Case PATH_FILE_SYSTEM_PROP
			path = code & "\mediatek\config\" & optionName & "\system.prop"
		Case PATH_FILE_PRODUCT_ROCO_MK
			path = code & "\mediatek\config\" & optionName & "\product_roco.mk"
		Case PATH_FILE_CUSTOM_CONF
			path = code & "\mediatek\config\" & optionName & "\custom.conf"			
		Case PATH_FILE_GMS_MK
			path = code & "\mediatek\config\" & optionName & "\gms.mk"
		Case PATH_FOLDER_BINARY
			path = code & "\mediatek\binary\packages\" & optionName
		Case PATH_FOLDER_CONFIG
			path = code & "\mediatek\config\" & optionName
		Case PATH_FOLDER_CUSTOM
			path = code & "\mediatek\custom\" & optionName
		Case PATH_FOLDER_OVERLAY
			Dim sTmp
			sTmp = Replace(optionName, "[", "-")
			sTmp = Replace(sTmp, "]", "")
			path = code & "\mediatek\custom\common\resource_overlay\roco\resandroid\" & sTmp
			If Not oFso.FolderExists(path) Then path = Replace(path, "resandroid", "reslight")
			If Not oFso.FolderExists(path) Then Exit Function
		Case PATH_FOLDER_OUT
			path = code & "\out\target\product\" & projectName
	End Select

	Select Case doWhat
		Case DO_OPEN_PATH : Call runOpenPath(path)
		Case DO_RETURN_PATH : handlePath = path
	End Select
End Function

Function getProjectName(optionName)
	Dim iInStr : iInStr = InStr(optionName, "[")
	If iInStr > 0 Then
		getProjectName = Mid(optionName, 1, InStr(optionName, "[") - 1)
	Else
		getProjectName = optionName
	End If
End Function

Sub runOpenPath(path)
	If oFso.FolderExists(path) Then
		oWs.Run "explorer.exe " & path
	ElseIf oFso.FileExists(path) Then
		oWs.Run """" & PATH_TEXT_EDITOR & """" & " " & path
	Else
		MsgBox("not found :" & Vblf & path)
	End If
End Sub
