Const PATH_TEXT_EDITOR = "F:\tools\Sublime_Text_3\sublime_text.exe"
Const ID_INPUT_OPEN_PATH = "input_open_path"
Const ID_LIST_OPEN_PATH = "list_open_path"
Const ID_UL_OPEN_PATH = "ul_open_path"

Const ID_INPUT_PROJECT = "input_project"

Const PATH_PROJECTCONFIG_MK = "..\ProjectConfig.mk"
Const PATH_SYSTEM_PROP = "..\system.prop"
Const PATH_PRODUCT_ROCO_MK = "..\product_roco.mk"
Const PATH_CUSTOM_CONF = "..\custom.conf"
Const PATH_GMS_MK = "..\gms.mk"
Const PATH_BINARY = "..\binary"
Const PATH_CONFIG = "..\config"
Const PATH_CUSTOM = "..\custom"
Const PATH_OVERLAY = "..\overlay\.."
Const PATH_OUT = "out\..\"

Dim aFUPath : aFUPath = Array( _
	    PATH_PROJECTCONFIG_MK, _
	    PATH_SYSTEM_PROP, _
	    PATH_PRODUCT_ROCO_MK, _
	    PATH_CUSTOM_CONF, _
	    PATH_GMS_MK, _
	    PATH_BINARY, _
	    PATH_CONFIG, _
	    PATH_CUSTOM, _
	    PATH_OVERLAY, _
	    PATH_OUT)

Call onloadFUPath()

Sub onloadFUPath()
	Dim i
	For i = 0 To UBound(aFUPath)
		Call addAfterLi(aFUPath(i), ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH, ID_UL_OPEN_PATH)
	Next
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
		Case PATH_PROJECTCONFIG_MK
			path = code & "\mediatek\config\" & optionName & "\ProjectConfig.mk"
		Case PATH_SYSTEM_PROP
			path = code & "\mediatek\config\" & optionName & "\system.prop"
		Case PATH_PRODUCT_ROCO_MK
			path = code & "\mediatek\config\" & optionName & "\product_roco.mk"
		Case PATH_CUSTOM_CONF
			path = code & "\mediatek\config\" & optionName & "\custom.conf"			
		Case PATH_GMS_MK
			path = code & "\mediatek\config\" & optionName & "\gms.mk"
		Case PATH_BINARY
			path = code & "\mediatek\binary\packages\" & optionName
		Case PATH_CONFIG
			path = code & "\mediatek\config\" & optionName
		Case PATH_CUSTOM
			path = code & "\mediatek\custom\" & optionName
		Case PATH_OVERLAY
			Dim sTmp
			sTmp = Replace(optionName, "[", "-")
			sTmp = Replace(sTmp, "]", "")
			path = code & "\mediatek\custom\common\resource_overlay\roco\resandroid\" & sTmp
			If Not oFso.FolderExists(path) Then path = Replace(path, "resandroid", "reslight")
			If Not oFso.FolderExists(path) Then Exit Function
		Case PATH_OUT
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
