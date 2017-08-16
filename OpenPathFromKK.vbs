Const ID_INPUT_OPEN_PATH_KK = "input_open_path_kk"
Const ID_LIST_OPEN_PATH_KK = "list_open_path_kk"
Const ID_UL_OPEN_PATH_KK = "ul_open_path_kk"

Const ID_INPUT_PROJECT_KK = "input_project_kk"

Const PATH_KK_PROJECTCONFIG_MK = "..\ProjectConfig.mk"
Const PATH_KK_SYSTEM_PROP = "..\system.prop"
Const PATH_KK_PRODUCT_ROCO_MK = "..\product_roco.mk"
Const PATH_KK_CUSTOM_CONF = "..\custom.conf"
Const PATH_KK_GMS_MK = "..\gms.mk"
Const PATH_KK_BINARY = "..\binary"
Const PATH_KK_CONFIG = "..\config"
Const PATH_KK_CUSTOM = "..\custom"
Const PATH_KK_OVERLAY = "..\overlay\.."
Const PATH_KK_OUT = "out\..\"

Dim aFUPathFromKK : aFUPathFromKK = Array( _
	    PATH_KK_PROJECTCONFIG_MK, _
	    PATH_KK_SYSTEM_PROP, _
	    PATH_KK_PRODUCT_ROCO_MK, _
	    PATH_KK_CUSTOM_CONF, _
	    PATH_KK_GMS_MK, _
	    PATH_KK_BINARY, _
	    PATH_KK_CONFIG, _
	    PATH_KK_CUSTOM, _
	    PATH_KK_OVERLAY, _
	    PATH_KK_OUT)

Call onloadFUPathFromKK()

Sub onloadFUPathFromKK()
	Dim i
	For i = 0 To UBound(aFUPathFromKK)
		Call addAfterLi(aFUPathFromKK(i), ID_INPUT_OPEN_PATH_KK, ID_LIST_OPEN_PATH_KK, ID_UL_OPEN_PATH_KK)
	Next
End Sub

Function handlePathFromKK(doWhat)
	Dim code : code = getElementValue(ID_INPUT_CODE_PATH_KK)
	Dim path : path = getElementValue(ID_INPUT_OPEN_PATH_KK)

	If Trim(code) = "" Then Exit Function

	If InStr(path, "..\") = 0 Then
		path = code & "\" & path
		path = Replace(path, "/", "\")
		Select Case doWhat
			Case DO_OPEN_PATH : Call runOpenPath(path)
			Case DO_RETURN_PATH : handlePathFromKK = path
		End Select
		Exit Function
	End If

	Dim optionName : optionName = getElementValue(ID_INPUT_PROJECT_KK)
	Dim projectName : projectName = getProjectName(optionName)

	Select Case path
		Case PATH_KK_PROJECTCONFIG_MK
			path = code & "\mediatek\config\" & optionName & "\ProjectConfig.mk"
		Case PATH_KK_SYSTEM_PROP
			path = code & "\mediatek\config\" & optionName & "\system.prop"
		Case PATH_KK_PRODUCT_ROCO_MK
			path = code & "\mediatek\config\" & optionName & "\product_roco.mk"
		Case PATH_KK_CUSTOM_CONF
			path = code & "\mediatek\config\" & optionName & "\custom.conf"			
		Case PATH_KK_GMS_MK
			path = code & "\mediatek\config\" & optionName & "\gms.mk"
		Case PATH_KK_BINARY
			path = code & "\mediatek\binary\packages\" & optionName
		Case PATH_KK_CONFIG
			path = code & "\mediatek\config\" & optionName
		Case PATH_KK_CUSTOM
			path = code & "\mediatek\custom\" & optionName
		Case PATH_KK_OVERLAY
			Dim sTmp
			sTmp = Replace(optionName, "[", "-")
			sTmp = Replace(sTmp, "]", "")
			path = code & "\mediatek\custom\common\resource_overlay\roco\resandroid\" & sTmp
			If Not oFso.FolderExists(path) Then path = Replace(path, "resandroid", "reslight")
			If Not oFso.FolderExists(path) Then Exit Function
		Case PATH_KK_OUT
			path = code & "\out\target\product\" & projectName
	End Select

	Select Case doWhat
		Case DO_OPEN_PATH : Call runOpenPath(path)
		Case DO_RETURN_PATH : handlePathFromKK = path
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
