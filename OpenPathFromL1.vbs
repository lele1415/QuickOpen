Const PATH_TEXT_EDITOR = "F:\tools\Sublime_Text_3\sublime_text.exe"

Const ID_INPUT_OPEN_PATH_L1 = "input_open_path_l1"
Const ID_LIST_OPEN_PATH_L1 = "list_open_path_l1"
Const ID_UL_OPEN_PATH_L1 = "ul_open_path_l1"

Const PATH_SYSTEM_PROP = "..\system.prop"
Const PATH_ITEMS_INI = "..\items.ini"
Const PATH_PROJECTCONFIG_MK = "..\ProjectConfig.mk"
Const PATH_DEVICE_MK = "..\device.mk"
Const PATH_GMS_MK = "..\gms.mk"

Const PATH_BUILD = "build"
Const PATH_DEVICE = "device"
Const PATH_DEVICE_COMMON = "device\mediatek\common"
Const PATH_DEVICE_OVERLAY = "device\mediatek\common\overlay\tablet"
Const PATH_FRAMEWORKS = "frameworks"
Const PATH_FRAMEWORKS_PACKAGES = "frameworks\base\packages"
Const PATH_FRAMEWORKS_RES = "frameworks\base\core\res\res"
Const PATH_KERNEL = "kernel.."
Const PATH_PACKAGES = "packages\apps"
Const PATH_VENDOR = "vendor"
Const PATH_VENDOR_PACKAGES = "vendor\mediatek\proprietary\packages\apps"

Const PATH_DEVICE_PROJECT = "device\..\[project]"
Const PATH_DEVICE_OPTION = "device\..\[option]"
Const PATH_KERNEL_IMGSENSOR = "kernel..\imgsensor"
Const PATH_VENDOR_HAL = "vendor\..\hal"
Const PATH_KERNEL_ARM_DTS = "kernel..\arm\..\dts"
Const PATH_KERNEL_ARM64_DTS = "kernel..\arm64\..\dts"
Const PATH_KERNEL_BATTERY = "kernel..\[battery]"
Const PATH_KERNEL_LCM = "kernel..\lcm"
Const PATH_LK_LCM = "..\lk\..\lcm"
Const PATH_OUT = "out\..\"

Dim aFUPath : aFUPath = Array( _
		PATH_SYSTEM_PROP, _
		PATH_ITEMS_INI, _
		PATH_PROJECTCONFIG_MK, _
		PATH_DEVICE_MK, _
		PATH_GMS_MK, _
		PATH_BUILD, _
		PATH_DEVICE, _
		PATH_DEVICE_COMMON, _
		PATH_DEVICE_OVERLAY, _
		PATH_FRAMEWORKS, _
		PATH_FRAMEWORKS_PACKAGES, _
		PATH_FRAMEWORKS_RES, _
		PATH_KERNEL, _
		PATH_PACKAGES, _
		PATH_VENDOR, _
		PATH_VENDOR_PACKAGES, _
		PATH_DEVICE_PROJECT, _
		PATH_DEVICE_OPTION, _
		PATH_KERNEL_IMGSENSOR, _
		PATH_VENDOR_HAL, _
		PATH_KERNEL_ARM_DTS, _
		PATH_KERNEL_ARM64_DTS, _
		PATH_KERNEL_BATTERY, _
		PATH_KERNEL_LCM, _
		PATH_LK_LCM, _
		PATH_OUT)

Call onloadFUPathFromL1()

Sub onloadFUPathFromL1()
	Dim i
	For i = 0 To UBound(aFUPath)
		Call addAfterLi(aFUPath(i), ID_INPUT_OPEN_PATH_L1, ID_LIST_OPEN_PATH_L1, ID_UL_OPEN_PATH_L1)
	Next
End Sub

Const DO_OPEN_PATH = 0
Const DO_RETURN_PATH = 1
Function handlePathFromL1(doWhat)
	Dim code : code = getElementValue(ID_INPUT_CODE_PATH_L1)
	Dim path : path = getElementValue(ID_INPUT_OPEN_PATH_L1)

	If Trim(code) = "" Then Exit Function

	Dim kernelName : kernelName = getKernelNameFromL1(code)

	If InStr(path, "..") = 0 Then
		path = code & "\" & path
		path = Replace(path, "/", "\")
		Select Case doWhat
			Case DO_OPEN_PATH : Call runOpenPath(path)
			Case DO_RETURN_PATH : handlePathFromL1 = path
		End Select
		Exit Function
	End If

	Dim projectName : projectName = getElementValue(ID_INPUT_PROJECT_L1)
	Dim optionName : optionName = getElementValue(ID_INPUT_OPTION_L1)

	'If Trim(projectName) = "" Or Trim(optionName) = "" Then Exit Function

	Select Case path
		Case PATH_SYSTEM_PROP
			path = code & "\device\joya_sz\" & projectName & "\roco\" & optionName & "\system.prop"
		Case PATH_ITEMS_INI
			path = code & "\device\joya_sz\" & projectName & "\roco\" & optionName & "\items.ini"
		Case PATH_PROJECTCONFIG_MK
			path = code & "\device\joya_sz\" & projectName & "\ProjectConfig.mk"
		Case PATH_DEVICE_MK
			path = code & "\device\mediatek\common\device.mk"
		Case PATH_GMS_MK
			path = code & "\vendor\google\products\gms.mk"
		Case PATH_DEVICE_PROJECT
			path = code & "\device\joya_sz\" & projectName
		Case PATH_DEVICE_OPTION
			path = code & "\device\joya_sz\" & projectName & "\roco\" & optionName
		Case PATH_KERNEL
			path = code & kernelName
		Case PATH_KERNEL_IMGSENSOR
			path = code & kernelName & "\drivers\misc\mediatek\imgsensor"
		Case PATH_VENDOR_HAL
			path = code & "\vendor\mediatek\proprietary\custom\" & getPlatformNameFromL1(code) & "\hal"
		Case PATH_KERNEL_ARM_DTS
			path = code & kernelName & "\arch\arm\boot\dts"
		Case PATH_KERNEL_ARM64_DTS
			path = code & kernelName & "\arch\arm64\boot\dts"
		Case PATH_KERNEL_BATTERY
			path = getBatteryPathFromL1(code, kernelName, projectName)
		Case PATH_KERNEL_LCM
			path = code & kernelName & "\drivers\misc\mediatek\lcm"
		Case PATH_LK_LCM
			path = getLkLcmPathFromL1(code, kernelName, projectName)
		Case PATH_OUT
			path = code & "\out\target\product\" & projectName
	End Select

	Select Case doWhat
		Case DO_OPEN_PATH : Call runOpenPath(path)
		Case DO_RETURN_PATH : handlePathFromL1 = path
	End Select
End Function

Function getKernelNameFromL1(code)
	If InStr(code, "l1") > 0 Or InStr(code, "8312") > 0 Then
		getKernelNameFromL1 = "\kernel-3.10"
	Else
		getKernelNameFromL1 = "\kernel-3.18"
	End If
End Function

Function getPlatformNameFromL1(code)
	Select Case True
		Case InStr(code, "8127") > 0
			getPlatformNameFromL1 = "mt8127"
		Case InStr(code, "8163") > 0
			getPlatformNameFromL1 = "mt8163"
		Case InStr(code, "8312") > 0
			getPlatformNameFromL1 = "mt6572"
		Case InStr(code, "8321") > 0
			getPlatformNameFromL1 = "mt6580"
		Case InStr(code, "87") > 0
			getPlatformNameFromL1 = "mt6735"
	End Select
End Function

Function getBatteryPathFromL1(code, kernelName, projectName)
	Select Case True
		Case InStr(code, "l18127") > 0
			getBatteryPathFromL1 = code & kernelName & "\arch\arm\mach-mt8127\" & projectName & "\power"
		Case InStr(code, "l18163") > 0
			getBatteryPathFromL1 = code & kernelName & "\drivers\misc\mediatek\mach\mt8163\" & projectName & "\power"
		Case InStr(code, "8312") > 0
			getBatteryPathFromL1 = code & kernelName & "\arch\arm\mach-mt6572\" & projectName & "\power"
		Case InStr(code, "l18321") > 0
			getBatteryPathFromL1 = code & kernelName & "\misc\mediatek\mach\mt6580\" & projectName & "\power"
		Case InStr(code, "M0") > 0
			getBatteryPathFromL1 = code & kernelName & "\drivers\misc\mediatek\include\mt-plat\" & getPlatformNameFromL1(code) & "\include\mach"
	End Select
End Function

Function getLkLcmPathFromL1(code)
	If InStr(code, "l1") > 0 Then
		getLkLcmPathFromL1 = code & "\bootable\bootloader\lk\dev\lcm"
	Else
	    getLkLcmPathFromL1 = code & "\vendor\mediatek\proprietary\bootable\bootloader\lk\dev\lcm"
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