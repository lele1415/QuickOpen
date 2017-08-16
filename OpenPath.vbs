Const ID_INPUT_OPEN_PATH = "input_open_path"

Const ID_LIST_OPEN_PATH_FLIE = "list_open_path_file"
Const ID_UL_OPEN_PATH_FLIE = "ul_open_path_file"

Const ID_LIST_OPEN_PATH_DEVICE = "list_open_path_device"
Const ID_UL_OPEN_PATH_DEVICE = "ul_open_path_device"

Const ID_LIST_OPEN_PATH_FRAMEWORK = "list_open_path_framework"
Const ID_UL_OPEN_PATH_FRAMEWORK = "ul_open_path_framework"

Const ID_LIST_OPEN_PATH_KERNEL_LK = "list_open_path_kernel_lk"
Const ID_UL_OPEN_PATH_KERNEL_LK = "ul_open_path_kernel_lk"

Const ID_LIST_OPEN_PATH_PACKAGES = "list_open_path_packages"
Const ID_UL_OPEN_PATH_PACKAGES = "ul_open_path_packages"

Const ID_LIST_OPEN_PATH_VENDOR = "list_open_path_vendor"
Const ID_UL_OPEN_PATH_VENDOR = "ul_open_path_vendor"

Const ID_LIST_OPEN_PATH_OUT = "list_open_path_out"
Const ID_UL_OPEN_PATH_OUT = "ul_open_path_out"

Const PATH_FILE_SYSTEM_PROP = "..\system.prop"
Const PATH_FILE_ITEMS_INI = "..\items.ini"
Const PATH_FILE_PROJECTCONFIG_MK = "..\ProjectConfig.mk"
Const PATH_FILE_DEVICE_MK = "..\device.mk"
Const PATH_FILE_CUSTOM_CONF = "..\custom.conf"
Const PATH_FILE_GMS_MK = "..\gms.mk"
Const PATH_FILE_BUILD_PROP = "..\build.prop"

Const PATH_DEVICE = "device"
Const PATH_DEVICE_PROJECT = "device\..\[project]"
Const PATH_DEVICE_OPTION = "device\..\[option]"
Const PATH_DEVICE_COMMON = "device\mediatek\common"
Const PATH_DEVICE_OVERLAY = "device\mediatek\common\overlay\tablet"

Const PATH_FRAMEWORKS = "frameworks"
Const PATH_FRAMEWORKS_PACKAGES = "frameworks\base\packages"
Const PATH_FRAMEWORKS_PACKAGES_SETTINGSPROVIDER = "frameworks\base\packages\SettingsProvider"
Const PATH_FRAMEWORKS_PACKAGES_SETTINGSPROVIDER_DEFAULTS = "frameworks\base\packages\SettingsProvider\res\values\defaults.xml"
Const PATH_FRAMEWORKS_PACKAGES_SYSTEMUI = "frameworks\base\packages\SystemUI"
Const PATH_FRAMEWORKS_RES = "frameworks\base\core\res\res"
Const PATH_FRAMEWORKS_RES_CONFIG = "frameworks\base\core\res\res\values\config.xml"

Const PATH_KERNEL = "kernel.."
Const PATH_KERNEL_ARM_DTS = "kernel..\arm\..\dts"
Const PATH_KERNEL_ARM64_DTS = "kernel..\arm64\..\dts"
Const PATH_KERNEL_BATTERY = "kernel..\[battery]"
Const PATH_KERNEL_LCM = "kernel..\lcm"
Const PATH_KERNEL_IMGSENSOR = "kernel..\imgsensor"
Const PATH_LK_LCM = "..\lk\..\lcm"

Const PATH_PACKAGES = "packages\apps"
Const PATH_PACKAGES_LAUNCHER3 = "packages\apps\Launcher3"
Const PATH_PACKAGES_SETTINGS = "packages\apps\Settings"

Const PATH_VENDOR = "vendor"
Const PATH_VENDOR_MEDIATEK = "vendor\mediatek\proprietary"
Const PATH_VENDOR_PACKAGES = "vendor\mediatek\proprietary\packages\apps"
Const PATH_VENDOR_GOOGLE = "vendor\google"
Const PATH_VENDOR_HAL = "vendor\..\hal"

Const PATH_OUT = "out\..\"
Const PATH_OUT_SYSTEM = "out\..\system"
Const PATH_OUT_TARGET_FILES = "out\..\target_files_intermediates"

Dim pathDict
Set pathDict = CreateObject("Scripting.Dictionary")

Dim aFUPath_File : aFUPath_File = Array( _
		PATH_FILE_SYSTEM_PROP, _
		PATH_FILE_ITEMS_INI, _
		PATH_FILE_PROJECTCONFIG_MK, _
		PATH_FILE_DEVICE_MK, _
		PATH_FILE_GMS_MK)

Dim aFUPath_device : aFUPath_device = Array( _
		PATH_DEVICE, _
		PATH_DEVICE_PROJECT, _
		PATH_DEVICE_OPTION, _
		PATH_DEVICE_COMMON, _
		PATH_DEVICE_OVERLAY)

Dim aFUPath_framework : aFUPath_framework = Array( _
		PATH_FRAMEWORKS, _
		PATH_FRAMEWORKS_PACKAGES, _
		PATH_FRAMEWORKS_PACKAGES_SETTINGSPROVIDER, _
		PATH_FRAMEWORKS_PACKAGES_SETTINGSPROVIDER_DEFAULTS, _
		PATH_FRAMEWORKS_PACKAGES_SYSTEMUI, _
		PATH_FRAMEWORKS_RES, _
		PATH_FRAMEWORKS_RES_CONFIG)

Dim aFUPath_kernel_lk : aFUPath_kernel_lk = Array( _
		PATH_KERNEL, _
		PATH_KERNEL_ARM_DTS, _
		PATH_KERNEL_ARM64_DTS, _
		PATH_KERNEL_BATTERY, _
		PATH_KERNEL_LCM, _
		PATH_KERNEL_IMGSENSOR, _
		PATH_LK_LCM)

Dim aFUPath_packages : aFUPath_packages = Array( _
		PATH_PACKAGES, _
		PATH_PACKAGES_LAUNCHER3, _
		PATH_PACKAGES_SETTINGS)

Dim aFUPath_vendor : aFUPath_vendor = Array( _
		PATH_VENDOR, _
		PATH_VENDOR_MEDIATEK, _
		PATH_VENDOR_PACKAGES, _
		PATH_VENDOR_GOOGLE, _
		PATH_VENDOR_HAL)

Dim aFUPath_out : aFUPath_out = Array( _
		PATH_OUT, _
		PATH_OUT_SYSTEM, _
		PATH_OUT_TARGET_FILES)

Call onloadFUPath(aFUPath_File, ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH_FLIE, ID_UL_OPEN_PATH_FLIE)
Call onloadFUPath(aFUPath_device, ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH_DEVICE, ID_UL_OPEN_PATH_DEVICE)
Call onloadFUPath(aFUPath_framework, ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH_FRAMEWORK, ID_UL_OPEN_PATH_FRAMEWORK)
Call onloadFUPath(aFUPath_kernel_lk, ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH_KERNEL_LK, ID_UL_OPEN_PATH_KERNEL_LK)
Call onloadFUPath(aFUPath_packages, ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH_PACKAGES, ID_UL_OPEN_PATH_PACKAGES)
Call onloadFUPath(aFUPath_vendor, ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH_VENDOR, ID_UL_OPEN_PATH_VENDOR)
Call onloadFUPath(aFUPath_out, ID_INPUT_OPEN_PATH, ID_LIST_OPEN_PATH_OUT, ID_UL_OPEN_PATH_OUT)
Call getWholePath()

Sub onloadFUPath(aFUPath, inputId, listId, ulId)
	Dim i
	For i = 0 To UBound(aFUPath)
		Call addAfterLi(aFUPath(i), inputId, listId, ulId)
	Next
End Sub

Sub getWholePath()
	Call pathDict.Add(PATH_FILE_SYSTEM_PROP, "[code]\device\joya_sz\[projectName]\roco\[optionName]\system.prop")
	Call pathDict.Add(PATH_FILE_ITEMS_INI, "[code]\device\joya_sz\[projectName]\roco\[optionName]\items.ini")
	Call pathDict.Add(PATH_FILE_PROJECTCONFIG_MK, "[code]\device\joya_sz\[projectName]\ProjectConfig.mk")
	Call pathDict.Add(PATH_FILE_DEVICE_MK, "[code]\device\mediatek\common\device.mk")
	Call pathDict.Add(PATH_FILE_CUSTOM_CONF, "[code]\device\mediatek\common\custom.conf")
	Call pathDict.Add(PATH_FILE_BUILD_PROP, "[code]\out\target\product\[projectName]\system\build.prop")
	Call pathDict.Add(PATH_FILE_GMS_MK, "[code]\vendor\google\products\gms.mk")
	Call pathDict.Add(PATH_DEVICE_PROJECT, "[code]\device\joya_sz\[projectName]")
	Call pathDict.Add(PATH_DEVICE_OPTION, "[code]\device\joya_sz\[projectName]\roco\[optionName]")
	Call pathDict.Add(PATH_KERNEL, "[code]\[kernelName]")
	Call pathDict.Add(PATH_KERNEL_IMGSENSOR, "[code]\[kernelName]\drivers\misc\mediatek\imgsensor")
	Call pathDict.Add(PATH_VENDOR_HAL, "[code]\vendor\mediatek\proprietary\custom\{getPlatformName}\hal")
	Call pathDict.Add(PATH_KERNEL_ARM_DTS, "[code]\[kernelName]\arch\arm\boot\dts")
	Call pathDict.Add(PATH_KERNEL_ARM64_DTS, "[code]\[kernelName]\arch\arm64\boot\dts")
	Call pathDict.Add(PATH_KERNEL_BATTERY, "{getBatteryPath}")
	Call pathDict.Add(PATH_KERNEL_LCM, "[code]\[kernelName]\drivers\misc\mediatek\lcm")
	Call pathDict.Add(PATH_LK_LCM, "{getLkLcmPath}")
	Call pathDict.Add(PATH_OUT, "[code]\out\target\product\[projectName]")
	Call pathDict.Add(PATH_OUT_SYSTEM, "[code]\out\target\product\[projectName]\system")
	Call pathDict.Add(PATH_OUT_TARGET_FILES, "[code]\out\target\product\[projectName]\obj\PACKAGING\target_files_intermediates")


End Sub

Const DO_OPEN_PATH = 0
Const DO_RETURN_PATH = 1
Function handlePath(doWhat)
	Dim code : code = getElementValue(ID_INPUT_CODE_PATH)
	Dim path : path = getElementValue(ID_INPUT_OPEN_PATH)

	If Trim(code) = "" Then Exit Function

	If InStr(path, "..") = 0 Then
		path = code & "\" & path
		path = Replace(path, "/", "\")
		Select Case doWhat
			Case DO_OPEN_PATH : Call runOpenPath(path)
			Case DO_RETURN_PATH : handlePath = path
		End Select
		Exit Function
	End If

	Dim projectName : projectName = getElementValue(ID_INPUT_PROJECT)
	Dim optionName : optionName = getElementValue(ID_INPUT_OPTION)
	Dim kernelName : kernelName = getKernelName(code)

	path = pathDict.Item(path)
	If InStr(path, "[code]") > 0 Then path = Replace(path, "[code]", code)
	If InStr(path, "[projectName]") > 0 Then path = Replace(path, "[projectName]", projectName)
	If InStr(path, "[optionName]") > 0 Then path = Replace(path, "[optionName]", optionName)
	If InStr(path, "[kernelName]") > 0 Then path = Replace(path, "[kernelName]", kernelName)
	If InStr(path, "{getPlatformName}") > 0 Then path = Replace(path, "{getPlatformName}", getPlatformName(code))
	If InStr(path, "{getBatteryPath}") > 0 Then path = Replace(path, "{getBatteryPath}", getBatteryPath(code, kernelName, projectName))
	If InStr(path, "{getLkLcmPath}") > 0 Then path = Replace(path, "{getLkLcmPath}", getLkLcmPath(code))

	Select Case doWhat
		Case DO_OPEN_PATH : Call runOpenPath(path)
		Case DO_RETURN_PATH : handlePath = path
	End Select
End Function

Function getKernelName(code)
	If InStr(code, "l1") > 0 Or InStr(code, "8312") > 0 Then
		getKernelName = "kernel-3.10"
	ElseIf InStr(code, "8167") > 0 Then
		getKernelName = "kernel-4.4"
	Else
		getKernelName = "kernel-3.18"
	End If
End Function

Function getPlatformName(code)
	Select Case True
		Case InStr(code, "8127") > 0
			getPlatformName = "mt8127"
		Case InStr(code, "8163") > 0
			getPlatformName = "mt8163"
		Case InStr(code, "8312") > 0
			getPlatformName = "mt6572"
		Case InStr(code, "8321") > 0
			getPlatformName = "mt6580"
		Case InStr(code, "87") > 0
			getPlatformName = "mt6735"
	End Select
End Function

Function getBatteryPath(code, kernelName, projectName)
	Select Case True
		Case InStr(code, "l18127") > 0
			getBatteryPath = code & kernelName & "\arch\arm\mach-mt8127\" & projectName & "\power"
		Case InStr(code, "l18163") > 0
			getBatteryPath = code & kernelName & "\drivers\misc\mediatek\mach\mt8163\" & projectName & "\power"
		Case InStr(code, "8312") > 0
			getBatteryPath = code & kernelName & "\arch\arm\mach-mt6572\" & projectName & "\power"
		Case InStr(code, "l18321") > 0
			getBatteryPath = code & kernelName & "\misc\mediatek\mach\mt6580\" & projectName & "\power"
		Case InStr(code, "M0") > 0
			getBatteryPath = code & kernelName & "\drivers\misc\mediatek\include\mt-plat\" & getPlatformName(code) & "\include\mach"
	End Select
End Function

Function getLkLcmPath(code)
	If InStr(code, "l1") > 0 Then
		getLkLcmPath = code & "\bootable\bootloader\lk\dev\lcm"
	Else
	    getLkLcmPath = code & "\vendor\mediatek\proprietary\bootable\bootloader\lk\dev\lcm"
	End If
End Function

Sub runOpenPath(path)
	If oFso.FolderExists(path) Then
		oWs.Run "explorer.exe " & path
	ElseIf oFso.FileExists(path) Then
		oWs.Run mTextEditorPath & " " & path
	Else
		MsgBox("not found :" & Vblf & path)
	End If
End Sub