

Dim pPathText : pPathText = oWs.CurrentDirectory & "\path.ini"
Dim pathDict
Set pathDict = CreateObject("Scripting.Dictionary")

Dim vaPathDirectory : Set vaPathDirectory = New VariableArray

Call readPathText(pPathText)
Call setOpenPathParentIds()
Call addPathList()
Call getWholePath()




Sub readPathText(DictPath)
    If Not oFso.FileExists(DictPath) Then Exit Sub
    
    Dim oText
    Set oText = oFso.OpenTextFile(DictPath, FOR_READING)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine

        If InStr(sReadLine, "{") > 0 Then
            Call getAllPath(oText, sReadLine)
        End If
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub getAllPath(oText, sReadLine)
    Dim directoryName : directoryName = Trim(Replace(sReadLine, "{", ""))
    if directoryName <> "" Then
        Dim vaPath : Set vaPath = New VariableArray
        vaPath.Name = directoryName

        sReadLine = oText.ReadLine
        Do until InStr(sReadLine, "}") > 0
            vaPath.Append(Trim(sReadLine))
            sReadLine = oText.ReadLine
        Loop

        vaPathDirectory.Append(vaPath)
    End If
End Sub

Sub addPathList()
    If vaPathDirectory.Bound <> -1 Then
        Call setOpenPathDirectoryIds()
        Call addListUL()
        Dim i, j, category

        For i = 0 To vaPathDirectory.Bound
            category = LCase(vaPathDirectory.V(i).Name)
            Call setOpenPathDirectoryIds()
            Call addListDirectoryLi(category, getOpenPathDivId() & category)

            Call setOpenPathListIds(category)
            Call addListUL()

            if vaPathDirectory.V(i).Bound <> -1 Then
                For j = 0 To vaPathDirectory.V(i).Bound
                    Call addListLi(vaPathDirectory.V(i).V(j))
                Next
            End If
        Next
    End If
End Sub

Sub setOpenPathParentIds()
	Call setListParentAndInputIds(getParentOpenPathId(), getOpenPathInputId())
	Call setListDirectoryDivId(getOpenPathDirectoryDivId())
End Sub

Sub setOpenPathDirectoryIds()
	Call setListDivIds(getOpenPathDirectoryDivId(), getOpenPathDirectoryULId())
End Sub

Sub setOpenPathListIds(category)
	Call setListDivIds(getOpenPathDivId() & category, getOpenPathULId() & category)
End Sub

Sub getWholePath()
	Call pathDict.Add("../buildinfo.sh", "build/make/tools/buildinfo.sh")
	Call pathDict.Add("../buildinfo_common.sh", "build/make/tools/buildinfo_common.sh")
	Call pathDict.Add("../Makefile", "build/make/core/Makefile")
	Call pathDict.Add("../ProjectConfig.mk", "device/mediateksample/[product]/ProjectConfig.mk")
	Call pathDict.Add("../SystemConfig.mk", "device/mediatek/system/[sys_target_project]/SystemConfig.mk")
	Call pathDict.Add("../full_[product].mk", "device/mediateksample/[product]/full_[product].mk")
	Call pathDict.Add("../vnd_[product].mk", "device/mediateksample/[product]/vnd_[product].mk")
	Call pathDict.Add("../device.mk", "device/mediatek/system/common/device.mk")
	Call pathDict.Add("../apns-conf.xml", "device/mediatek/config/apns-conf.xml")
	Call pathDict.Add("../system.prop", "device/mediatek/system/common/system.prop")
	Call pathDict.Add("../custom.conf", "device/mediatek/vendor/common/custom.conf")
	Call pathDict.Add("frameworks/../android", "frameworks/base/core/java/android")
	Call pathDict.Add("frameworks/../services", "frameworks/base/services/core/java/com/android/server")
	Call pathDict.Add("frameworks/../values", "frameworks/base/core/res/res/values")
	Call pathDict.Add("frameworks/../config.xml", "frameworks/base/core/res/res/values/config.xml")
	Call pathDict.Add("../SystemUI/../config.xml", "vendor/mediatek/proprietary/packages/apps/SystemUI/res/values/config.xml")
	Call pathDict.Add("../SettingsProvider/../defaults.xml", "vendor/mediatek/proprietary/packages/apps/SettingsProvider/res/values/defaults.xml")
	Call pathDict.Add("../SettingsProvider/../DatabaseHelper.java", "vendor/mediatek/proprietary/packages/apps/SettingsProvider/src/com/android/providers/settings/DatabaseHelper.java")
	Call pathDict.Add("../label.ini", "vendor/mediatek/proprietary/buildinfo_sys/label.ini")
	Call pathDict.Add("build.log", "build.log")
	Call pathDict.Add("out/..", "out/target/product/[product]")
	Call pathDict.Add("../target_files", "out/target/product/[product]/obj/PACKAGING/target_files_intermediates")
	Call pathDict.Add("../system/build.prop", "out/target/product/[product]/system/build.prop")
	Call pathDict.Add("../vendor/build.prop", "out/target/product/[product]/vendor/build.prop")
	Call pathDict.Add("../product/build.prop", "out/target/product/[product]/product/build.prop")
	Call pathDict.Add("../product/etc/build.prop", "out/target/product/[product]/product/etc/build.prop")
End Sub

Sub replaceOpenPath()
	Dim path : path = getOpenPath()
	If InStr(path, "..") > 0 Then
		path = pathDict.Item(path)
		If InStr(path, "[product]") > 0 Then
			path = Replace(path, "[product]", getProduct())
		End If
		If InStr(path, "[sys_target_project]") > 0 Then
			path = Replace(path, "[sys_target_project]", getSysTargetProject())
		End If
		Call setOpenPath(path)
    End If
End Sub

Function getSysTargetProject()
	Dim fullMkPath
	fullMkPath = getSdkPath() & "/" & "device/mediateksample/" & getProduct() & "/full_" & getProduct() & ".mk"
	If Not oFso.FileExists(fullMkPath) Then Exit Function

	Dim oText, sReadLine, exitFlag
	Set oText = oFso.OpenTextFile(fullMkPath, FOR_READING)

	Do Until oText.AtEndOfStream
	    sReadLine = oText.ReadLine
	    If InStr(sReadLine, "SYS_TARGET_PROJECT") > 0 Then
	        getSysTargetProject = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
	        Exit Do
	    End If
	Loop

	oText.Close
	Set oText = Nothing
	MsgBox("getSysTargetProject end")
End Function

Function getOpenPath()
	getOpenPath = getElementValue(getOpenPathInputId())
End Function

Function setOpenPath(path)
	Call setElementValue(getOpenPathInputId(), path)
End Function

Const DO_OPEN_PATH = 0
Const DO_RETURN_PATH = 1
Const DO_COPY_PATH = 2
Function handlePath(doWhat)
	If Trim(getSdkPath()) = "" Then Exit Function

	path = getSdkPath() & "\" & getOpenPath()
	path = Replace(path, "/", "\")
	Select Case doWhat
		Case DO_OPEN_PATH : Call runOpenPath(path)
		Case DO_RETURN_PATH : handlePath = path
		Case DO_COPY_PATH : Call CopyString(path)
	End Select
End Function

Function getKernelName(code, product)
	If InStr(code, "l1") > 0 Or InStr(code, "8312") > 0 Then
		getKernelName = "kernel-3.10"
	ElseIf InStr(code, "8167") > 0 Then
		getKernelName = "kernel-4.4"
	ElseIf InStr(code, "O18735B") > 0 Then
		If InStr(product, "8735") > 0 Then
			getKernelName = "kernel-3.18"
		Else
			getKernelName = "kernel-4.4"
		End If
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
		Case InStr(code, "8167") > 0
			getPlatformName = "mt8167"
		Case InStr(code, "8312") > 0
			getPlatformName = "mt6572"
		Case InStr(code, "8321") > 0
			getPlatformName = "mt6580"
		Case InStr(code, "87") > 0
			getPlatformName = "mt6735"
	End Select
End Function

Function getHalDir(product)
	Select Case True
		Case InStr(product, "tb8735ap1") > 0
			getHalDir = "D1"
		Case InStr(product, "tb8735ba1") > 0
			getHalDir = "D2"
		Case InStr(product, "tb8735ma1") > 0
			getHalDir = "D2"
		Case InStr(product, "tb8735p1") > 0
			getHalDir = "D1"
	End Select
End Function

Function getBatteryPath(code, kernelName, product)
	Select Case True
		Case InStr(code, "l18127") > 0
			getBatteryPath = kernelName & "\arch\arm\mach-mt8127\" & product & "\power"
		Case InStr(code, "l18163") > 0
			getBatteryPath = kernelName & "\drivers\misc\mediatek\mach\mt8163\" & product & "\power"
		Case InStr(code, "8312") > 0
			getBatteryPath = kernelName & "\arch\arm\mach-mt6572\" & product & "\power"
		Case InStr(code, "l18321") > 0
			getBatteryPath = kernelName & "\misc\mediatek\mach\mt6580\" & product & "\power"
		Case InStr(code, "M0") > 0
			getBatteryPath = kernelName & "\drivers\misc\mediatek\include\mt-plat\" & getPlatformName(code) & "\include\mach"
	End Select
End Function

Function getLkLcmPath(code)
	If InStr(code, "l1") > 0 Then
		getLkLcmPath = "\bootable\bootloader\lk\dev\lcm"
	Else
	    getLkLcmPath = "\vendor\mediatek\proprietary\bootable\bootloader\lk\dev\lcm"
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

Sub runBeyondCompare()
	Dim inputPath, wholePath
	inputPath = getOpenPath()
	wholePath = getSdkPath() & "/" & inputPath

	If oFso.FileExists(wholePath) Or oFso.FolderExists(wholePath) Then
	    Dim leftPath, rightPath, projectPath

		projectPath = getProjectPathWithoutSdk()

		If InStr(inputPath, projectPath) > 0 Then
			leftPath = wholePath
			rightPath = Replace(wholePath, projectPath, "")
			rightPath = Replace(rightPath, "//", "/")
		Else
			leftPath = getSdkPath() & "/" & projectPath & "/" & inputPath
			rightPath = wholePath
		End If

		leftPath = """" & Replace(leftPath, "/", "\") & """"
		rightPath = """" & Replace(rightPath, "/", "\") & """"

		Dim command : command = mBeyondComparePath & " " & leftPath & " " & rightPath
		oWs.Run command
	Else
		MsgBox("Not found :" & Vblf & wholePath)
	End If
End Sub
