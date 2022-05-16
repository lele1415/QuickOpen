

Dim pPathText : pPathText = oWs.CurrentDirectory & "\res\path.ini"
Dim pathDict
Set pathDict = CreateObject("Scripting.Dictionary")

Dim vaPathDirectory : Set vaPathDirectory = New VariableArray

Call readPathText(pPathText)
Call setOpenPathParentIds()
Call addPathList()



'Open path
Sub onOpenPathChange()
    Call replaceOpenPath()
End Sub

Function getOpenPath()
    getOpenPath = mOpenPathInput.text
End Function

Sub setOpenPath(path)
    mOpenPathInput.setText(path)
End Sub

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
        	Dim a
        	a = Split(sReadLine, ":")
            vaPath.Append(Trim(a(0)))
            Call pathDict.Add(Trim(a(0)), Trim(a(1)))
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

Sub replaceOpenPath()
	Dim path : path = getOpenPath()
	If InStr(path, "..") > 0 Then
		path = pathDict.Item(path)

		If InStr(path, "[product]") > 0 Then
			path = Replace(path, "[product]", mIp.Infos.Product)
		End If

		If InStr(path, "[project]") > 0 Then
			path = Replace(path, "[project]", mIp.Infos.Project)
		End If

		If InStr(path, "[project-driver]") > 0 Then
			path = Replace(path, "[project-driver]", Replace(mIp.Infos.Project, "-MMI", ""))
		End If

		If InStr(path, "[boot_logo]") > 0 Then
			path = Replace(path, "[boot_logo]", getBootLogo())
		End If
		
		If InStr(path, "[sys_target_project]") > 0 Then
			path = Replace(path, "[sys_target_project]", getSysTargetProject())
		End If

		Call setOpenPath(path)
    End If
End Sub

Function getSysTargetProject()
	Dim fullMkPath
	fullMkPath = mIp.Infos.Sdk & "/" & "device/mediateksample/" & mIp.Infos.Product & "/full_" & mIp.Infos.Product & ".mk"
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
End Function

' Function getOpenPath()
' 	getOpenPath = getElementValue(getOpenPathInputId())
' End Function

' Function setOpenPath(path)
' 	Call setElementValue(getOpenPathInputId(), path)
' End Function

Const DO_OPEN_PATH = 0
Const DO_RETURN_PATH = 1
Const DO_COPY_PATH = 2
Function handlePath(doWhat)
	If Trim(mIp.Infos.Sdk) = "" Then Exit Function

	path = mIp.Infos.Sdk & "\" & getOpenPath()
	path = Replace(path, "/", "\")
	Select Case doWhat
		Case DO_OPEN_PATH : Call runOpenPath(path)
		Case DO_RETURN_PATH : handlePath = path
		Case DO_COPY_PATH : Call CopyString(path)
	End Select
End Function

' Function getKernelName(sdk, product)
' 	If InStr(sdk, "l1") > 0 Or InStr(sdk, "8312") > 0 Then
' 		getKernelName = "kernel-3.10"
' 	ElseIf InStr(sdk, "8167") > 0 Then
' 		getKernelName = "kernel-4.4"
' 	ElseIf InStr(sdk, "O18735B") > 0 Then
' 		If InStr(product, "8735") > 0 Then
' 			getKernelName = "kernel-3.18"
' 		Else
' 			getKernelName = "kernel-4.4"
' 		End If
' 	Else
' 		getKernelName = "kernel-3.18"
' 	End If
' End Function

' Function getPlatformName(sdk)
' 	Select Case True
' 		Case InStr(sdk, "8127") > 0
' 			getPlatformName = "mt8127"
' 		Case InStr(sdk, "8163") > 0
' 			getPlatformName = "mt8163"
' 		Case InStr(sdk, "8167") > 0
' 			getPlatformName = "mt8167"
' 		Case InStr(sdk, "8312") > 0
' 			getPlatformName = "mt6572"
' 		Case InStr(sdk, "8321") > 0
' 			getPlatformName = "mt6580"
' 		Case InStr(sdk, "87") > 0
' 			getPlatformName = "mt6735"
' 	End Select
' End Function

' Function getHalDir(product)
' 	Select Case True
' 		Case InStr(product, "tb8735ap1") > 0
' 			getHalDir = "D1"
' 		Case InStr(product, "tb8735ba1") > 0
' 			getHalDir = "D2"
' 		Case InStr(product, "tb8735ma1") > 0
' 			getHalDir = "D2"
' 		Case InStr(product, "tb8735p1") > 0
' 			getHalDir = "D1"
' 	End Select
' End Function

' Function getBatteryPath(sdk, kernelName, product)
' 	Select Case True
' 		Case InStr(sdk, "l18127") > 0
' 			getBatteryPath = kernelName & "\arch\arm\mach-mt8127\" & product & "\power"
' 		Case InStr(sdk, "l18163") > 0
' 			getBatteryPath = kernelName & "\drivers\misc\mediatek\mach\mt8163\" & product & "\power"
' 		Case InStr(sdk, "8312") > 0
' 			getBatteryPath = kernelName & "\arch\arm\mach-mt6572\" & product & "\power"
' 		Case InStr(sdk, "l18321") > 0
' 			getBatteryPath = kernelName & "\misc\mediatek\mach\mt6580\" & product & "\power"
' 		Case InStr(sdk, "M0") > 0
' 			getBatteryPath = kernelName & "\drivers\misc\mediatek\include\mt-plat\" & getPlatformName(sdk) & "\include\mach"
' 	End Select
' End Function

' Function getLkLcmPath(sdk)
' 	If InStr(sdk, "l1") > 0 Then
' 		getLkLcmPath = "\bootable\bootloader\lk\dev\lcm"
' 	Else
' 	    getLkLcmPath = "\vendor\mediatek\proprietary\bootable\bootloader\lk\dev\lcm"
' 	End If
' End Function

Sub runOpenPath(path)
	If oFso.FolderExists(path) Then
		oWs.Run "explorer.exe " & path
	ElseIf oFso.FileExists(path) Then
	    If isPictureFilePath(path) Or isCompressFilePath(path) Then
	    	oWs.Run "explorer.exe " & path
	    Else
		    oWs.Run mTextEditorPath & " " & path
		End If
	Else
		MsgBox("not found :" & Vblf & path)
	End If
End Sub

Sub runBeyondCompare()
	Dim inputPath, wholePath
	inputPath = getOpenPath()
	wholePath = mIp.Infos.Sdk & "/" & inputPath

	If oFso.FileExists(wholePath) Or oFso.FolderExists(wholePath) Then
	    Dim leftPath, rightPath, projectPath

		projectPath = mIp.Infos.ProjectPath

		If InStr(inputPath, projectPath) > 0 Then
			leftPath = wholePath
			rightPath = Replace(wholePath, projectPath, "")
			rightPath = Replace(rightPath, "//", "/")
		Else
			leftPath = mIp.Infos.ProjectSdkPath & "/" & inputPath
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

Sub runWebsite(path)
	oWs.Run mBrowserPath & " " & path
End Sub
