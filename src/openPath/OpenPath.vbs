Option Explicit

Dim pPathText : pPathText = oWs.CurrentDirectory & "\res\path.ini"
Dim pathDict, mOverlayPathDict
Set pathDict = CreateObject("Scripting.Dictionary")
Set mOverlayPathDict = CreateObject("Scripting.Dictionary")

Dim vaPathDirectory : Set vaPathDirectory = New VariableArray

Dim mLeftComparePath, mRightComparePath



'Open path
Sub onOpenPathChange()
    Call replaceOpenPath()
End Sub

Sub onOpenListClick()
	Call mOpenPathList.toggleList()
End Sub

Sub onOpenButtonClick()
	If mCmdInput.text <> "" Then Call handleCmdInput() : Exit Sub
	Call removeOpenButtonList()
	Call makeOpenButton()
	If mOpenButtonList.VaArray.Bound = -1 Then
		Call runOpenPath()
	Else
	    Call mOpenButtonList.toggleButtonList()
	End If
End Sub

Sub onOutButtonClick()
	Call mOutFileList.toggleButtonList()
End Sub

Function getOpenPath()
    getOpenPath = mOpenPathInput.text
End Function

Sub setOpenPath(path)
    mOpenPathInput.setText(path)
End Sub

Sub addOpenPathList()
	Call readPathText()
	Call mOpenPathList.addList(vaPathDirectory)
End Sub

Sub readPathText()
    If Not oFso.FileExists(pPathText) Then Exit Sub
    
    Dim oText
    Set oText = oFso.OpenTextFile(pPathText, FOR_READING)

    Dim sReadLine
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

Sub makeOpenButton()
	Dim inputPath
	inputPath = getOpenPath()
	If Trim(inputPath) = "" Or _
	        InStr(inputPath, "weibu/") = 1 Or _
	        InStr(inputPath, "out/") = 1 Then
		Exit Sub
	End If

    If mIp.hasProjectInfos() Then
		Dim wholePath, isFile, fileName
		wholePath = mIp.Infos.Sdk & "/" & inputPath
		If oFso.FileExists(wholePath) Then
			isFile = True
		ElseIf oFso.FolderExists(wholePath) Then
		    isFile = False
		Else
		    Exit Sub
		End If
		fileName = getFileNameFromPath(inputPath)

		Call findOverlayPath("MMI", isFile, fileName)
		Call findOverlayPath("Driver", isFile, fileName)

	    If mOpenButtonList.VaArray.Bound > -1 Then
	    	Call mOverlayPathDict.Add("Origin", wholePath)
	    	Call mOpenButtonList.VaArray.Append("Origin")
	    	Call mOpenButtonList.addList()
	    End If
	End If
End Sub

Sub findOverlayPath(where, isFile, fileName)
	Dim path : path = getOverlayPath(where, isFile, fileName)
	If path <> "" Then
		Call mOverlayPathDict.Add(where, path)
		Call mOpenButtonList.VaArray.Append(where)
	End If
End Sub

Function getOverlayPath(where, isFile, fileName)
	Dim inputPath, wholePath
	inputPath = getOpenPath()

	If where = "MMI" Then
		wholePath = mIp.Infos.getOverlaySdkPath(inputPath)
		If isFile And (Not oFso.FileExists(wholePath)) And mIp.hasProjectAlps() Then
			If oFso.FileExists(mIp.Infos.ProjectSdkPath & "/config/" & fileName) Then
				wholePath = mIp.Infos.ProjectSdkPath & "/config/" & fileName
			End If
		End If
    ElseIf where = "Driver" Then
        wholePath = mIp.Infos.getDriverOverlaySdkPath(inputPath)
        If isFile And (Not oFso.FileExists(wholePath)) And mIp.hasProjectAlps() Then
			If oFso.FileExists(mIp.Infos.DriverProjectSdkPath & "/config/" & fileName) Then
				wholePath = mIp.Infos.DriverProjectSdkPath & "/config/" & fileName
			End If
		End If
    End If

    If isFile And oFso.FileExists(wholePath) Then
    	getOverlayPath = wholePath
    ElseIf (Not isFile) And oFso.FolderExists(wholePath) Then
    	getOverlayPath = wholePath
    Else
        getOverlayPath = ""
    End If
End Function

Function getOpenButtonListPath(where)
	getOpenButtonListPath = mOverlayPathDict.Item(where)
End Function

Sub removeOpenButtonList()
	If mOverlayPathDict.Exists("MMI") Then Call mOverlayPathDict.Remove("MMI")
	If mOverlayPathDict.Exists("Driver") Then Call mOverlayPathDict.Remove("Driver")
	If mOverlayPathDict.Exists("Origin") Then Call mOverlayPathDict.Remove("Origin")
	Call mOpenButtonList.removeList()
End Sub

Sub addOutFileList()
	mOutFileList.VaArray.Append("build.log")
	mOutFileList.VaArray.Append("out")
	mOutFileList.VaArray.Append("target_files")
	mOutFileList.VaArray.Append("system/build.prop")
	mOutFileList.VaArray.Append("vendor/build.prop")
	mOutFileList.VaArray.Append("product/build.prop")
    Call mOutFileList.addList()
End Sub

Function getOutProductPath()
	If mIp.hasProjectInfos() Then
		Dim path : path = mIp.Infos.Sdk & "/out/target/product/" & mIp.Infos.Product

		If Not oFso.FolderExists(path) Then
			If Not oFso.FolderExists(mIp.Infos.Sdk & "/out") Then
				MsgBox("Not found out/")
				runFolderPath(mIp.Infos.Sdk)
				getOutProductPath = ""
			ElseIf Not oFso.FolderExists(mIp.Infos.Sdk & "/out/target/product") Then
				MsgBox("Not found out/target/product/")
				runFolderPath(mIp.Infos.Sdk & "/out")
				getOutProductPath = ""
			Else
			    Dim vaOutProduct : Set vaOutProduct = searchFolder(mIp.Infos.Sdk & "/out/target/product", "", _
		                SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
			    If vaOutProduct.Bound < 0 Then
			        MsgBox("No product folders found in out/target/product/")
			        runFolderPath(mIp.Infos.Sdk & "/out/target/product")
			        getOutProductPath = ""
			    ElseIf vaOutProduct.V(0) <> mIp.Infos.Product Then
			        MsgBox("Other product found in out/target/product/" & Vblf & vaOutProduct.V(0))
			        getOutProductPath = mIp.Infos.Sdk & "/out/target/product/" & vaOutProduct.V(0)
			    End If
			End If
		Else
		    getOutProductPath = path
		End If
	Else
	    getOutProductPath = ""
	End If
End Function

Function getOutListPath(where)
    Dim outProductPath : outProductPath = getOutProductPath()
    If outProductPath = "" Then
    	getOutListPath = ""
    	Exit Function
    End If

    If where = "build.log" Then
        getOutListPath = mIp.Infos.Sdk & "/build.log"
    ElseIf where = "out" Then
        getOutListPath = outProductPath
    ElseIf where = "target_files" Then
        getOutListPath = outProductPath & "/obj/PACKAGING/target_files_intermediates"
    ElseIf where = "system/build.prop" Then
        getOutListPath = outProductPath & "/system/build.prop"
    ElseIf where = "vendor/build.prop" Then
        getOutListPath = outProductPath & "/vendor/build.prop"
    ElseIf where = "product/build.prop" Then
        Dim path_r, path_s
        path_r = outProductPath & "/product/build.prop"
        path_s = outProductPath & "/product/etc/build.prop"
        If oFso.FileExists(path_r) Then
            getOutListPath = path_r
        Else
            getOutListPath = path_s
        End If
    End If
End Function

Sub replaceOpenPath()
	Dim path : path = getOpenPath()

	If InStr(path, "..") > 0 Then
		path = pathDict.Item(path)
    End If
	
	path = Replace(path, "\", "/")
	Call replaceProjectInfoStr(path)
	Call setOpenPath(path)
End Sub

Sub replaceProjectInfoStr(path)
	If InStr(path, "[product]") > 0 Then
		path = Replace(path, "[product]", mIp.Infos.Product)
	End If

	If InStr(path, "[project]") > 0 Then
		path = Replace(path, "[project]", mIp.Infos.Project)
	End If

	If InStr(path, "[project-driver]") > 0 Then
		path = Replace(path, "[project-driver]", mIp.Infos.DriverProject)
	End If

	If InStr(path, "[boot_logo]") > 0 Then
		path = Replace(path, "[boot_logo]", mIp.Infos.BootLogo)
	End If
	
	If InStr(path, "[sys_target_project]") > 0 Then
		path = Replace(path, "[sys_target_project]", mIp.Infos.SysTarget)
	End If

	If InStr(path, "[kernel_version]") > 0 Then
		path = Replace(path, "[kernel_version]", mIp.Infos.KernelVer)
	End If

	If InStr(path, "[target_arch]") > 0 Then
		path = Replace(path, "[target_arch]", mIp.Infos.TargetArch)
	End If

	Call setOpenPath(path)
End Sub

Function runOpenPath()
	If mIp.hasProjectInfos() Then Call runPath(mIp.Infos.Sdk & "\" & getOpenPath())
End Function

Sub compareForProject()
	If Not mIp.hasProjectInfos() Then Exit Sub

    Dim inputPath, wholePath
    inputPath = getOpenPath()
    wholePath = mIp.Infos.Sdk & "/" & inputPath

    If oFso.FileExists(wholePath) Or oFso.FolderExists(wholePath) Then
        Dim projectPath, driverPath

        projectPath = mIp.Infos.ProjectPath & mIp.Infos.ProjectAlps
        driverPath = mIp.Infos.DriverProjectPath & mIp.Infos.ProjectAlps

        If InStr(inputPath, projectPath) > 0 Then
            mLeftComparePath = wholePath
            mRightComparePath = Replace(wholePath, projectPath & "/", "")
        ElseIf InStr(inputPath, driverPath) > 0 Then
            mLeftComparePath = wholePath
            mRightComparePath = Replace(wholePath, driverPath & "/", "")
        Else
            mLeftComparePath = mIp.Infos.getOverlaySdkPath(inputPath)
            mRightComparePath = wholePath
        End If

        mLeftComparePath = """" & Replace(mLeftComparePath, "/", "\") & """"
        mRightComparePath = """" & Replace(mRightComparePath, "/", "\") & """"

        Call runBeyondCompare(mLeftComparePath, mRightComparePath)
    Else
        MsgBox("Not found :" & Vblf & wholePath)
    End If
End Sub

Sub selectForCompare()
    Dim inputPath, wholePath
    inputPath = getOpenPath()
    wholePath = mIp.Infos.Sdk & "/" & inputPath

    If oFso.FileExists(wholePath) Or oFso.FolderExists(wholePath) Then
    	mLeftComparePath = """" & Replace(wholePath, "/", "\") & """"
    	Call hideElement(ID_SELECT_FOR_COMPARE)
    	Call showElement(ID_COMPARE_TO)
    Else
        MsgBox("Not found :" & Vblf & wholePath)
    End If
End Sub

Sub compareTo()
    Dim inputPath, wholePath
    inputPath = getOpenPath()
    wholePath = mIp.Infos.Sdk & "/" & inputPath

    If oFso.FileExists(wholePath) Or oFso.FolderExists(wholePath) Then
    	mRightComparePath = """" & Replace(wholePath, "/", "\") & """"
    	Call hideElement(ID_COMPARE_TO)
    	Call showElement(ID_SELECT_FOR_COMPARE)
    	Call runBeyondCompare(mLeftComparePath, mRightComparePath)
    Else
        MsgBox("Not found :" & Vblf & wholePath)
    End If
End Sub

Sub cleanOpenPath()
	setOpenPath("")
End Sub

Sub openMMI()
	If mIp.hasProjectInfos() Then runFolderPath(mIp.Infos.ProjectSdkPath)
End Sub

Sub openDriver()
	If mIp.hasProjectInfos() Then runFolderPath(mIp.Infos.DriverProjectSdkPath)
End Sub

Sub replaceSlash()
	Call setOpenPath(Replace(getOpenPath(), "\", "/"))
End Sub

Sub addProjectPath()
	Call setOpenPath(mIp.Infos.ProjectPath & mIp.Infos.ProjectAlps & "/" & getOpenPath())
End Sub

Sub addDriverProjectPath()
	Call setOpenPath(mIp.Infos.DriverProjectPath & mIp.Infos.ProjectAlps & "/" & getOpenPath())
End Sub

Sub cutSdkPath()
	Call mIp.cutSdkInOpenPath()
End Sub

Sub cutProjectPath()
	Call mIp.cutProjectInOpenPath()
End Sub

Dim pJavaFileText : pJavaFileText = oWs.CurrentDirectory & "\res\file_list_all_java.txt"
Dim pFrameworksJavaFileText : pFrameworksJavaFileText = oWs.CurrentDirectory & "\res\file_list_frameworks_java.txt"
Dim pAndroidmkFileText : pAndroidmkFileText = oWs.CurrentDirectory & "\res\file_list_androidmk.txt"

Sub addFileList()
	If mFileButtonList.VaArray.Bound = 0 Then
        Call setOpenPath(mFileButtonList.VaArray.V(0))
        Call mFileButtonList.VaArray.ResetArray()
    ElseIf mFileButtonList.VaArray.Bound > 0 Then
        Call mFileButtonList.addList()
        Call mFileButtonList.toggleButtonList()
    End If
End Sub

Function getFileListPathFromRes(name)
	getFileListPathFromRes = oWs.CurrentDirectory & "\res\filelist\" & getSdkSimpleName() & "\" & name
End Function

Sub findJavaFile()
    If mFileButtonList.hideListIfShowing() Then Exit Sub

	Dim fileListPath : fileListPath = getFileListPathFromRes("java.txt")
	If oFso.FileExists(fileListPath) Then
		Call makeFileList(fileListPath, ".java")
		If mFileButtonList.VaArray.Bound = -1 Then
			fileListPath = getFileListPathFromRes("kotlin.txt")
			If oFso.FileExists(fileListPath) Then
				Call makeFileList(fileListPath, ".kt")
			End If
		End If
	Else
		Call CopyString("find -type f -name *.java > java.txt")
		MsgBox("java.txt not exist!")
	End If
End Sub

Sub findFrameworksJavaFile()
    If mFileButtonList.hideListIfShowing() Then Exit Sub

	Dim fileListPath : fileListPath = getFileListPathFromRes("f-java.txt")
	If oFso.FileExists(fileListPath) Then
		Call makeFileList(getFileListPathFromRes("f-java.txt"), ".java")
	Else
		Call CopyString("find frameworks/ -type f -name *.java > f-java.txt")
		MsgBox("f-java.txt not exist!")
	End If
End Sub

Sub findXmlFile()
    If mFileButtonList.hideListIfShowing() Then Exit Sub

	Dim fileListPath : fileListPath = getFileListPathFromRes("xml.txt")
	If oFso.FileExists(fileListPath) Then
		Call makeFileList(getFileListPathFromRes("xml.txt"), ".xml")
	Else
		Call CopyString("find -type f -name *.xml > xml.txt")
		MsgBox("xml.txt not exist!")
	End If
End Sub

Sub findAppFolder()
    If mFileButtonList.hideListIfShowing() Then Exit Sub

	Dim fileListPath : fileListPath = getFileListPathFromRes("app.txt")
	If oFso.FileExists(fileListPath) Then
		Call makeFileList(getFileListPathFromRes("app.txt"), "app")
	Else
		Call CopyString("find -type f -name Android.* > app.txt")
		MsgBox("app.txt not exist!")
	End If
End Sub

Sub findPackageFile()
    If mFileButtonList.hideListIfShowing() Then Exit Sub
	If Not mIp.hasProjectInfos() Or Trim(getOpenPath()) = "" Then Exit Sub
	
	If Not oFso.FileExists(mIp.Infos.Sdk & "\pkg.txt") Then
		If Not oFso.FolderExists(mIp.Infos.Sdk & "\" & getOpenPath()) Then
			MsgBox("Path not exist!" & VbLf & mIp.Infos.Sdk & "\" & getOpenPath())
		    Exit Sub
		End If
		Call replaceSlash()
		Call CopyString("find " & getOpenPath() & " -type f -name *.java > pkg.txt")
		MsgBox("pkg.txt not exist!")
	Else
	    Call makeFileList(mIp.Infos.Sdk & "\pkg.txt", ".java")
	End If
End Sub

Sub makeFileList(fileListPath, suffix)
	If Trim(getOpenPath()) = "" Or InStr(getOpenPath(), "/") > 0 Or InStr(getOpenPath(), "\") > 0 Then Exit Sub

	If suffix <> "app" Then
        Call findFileInListText(getOpenPath(), suffix, fileListPath)
    Else
        Call findAppFolderInListText(getOpenPath(), fileListPath)
    End If

    Call addFileList()
End Sub

Sub findFileInListText(input, suffix, path)
	Dim oText, sReadLine, keyStr, count
	If Not InStr(input, ".") > 0 Then
		keyStr = input & suffix
	Else
	    keyStr = input
	End If
	keyStr = "/" & keyStr
	Set oText = oFso.OpenTextFile(path, FOR_READING)
	count = 0

	Do Until oText.AtEndOfStream
	    sReadLine = oText.ReadLine
	    If count > 10 Then Exit Do
	    If Right(sReadLine, Len(keyStr)) = keyStr Then
	        Call mFileButtonList.VaArray.Append(Replace(sReadLine, "./", ""))
	        count = count + 1
	    End If
	Loop

	oText.Close
	Set oText = Nothing
End Sub

Sub findAppFolderInListText(input, path)
	Dim oText, sReadLine, keyStr, count
	keyStr = "/" & input & "/Android."
	Set oText = oFso.OpenTextFile(path, FOR_READING)
	count = 0

	Do Until oText.AtEndOfStream
	    If count > 10 Then Exit Do
	    sReadLine = oText.ReadLine
	    If InStr(sReadLine, keyStr) > 0 Then
	    	path = Left(sReadLine, InStr(sReadLine, keyStr) + Len(input))
	        Call mFileButtonList.VaArray.Append(Replace(path, "./", ""))
	        count = count + 1
	    End If
	Loop
End Sub

Sub pasteAndOpenPath()
    Call setElementValue(ID_INPUT_OPEN_PATH, "")
    Call focusElement(ID_INPUT_OPEN_PATH)
    oWs.SendKeys "^v"
    oWs.SendKeys "{ENTER}"
End Sub

Sub tabOpenPath()
    Dim tabStr
	tabStr = getTabStr()
    If tabStr <> "" Then Call setOpenPath(getOpenPath() & tabStr)
End Sub
