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
	If mCmdInput.text <> "" Then
	    Dim cmd : cmd = mCmdInput.text
	    Call handleCmdInput()
		If mCmdInput.text = "" Then Call saveHistoryCmd(cmd)
		Exit Sub
	End If
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

Function getCmdText()
    getCmdText = mCmdInput.text
End Function

Sub setCmdText(path)
    mCmdInput.setText(path)
End Sub

Sub addOpenPathList()
	Call readPathText()
	Call mOpenPathList.addList(vaPathDirectory)
End Sub

Sub readPathText()
    If Not isFileExists(pPathText) Then Exit Sub
    
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
		Dim isFile
		If isFileExists(inputPath) Then
			isFile = True
		ElseIf isFolderExists(inputPath) Then
		    isFile = False
		Else
		    Exit Sub
		End If

		Call findOverlayPath(isFile)

	    If mOpenButtonList.VaArray.Bound > -1 Then
	    	Call mOverlayPathDict.Add("Origin", inputPath)
	    	Call mOpenButtonList.VaArray.Append("Origin")
	    	Call mOpenButtonList.addList()
	    End If
	End If
End Sub

Function findOverlayPath(isFile)
	Dim projectName, newName, inputPath, wholePath, configFilePath
	projectName = mIp.Infos.Project
	inputPath = getOpenPath()
	wholePath = mIp.Infos.getOverlayPath(inputPath)
	If isFile Then configFilePath = mIp.Infos.ProjectPath & "/config/" & getFileNameFromPath(inputPath)

    Do
		If isFile And (Not isFileExists(wholePath)) Then
			If isFileExists(configFilePath) Then
				wholePath = configFilePath
			End If
		End If

		If isFile Then
		    If isFileExists(wholePath) Then
				Call mOverlayPathDict.Add(projectName, wholePath)
				Call mOpenButtonList.VaArray.Append(projectName)
			ElseIf isFileExists(configFilePath) Then
			    Call mOverlayPathDict.Add(projectName, configFilePath)
				Call mOpenButtonList.VaArray.Append(projectName)
			End If
			    
		ElseIf isFolderExists(wholePath) Then
			Call mOverlayPathDict.Add(projectName, wholePath)
			Call mOpenButtonList.VaArray.Append(projectName)
		End If

		If InStr(projectName, "-") > 0 Then
		    newName = Left(projectName, InStrRev(projectName, "-") - 1)
			wholePath = Replace(wholePath, projectName, newName)
			If isFile Then configFilePath = Replace(configFilePath, projectName, newName)
			projectName = newName
		Else
		    Exit Do
		End If
	Loop
End Function

Function getOpenButtonListPath(where)
	getOpenButtonListPath = mOverlayPathDict.Item(where)
End Function

Sub removeOpenButtonList()
    Call mOverlayPathDict.RemoveAll()
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

Function getOutProductPath(sys)
	If mIp.hasProjectInfos() Then
		Dim path, outName
		If sys Then
			path = mIp.Infos.SysOutPath
		Else
			path = mIp.Infos.OutPath
		End If
		outName = Left(path, InStr(path, "/") - 1)

		If Not isFolderExists(path) Then
			If Not isFolderExists(outName) Then
				MsgBox("Not found " & outName)
				getOutProductPath = ""
			ElseIf Not isFolderExists(outName & "/target/product") Then
				MsgBox("Not found out/target/product/")
				getOutProductPath = ""
			Else
			    Dim vaOutProduct : Set vaOutProduct = searchFolder(outName & "/target/product", "", _
		                SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
			    If vaOutProduct.Bound < 0 Then
			        MsgBox("No product folders found in out/target/product/")
			        getOutProductPath = ""
			    ElseIf vaOutProduct.V(0) <> mIp.Infos.Product Then
			        'MsgBox("Other product found in " & outName & "/target/product/" & Vblf & vaOutProduct.V(0))
			        getOutProductPath = outName & "/target/product/" & vaOutProduct.V(0)
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
    Dim outProductPath : outProductPath = getOutProductPath(False)
    If outProductPath = "" Then
    	getOutListPath = ""
    	Exit Function
    End If

    If where = "build.log" Then
        getOutListPath = "build.log"
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
        If isFileExists(path_r) Then
            getOutListPath = path_r
        Else
            getOutListPath = path_s
        End If
    End If
End Function

Sub replaceOpenPath()
	Dim path : path = getOpenPath()
	If InStr(path, ":\") > 0 Or InStr(path, "\\192.168") > 0 Then Exit Sub

	If InStr(path, "..") > 0 Then
		path = pathDict.Item(path)
    End If
	path = replaceProjectInfoStr(path)
	Call setOpenPath(relpaceSlashInPath(path))

	'Call cutSdkPath()
End Sub

Function replaceProjectInfoStr(path)
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

	replaceProjectInfoStr = path
End Function

Function runOpenPath()
	If mIp.hasProjectInfos() Then Call runPath(getOpenPath())
End Function

Sub compareForProject()
	If Not mIp.hasProjectInfos() Then Exit Sub

    Dim inputPath, wholePath
    inputPath = getOpenPath()

    If isFileExists(inputPath) Or isFolderExists(inputPath) Then
        Dim projectPath, driverPath

        projectPath = mIp.Infos.ProjectPath & mIp.Infos.ProjectAlps
        driverPath = mIp.Infos.DriverProjectPath & mIp.Infos.ProjectAlps

        If InStr(inputPath, projectPath) > 0 Then
            mLeftComparePath = inputPath
            mRightComparePath = Replace(inputPath, projectPath & "/", "")
        ElseIf InStr(inputPath, driverPath) > 0 Then
            mLeftComparePath = inputPath
            mRightComparePath = Replace(inputPath, driverPath & "/", "")
        Else
            mLeftComparePath = mIp.Infos.getOverlayPath(inputPath)
            mRightComparePath = inputPath
        End If

        mLeftComparePath = """" & mLeftComparePath & """"
        mRightComparePath = """" & mRightComparePath & """"

        Call runBeyondCompare(mLeftComparePath, mRightComparePath)
    Else
        MsgBox("Not found :" & Vblf & inputPath)
    End If
End Sub

Sub selectForCompare()
    Dim inputPath
    inputPath = getOpenPath()

    If isFileExists(inputPath) Or isFolderExists(inputPath) Then
    	mLeftComparePath = """" & mIp.Infos.getPathWithDriveSdk(inputPath) & """"
    	Call hideElement(ID_SELECT_FOR_COMPARE)
    	Call showElement(ID_COMPARE_TO)
    Else
        MsgBox("Not found :" & Vblf & inputPath)
    End If
End Sub

Sub compareTo()
    Dim inputPath
    inputPath = getOpenPath()

    If isFileExists(inputPath) Or isFolderExists(inputPath) Then
    	mRightComparePath = """" & mIp.Infos.getPathWithDriveSdk(inputPath) & """"
    	Call hideElement(ID_COMPARE_TO)
    	Call showElement(ID_SELECT_FOR_COMPARE)
    	Call runBeyondCompare(mLeftComparePath, mRightComparePath)
    Else
        MsgBox("Not found :" & Vblf & inputPath)
    End If
End Sub

Sub cleanOpenPath()
	setOpenPath("")
End Sub

Sub openMMI()
	If mIp.hasProjectInfos() Then runFolderPath(mIp.Infos.ProjectPath)
End Sub

Sub openDriver()
	If mIp.hasProjectInfos() Then runFolderPath(mIp.Infos.DriverProjectPath)
End Sub

Sub replaceSlash()
	Call setOpenPath(relpaceSlashInPath(getOpenPath()))
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
	If isFileExists(fileListPath) Then
		Call makeFileList(fileListPath, ".java")
		If mFileButtonList.VaArray.Bound = -1 Then
			fileListPath = getFileListPathFromRes("kotlin.txt")
			If isFileExists(fileListPath) Then
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
	If isFileExists(fileListPath) Then
		Call makeFileList(getFileListPathFromRes("f-java.txt"), ".java")
	Else
		Call CopyString("find frameworks/ -type f -name *.java > f-java.txt")
		MsgBox("f-java.txt not exist!")
	End If
End Sub

Sub findXmlFile()
    If mFileButtonList.hideListIfShowing() Then Exit Sub

	Dim fileListPath : fileListPath = getFileListPathFromRes("xml.txt")
	If isFileExists(fileListPath) Then
		Call makeFileList(getFileListPathFromRes("xml.txt"), ".xml")
	Else
		Call CopyString("find -type f -name *.xml > xml.txt")
		MsgBox("xml.txt not exist!")
	End If
End Sub

Sub findAppFolder()
    If mFileButtonList.hideListIfShowing() Then Exit Sub

	Dim fileListPath : fileListPath = getFileListPathFromRes("app.txt")
	If isFileExists(fileListPath) Then
		Call makeFileList(getFileListPathFromRes("app.txt"), "app")
	Else
		Call CopyString("find -type f -name Android.* > app.txt")
		MsgBox("app.txt not exist!")
	End If
End Sub

Sub findPackageFile()
    If mFileButtonList.hideListIfShowing() Then Exit Sub
	If Not mIp.hasProjectInfos() Or Trim(getOpenPath()) = "" Then Exit Sub
	
	If Not isFileExists("pkg.txt") Then
		If Not isFolderExists(getOpenPath()) Then
			MsgBox("Path not exist!" & VbLf & getOpenPath())
		    Exit Sub
		End If
		Call replaceSlash()
		Call CopyString("find " & getOpenPath() & " -type f -name *.java > pkg.txt")
		MsgBox("pkg.txt not exist!")
	Else
	    Call makeFileList(mIp.Infos.getPathWithDriveSdk("pkg.txt"), ".java")
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
