

Dim pPathText : pPathText = oWs.CurrentDirectory & "\res\path.ini"
Dim pathDict, mOverlayPathDict
Set pathDict = CreateObject("Scripting.Dictionary")
Set mOverlayPathDict = CreateObject("Scripting.Dictionary")

Dim vaPathDirectory : Set vaPathDirectory = New VariableArray
Dim vaOpenPathList : Set vaOpenPathList = New VariableArray

Call addOpenPathList()
Call addBuildpropList()



'Open path
Sub onOpenPathChange()
    Call replaceOpenPath()
End Sub

Sub onOpenButtonClick()
	Call removeOpenButtonList()
	Call makeOpenButton()
	If vaOpenPathList.Bound = -1 Then
		Call runOpenPath()
	Else
	    Call mOpenButtonList.toggleButtonList()
	End If
End Sub

Function getOpenPath()
    getOpenPath = mOpenPathInput.text
End Function

Sub setOpenPath(path)
    mOpenPathInput.setText(path)
End Sub

Sub addOpenPathList()
	Call readPathText(pPathText)
	Call mOpenPathList.addList(vaPathDirectory)
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

	    If vaOpenPathList.Bound > -1 Then
	    	Call mOverlayPathDict.Add("Origin", wholePath)
	    	Call vaOpenPathList.Append("Origin")
	    	Call mOpenButtonList.addList(vaOpenPathList)
	    End If
	End If
End Sub

Sub findOverlayPath(where, isFile, fileName)
	path = getOverlayPath(where, isFile, fileName)
	If path <> "" Then
		Call mOverlayPathDict.Add(where, path)
		Call vaOpenPathList.Append(where)
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
	Call vaOpenPathList.ResetArray()
	Call mOpenButtonList.removeList()
End Sub

Sub addBuildpropList()
	Dim vaBuildprop : Set vaBuildprop = New VariableArray
	vaBuildprop.Append("build.log")
	vaBuildprop.Append("out")
	vaBuildprop.Append("target_files")
	vaBuildprop.Append("system/build.prop")
	vaBuildprop.Append("vendor/build.prop")
	vaBuildprop.Append("product/build.prop")
    Call mBuildpropList.addList(vaBuildprop)
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
        Dim product_r, pruduct_s
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
	path = Replace(path, "\", "/")
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
    End If

	Call setOpenPath(path)
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

Function runOpenPath()
	If mIp.hasProjectInfos() Then Call runPath(mIp.Infos.Sdk & "\" & getOpenPath())
End Function

Sub runBeyondCompare()
	If Not mIp.hasProjectInfos() Then Exit Sub

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
            leftPath = mIp.Infos.getOverlaySdkPath(inputPath)
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

Sub cleanOpenPath()
	setOpenPath("")
End Sub

Sub openMMI()
	If mIp.hasProjectInfos() Then runFolderPath(mIp.Infos.ProjectSdkPath)
End Sub

Sub openDriver()
	If mIp.hasProjectInfos() Then runFolderPath(mIp.Infos.DriverProjectSdkPath)
End Sub
