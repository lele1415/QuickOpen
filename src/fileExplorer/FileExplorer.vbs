Const ID_DIV_EXP_FILE = "exp_file"
Const ID_DIV_EXP_PATH = "exp_path"
Const EXP_SHOW = 0
Const EXP_HIDE = 1

Dim iCrtPathLength, sCrtPath
iCrtPathLength = 0
sCrtPath = ""

Dim vaSubFolderName : Set vaSubFolderName = New VariableArray
Dim vaSubFileName : Set vaSubFileName = New VariableArray

Sub initNewPath(doWhat)
	Dim cutLength : cutLength = iCrtPathLength
	Call delPath(cutLength)
	Call removeFile()

	If doWhat = EXP_SHOW Then
		Dim rootPath : rootPath = handlePath(1)
		If isValidRootPath(rootPath) Then
			Call addRootPath(rootPath)
			Call addFile()
		End If
	End If
End Sub

Sub clickPath(iPathLength)
	Dim cutLength : cutLength = iCrtPathLength - iPathLength

	If cutLength = 0 Then
		Call openFolder(sCrtPath)
	Else
		Call delPath(cutLength)
		Call removeFile()
		Call addFile()
	End If
End Sub

Sub clickPlus(folderName)
	Call addPath(folderName)
	Call removeFile()
	Call addFile()
End Sub

Sub clickFolder(folderName)
	Call openFolder(sCrtPath & "\" & folderName)
End Sub

Sub clickFile(fileName)
	Call openFile(sCrtPath & "\" & fileName)
End Sub



Sub addRootPath(rootPath)
	Dim aFileName, i
	aFileName = Split(rootPath, "\")
	For i = 0 To UBound(aFileName)
		If aFileName(i) <> "" Then
			addPath(aFileName(i))
		End If
	Next
End Sub

Sub addPath(folderName)
	iCrtPathLength = iCrtPathLength + 1
	If iCrtPathLength = 1 Then
		sCrtPath = folderName
	Else
		sCrtPath = sCrtPath & "\" & folderName
	End If

	Call addExpPath(ID_DIV_EXP_PATH, folderName, iCrtPathLength)
End Sub

Sub delPath(iCut)
	iCrtPathLength = iCrtPathLength - iCut
	Dim i, iInStrRev
	For i = 1 To iCut
		iInStrRev = InStrRev(sCrtPath, "\")
		If iInStrRev > 0 Then
			sCrtPath = Mid(sCrtPath, 1, iInStrRev - 1)
		Else
			sCrtPath = ""
		End If
	Next
	Call delExpPath(ID_DIV_EXP_PATH, iCut)
End Sub

Sub removeFile()
	Call removeAllButton(ID_DIV_EXP_FILE)
End Sub

Sub addFile()
	Call vaSubFolderName.ResetArray()
	Call vaSubFileName.ResetArray()
	Call getSubFolderAndFileName()
	Dim i
	For i = 0 To vaSubFolderName.Bound
		Call addButtonOfFolder(ID_DIV_EXP_FILE, vaSubFolderName.V(i))
	Next
	For i = 0 To vaSubFileName.Bound
		Call addButtonOfFile(ID_DIV_EXP_FILE, vaSubFileName.V(i))
	Next
End Sub

Sub openFolder(path)
	If oFso.FolderExists(path) Then
		oWs.Run "explorer.exe " & path
	End If
End Sub

Sub openFile(path)
	If oFso.FileExists(path) Then
		oWs.Run mTextEditorPath & " " & path
	End If
End Sub



Function isValidRootPath(sPath)
	If oFso.FolderExists(sPath) Then
		If Right(sPath, 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)
		isValidRootPath = True
	Else
		MsgBox("path is not exist!")
		isValidRootPath = False
	End If
End Function

Sub getSubFolderAndFileName()
	Dim oRootFolder : Set oRootFolder = oFso.GetFolder(sCrtPath)
	Dim subFolder, subFile

	For Each subFolder In oRootFolder.SubFolders
		vaSubFolderName.Append(subFolder.Name)
	Next
	vaSubFolderName.SortArray()

	For Each subFile In oRootFolder.Files
		vaSubFileName.Append(subFile.Name)
	Next
	vaSubFileName.SortArray()
End Sub