Const ID_DIV_SHOW_AREA_L1 = "show_area_l1"
Const ID_DIV_CURRENT_PATH_L1 = "current_path_l1"
Const ID_DIV_SHOW_AREA_KK = "show_area_kk"
Const ID_DIV_CURRENT_PATH_KK = "current_path_kk"

Dim aDivAreaId : aDivAreaId = Array(ID_DIV_SHOW_AREA_L1, ID_DIV_SHOW_AREA_KK)
Dim aDivPathId : aDivPathId = Array(ID_DIV_CURRENT_PATH_L1, ID_DIV_CURRENT_PATH_KK)
Dim aCurrentPath : aCurrentPath = Array(1)
Dim aCurrentDeepCount : aCurrentDeepCount = Array(1)

Dim vaSubFolderName : Set vaSubFolderName = New VariableArray
Dim vaSubFileName : Set vaSubFileName = New VariableArray


Const FROM_L1 = 0
Const FROM_KK = 1
Sub initPath(where)
	Dim rootPath
	Select Case where
		Case FROM_L1 : rootPath = handlePathFromL1(1)
		Case FROM_KK : rootPath = handlePathFromKK(1)
	End Select
	
	If oFso.FolderExists(rootPath) Then
		If Right(rootPath, 1) = "\" Then rootPath = Mid(rootPath, 1, Len(rootPath)-1)
		Call AddSubFolder(rootPath, -1, aDivAreaId(where), aDivPathId(where), where)
	Else
		MsgBox("path is not exist!")
	End If
End Sub

Sub AddSubFolder(folderName, nextDeepCount, divAreaId, divPathId, where)
	If nextDeepCount = aCurrentDeepCount(where) Then Exit Sub

	Call updateCurrentPath(folderName, nextDeepCount, divAreaId, divPathId, where)
	Call resetShowArea(divAreaId)
	Call getSubFolderAndFileName(where)

	Dim i
	For i = 0 To vaSubFolderName.Length
		Call addButtonOfFolder(divAreaId, vaSubFolderName.Value(i), divPathId, where)
	Next
	For i = 0 To vaSubFileName.Length
		Call addButtonOfFile(divAreaId, vaSubFileName.Value(i), where)
	Next
End Sub

Sub updateCurrentPath(folderName, nextDeepCount, divAreaId, divPathId, where)
	If nextDeepCount = -1 Then
	'//when change root path
		aCurrentPath(where) = folderName
		Call removeButtonOfCurrentPath(divPathId, aCurrentDeepCount(where), 0)
		Call addButtonOfCurrentPath(divPathId, folderName, aCurrentDeepCount(where), divAreaId, where)
		aCurrentDeepCount(where) = 1
	ElseIf nextDeepCount = 0 Then
	'//when show sub folder
		aCurrentDeepCount(where) = aCurrentDeepCount(where) + 1
		aCurrentPath(where) = aCurrentPath(where) & "\" & folderName
		Call addButtonOfCurrentPath(divPathId, folderName, aCurrentDeepCount(where), divAreaId, where)
	Else
	'//when show past folder in path
		Dim i
		For i = 1 To aCurrentDeepCount(where) - nextDeepCount
			aCurrentPath(where) = removeLastFolderNameOfPath(aCurrentPath(where))
		Next

		Call removeButtonOfCurrentPath(divPathId, aCurrentDeepCount(where), nextDeepCount)
		aCurrentDeepCount(where) = nextDeepCount
	End If
End Sub

Function removeLastFolderNameOfPath(path)
	Dim iInStr : iInStr = InStrRev(path, "\")
	removeLastFolderNameOfPath = Mid(path, 1, iInStr-1)
End Function

Sub resetShowArea(divAreaId)
	Call removeAllButton(divAreaId)
	Call vaSubFolderName.ResetArray()
	Call vaSubFileName.ResetArray()
End Sub

Sub getSubFolderAndFileName(where)
	Dim oRootFolder : Set oRootFolder = oFso.GetFolder(aCurrentPath(where))
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

Sub OpenFolder(folderName, where)
	runOpenPath(aCurrentPath(where) & "\" & folderName)
End Sub

Sub runOpenPath(path)
	If oFso.FolderExists(path) Then
		oWs.Run "explorer.exe " & path
	ElseIf oFso.FileExists(path) Then
		oWs.Run """" & PATH_TEXT_EDITOR & """" & " " & path
	Else
		MsgBox("not found :" & Vblf & path)
	End If
End Sub
