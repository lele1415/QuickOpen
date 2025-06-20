Option Explicit

Const FOR_READING = 1
Const FOR_APPENDING = 8

Const SEARCH_FILE = 0
Const SEARCH_FOLDER = 1
Const SEARCH_FILE_FOLDER = 2
Const SEARCH_ROOT = 0
Const SEARCH_SUB = 1
Const SEARCH_WHOLE_NAME = 0
Const SEARCH_PART_NAME = 1
Const SEARCH_START_NAME = 2
Const SEARCH_ONE = 0
Const SEARCH_ALL = 1
Const SEARCH_RETURN_PATH = 0
Const SEARCH_RETURN_NAME = 1

Dim oFso
Set oFso=CreateObject("Scripting.FileSystemObject")

Dim PATH_CONFIG : PATH_CONFIG = oWs.CurrentDirectory & "\res\config.ini"
Dim PATH_TASK_LIST : PATH_TASK_LIST = oWs.CurrentDirectory & "\res\tasklist.ini"
Dim PATH_LAST_TASK : PATH_LAST_TASK = oWs.CurrentDirectory & "\res\last_task.ini"
Dim vaCmdHistory : Set vaCmdHistory = New VariableArray
Dim vaPathHistory : Set vaPathHistory = New VariableArray

Function searchFolder(path, str, searchType, searchWhere, searchMode, searchTimes, returnType)
    Dim pRootFolder : pRootFolder = checkDriveSdkPath(path)
    If Not isFolderExists(pRootFolder) Then
        MsgBox("Path does not exist!" & VbLf & pRootFolder)
        If searchTimes = SEARCH_ONE Then
            searchFolder = ""
        Else
            Set searchFolder = new VariableArray
        End If
        Exit Function
    End If
    If searchMode = SEARCH_WHOLE_NAME Then searchTimes = SEARCH_ONE

    Dim oRootFolder : Set oRootFolder = oFso.GetFolder(pRootFolder)

    Dim Folder, sTmp
    Select Case True
        Case searchType = SEARCH_FILE And searchWhere = SEARCH_ROOT
            If searchTimes = SEARCH_ALL Then
                Set searchFolder = startSearch(oRootFolder.Files, pRootFolder, str, searchMode, SEARCH_ALL, returnType)
            Else
                searchFolder = startSearch(oRootFolder.Files, pRootFolder, str, searchMode, SEARCH_ONE, returnType)
            End If

        Case searchType = SEARCH_FOLDER And searchWhere = SEARCH_ROOT
            If searchTimes = SEARCH_ALL Then
                Set searchFolder = startSearch(oRootFolder.SubFolders, pRootFolder, str, searchMode, SEARCH_ALL, returnType)
            Else
                searchFolder = startSearch(oRootFolder.SubFolders, pRootFolder, str, searchMode, SEARCH_ONE, returnType)
            End If

        Case searchType = SEARCH_FILE_FOLDER And searchWhere = SEARCH_ROOT
            Dim vaAll, f
            Set vaAll = New VariableArray
            For Each f In oRootFolder.SubFolders
                vaAll.append(f)
            Next
            For Each f In oRootFolder.Files
                vaAll.append(f)
            Next
            If searchTimes = SEARCH_ALL Then
                Set searchFolder = startSearch(vaAll.InnerArray, pRootFolder, str, searchMode, SEARCH_ALL, returnType)
            Else
                searchFolder = startSearch(vaAll.InnerArray, pRootFolder, str, searchMode, SEARCH_ONE, returnType)
            End If

        Case searchType = SEARCH_FILE And searchWhere = SEARCH_SUB
            For Each Folder In oRootFolder.SubFolders
                sTmp = startSearch(Folder.Files, pRootFolder & "\" & Folder.Name, str, searchMode, SEARCH_ONE, returnType)
                If sTmp <> "" Then
                    searchFolder = sTmp
                    Exit Function
                End If
            Next
            searchFolder = ""

        Case searchType = SEARCH_FOLDER And searchWhere = SEARCH_SUB
            For Each Folder In oRootFolder.SubFolders
                sTmp = startSearch(Folder.SubFolders, pRootFolder & "\" & Folder.Name, str, searchMode, SEARCH_ONE, returnType)
                If sTmp <> "" Then
                    searchFolder = sTmp
                    Exit Function
                End If
            Next
            searchFolder = ""
    End Select
End Function

Function startSearch(oAll, pRootFolder, str, searchMode, searchTimes, returnType)
    Dim oSingle

    If searchTimes = SEARCH_ALL Then
        Dim vaStr : Set vaStr = New VariableArray
        For Each oSingle In oAll
            If checkSearchName(oSingle.Name, str, searchMode) Then
                If returnType = SEARCH_RETURN_PATH Then
                    vaStr.append(pRootFolder & "\" & oSingle.Name)
                Else
                    vaStr.append(oSingle.Name)
                End If
            End If
        Next
        Set startSearch = vaStr
        Exit Function
    Else
        For Each oSingle In oAll
            If checkSearchName(oSingle.Name, str, searchMode) Then
                If returnType = SEARCH_RETURN_PATH Then
                    startSearch = pRootFolder & "\" & oSingle.Name
                Else
                    startSearch = oSingle.Name
                End If
                Exit Function
            End If
        Next
    End If
    startSearch = ""
End Function

Function checkSearchName(name, str, searchMode)
    If searchMode = SEARCH_WHOLE_NAME Then
        If StrComp(name, str) = 0 Then
            checkSearchName = True
        Else
            checkSearchName = False
        End If
    ELseIf searchMode = SEARCH_PART_NAME Then
        If InStr(name, str) > 0 Then
            checkSearchName = True
        Else
            checkSearchName = False
        End If
    ELseIf searchMode = SEARCH_START_NAME Then
        If str = "" Then
            checkSearchName = True
        ElseIf StrComp(Mid(name, 1, Len(str)), str) = 0 Then
            checkSearchName = True
        Else
            checkSearchName = False
        End If
    End If
End Function

Function initTxtFile(FilePath)
    If isFileExists(FilePath) Then
        Dim TxtFile
        Set TxtFile = oFso.getFile(FilePath)
        TxtFile.Delete
        Set TxtFile = Nothing
    End If    
    oFso.CreateTextFile FilePath, True
End Function

Function isFileExists(path)
    Dim newPath
    newPath = checkDriveSdkPath(path)
    isFileExists = oFso.FileExists(newPath)
End Function

Function isFolderExists(path)
    Dim newPath
    newPath = checkDriveSdkPath(path)
    isFolderExists = oFso.FolderExists(newPath)
End Function

Function readTextAndGetValue(keyStr, filePath)
    Dim path : path = checkDriveSdkPath(filePath)
    If Not isFileExists(path) Then readTextAndGetValue = "" : Exit Function
    
    Dim oText, sReadLine, result
    result = False
    Set oText = oFso.OpenTextFile(path, FOR_READING)

    Do Until oText.AtEndOfStream
        sReadLine = Trim(oText.ReadLine)
        Do While InStr(sReadLine, "  ") > 0
            sReadLine = Replace(sReadLine, "  ", " ")
        Loop
        If InStr(sReadLine, " =") > 0 Then
            sReadLine = Replace(sReadLine, " =", "=")
        End If

        If InStr(sReadLine, keyStr) = 1 Then
            readTextAndGetValue = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
            result = True
            Exit Do
        End If
    Loop

    oText.Close
    Set oText = Nothing
    If Not result Then readTextAndGetValue = ""
End Function

Function readLineOfTextFile(line, filePath)
    Dim path : path = checkDriveSdkPath(filePath)
    If Not isFileExists(path) Then Exit Function
    
    Dim oText, sReadLine, index
    Set oText = oFso.OpenTextFile(path, FOR_READING)
    index = 1

    Do Until oText.AtEndOfStream
        If index = line Then
            sReadLine = Trim(oText.ReadLine)
            Exit Do
        End If
        index = index + 1
    Loop
    readLineOfTextFile = sReadLine
End Function

Function getTabStr()
    Dim folderPath
    folderPath = getParentPath(getOpenPath())
    If Not isFolderExists(folderPath) Then getTabStr = "" : Exit Function
    
    Dim input, vaFileFolder, sameStartStr
    input = getFileNameFromPath(getOpenPath())
    Set vaFileFolder = searchFolder(folderPath, input, SEARCH_FILE_FOLDER, SEARCH_ROOT, SEARCH_START_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)

    sameStartStr = getSameStartStrFromArray(folderPath, input, vaFileFolder)
    If sameStartStr <> "" Then
        getTabStr = Replace(sameStartStr, input, "")
    Else
        getTabStr = sameStartStr
    End If
End Function

Sub saveHistoryPath(path)
    Dim index
    index = vaPathHistory.getIndexIfExist(path)
    If index > -1 Then Call vaPathHistory.moveToEnd(index) : Exit Sub

    If vaPathHistory.Bound = 19 Then Call vaPathHistory.popBySeq(0)
    Call vaPathHistory.append(path)
End Sub

Dim mCurrentHistoryPath
Sub showHistoryPath(keyCode)
    If vaPathHistory.Bound = -1 Then Exit Sub

    Dim index
    index = vaPathHistory.getIndexIfExist(getOriginPathFromOverlayPath(getOpenPath()))
    If index = -1 Then mCurrentHistoryPath = getOpenPath()

    If keyCode = KEYCODE_UP Then
        If index > 0 Then
            Call setOpenPath(vaPathHistory.V(index - 1))
        ElseIf index = -1 Then
            Call setOpenPath(vaPathHistory.V(vaPathHistory.Bound))
        End If
    ElseIf keyCode = KEYCODE_DOWN Then
        If index > -1 And index < vaPathHistory.Bound Then
            Call setOpenPath(vaPathHistory.V(index + 1))
        ElseIf index = vaPathHistory.Bound Then
            Call setOpenPath(mCurrentHistoryPath)
        End If
    End If
End Sub

Sub saveHistoryCmd(cmd)
    Dim index
    index = vaCmdHistory.getIndexIfExist(cmd)
    If index > -1 Then Call vaCmdHistory.moveToEnd(index) : Exit Sub

    If vaCmdHistory.Bound = 19 Then Call vaCmdHistory.popBySeq(0)
    Call vaCmdHistory.append(cmd)
End Sub

Sub showHistoryCmd(keyCode)
    If vaCmdHistory.Bound = -1 Then Exit Sub

    Dim index
    index = vaCmdHistory.getIndexIfExist(getCmdText())
    If index = -1 Then mCurrentHistoryPath = getCmdText()

    If keyCode = KEYCODE_UP Then
        If index > 0 Then
            Call setCmdText(vaCmdHistory.V(index - 1))
        ElseIf index = -1 Then
            Call setCmdText(vaCmdHistory.V(vaCmdHistory.Bound))
        End If
    ElseIf keyCode = KEYCODE_DOWN Then
        If index > -1 And index < vaCmdHistory.Bound Then
            Call setCmdText(vaCmdHistory.V(index + 1))
        ElseIf index = vaCmdHistory.Bound Then
            Call setCmdText(mCurrentHistoryPath)
        End If
    End If
End Sub

Function getParentProject()
    Dim vaParents, project, index, folderPath
    Set vaParents = New VariableArray
    project = mBuild.Project
    Do
        folderPath = "weibu/" & mBuild.Product & "/" & project
        If isFolderExists(folderPath) Then
            vaParents.Append(folderPath)
        End If
        index = InStrRev(project, "-")
        If index > 0 Then project = Left(project, index - 1)
    Loop While index > 0
    Set getParentProject = vaParents
End Function

Function getMultiOverlayPath(filePath)
    Dim configFile
    If mBuild.Infos.hasConfig() Then
        configFile = mBuild.Infos.ProjectPath & "/config/" & getFileNameFromPath(filePath)
        If isFileExists(configFile) Then
            getMultiOverlayPath = configFile
            Exit Function
        End If
    End If

    Dim overlayPath
    overlayPath = filePath
    If isFileExists(filePath) Then
        Dim vaParents, i
        Set vaParents = getParentProject()
        If vaParents.Bound > -1 Then
            For i = 0 To vaParents.Bound
                If isFileExists(vaParents.V(i) & "/alps/" & filePath) Then
                    overlayPath = vaParents.V(i) & "/alps/" & filePath
                    Exit For
                End If
            Next
        End If
    End If
    getMultiOverlayPath = overlayPath
End Function
