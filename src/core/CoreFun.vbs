Option Explicit

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

Function searchFolder(pRootFolder, str, searchType, searchWhere, searchMode, searchTimes, returnType)
    If Not oFso.FolderExists(pRootFolder) Then searchFolder = "" : Exit Function
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
                vaAll.Append(f)
            Next
            For Each f In oRootFolder.Files
                vaAll.Append(f)
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
                            vaStr.Append(pRootFolder & "\" & oSingle.Name)
                        Else
                            vaStr.Append(oSingle.Name)
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
                If StrComp(name ,str) = 0 Then
                    checkSearchName = True
                Else
                    checkSearchName = False
                End If
            ELseIf searchMode = SEARCH_PART_NAME Then
                If InStr(name ,str) > 0 Then
                    checkSearchName = True
                Else
                    checkSearchName = False
                End If
            ELseIf searchMode = SEARCH_START_NAME Then
                If str = "" Then
                    checkSearchName = True
                ElseIf StrComp(Mid(name ,1, Len(str)), str) = 0 Then
                    checkSearchName = True
                Else
                    checkSearchName = False
                End If
            End If
        End Function

Function getElementValue(elementId)
    getElementValue = document.getElementById(elementId).value
End Function

Sub setElementValue(elementId, str)
    document.getElementById(elementId).value = str
End Sub

Sub disableElement(inputId)
    document.getElementById(inputId).disabled = "disabled"
End Sub

Sub enableElement(inputId)
    document.getElementById(inputId).disabled = ""
End Sub

Function initTxtFile(FilePath)
    If oFso.FileExists(FilePath) Then
        Dim TxtFile
        Set TxtFile = oFso.getFile(FilePath)
        TxtFile.Delete
        Set TxtFile = Nothing
    End If    
    oFso.CreateTextFile FilePath, True
End Function

Sub freezeAllInput()
    Call disableElement(getWorkInputId())
    Call disableElement(getSdkPathInputId())
    Call disableElement(getProductInputId())
    Call disableElement(getProjectInputId())
    Call disableElement(ID_CREATE_SHORTCUTS)
    Call disableElement(ID_SHOW_SHORTCUTS)
    Call disableElement(ID_HIDE_SHORTCUTS)
End Sub

Sub unfreezeAllInput()
    Call enableElement(getWorkInputId())
    Call enableElement(getSdkPathInputId())
    Call enableElement(getProductInputId())
    Call enableElement(getProjectInputId())
    Call enableElement(ID_CREATE_SHORTCUTS)
    Call enableElement(ID_SHOW_SHORTCUTS)
    Call enableElement(ID_HIDE_SHORTCUTS)
End Sub

Function readTextAndGetValue(keyStr, filePath)
    If Not oFso.FileExists(filePath) Then Exit Function
    
    Dim oText, sReadLine, flag
    flag = False
    Set oText = oFso.OpenTextFile(filePath, FOR_READING)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        Do While InStr(sReadLine, "  ") > 0
            sReadLine = Replace(sReadLine, "  ", " ")
        Loop
        If InStr(sReadLine, " =") > 0 Then
            sReadLine = Replace(sReadLine, " =", "=")
        End If

        If InStr(sReadLine, keyStr) > 0 Then
            readTextAndGetValue = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
            flag = True
            Exit Do
        End If
    Loop

    oText.Close
    Set oText = Nothing
    If Not flag Then readTextAndGetValue = ""
End Function

Function isEndWith(str, endStr)
    If Right(str, Len(endStr)) = endStr Then
        isEndWith = True
    Else
        isEndWith = False
    End If
End Function

Function findStr(str, key)
    Dim i, num
    num = 0
    For i = 1 To (Len(str) - Len(key) + 1)
        If Mid(str, i, Len(key)) = key Then
            num = num + 1
        End If
    Next
    findStr = num
End Function

Function getFileNameFromPath(path)
    Dim str
    str = Replace(path, "\", "/")
    If InStr(str, "/") > 0 Then
        str = Replace(str, Left(str, InStrRev(str, "/")), "")
    Else
        str = path
    End If
    getFileNameFromPath = str
End Function

Function getFolderPath(filePath)
    Dim str
    str = Replace(filePath, "\", "/")
    If InStr(str, "/") > 0 Then
        str = Left(str, InStrRev(str, "/") - 1)
    Else
        str = ""
    End If
    getFolderPath = str
End Function

Function getTabStr()
    Dim folderPath
    folderPath = mIp.Infos.Sdk & "/" & getFolderPath(getOpenPath())
    If Not oFso.FolderExists(folderPath) Then getTabStr = "" : Exit Function
    
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

Function getSameStartStrFromArray(folderPath, input, vaStr)
    If vaStr.Bound = -1 Then getSameStartStrFromArray = "" : Exit Function
    If vaStr.Bound = 0 Then
        If oFso.FolderExists(folderPath & "/" & vaStr.V(0)) Then
            getSameStartStrFromArray = vaStr.V(0) & "/" : Exit Function
        Else
            getSameStartStrFromArray = vaStr.V(0) : Exit Function
        End If
    End If

    Dim s1, s2, i, str
    s1 = vaStr.V(0)
    For i = 0 To vaStr.Bound - 1
        s2 = vaStr.V(i + 1)
        If InStr(s2, s1) <> 1 Then s1 = getSameStartStr(input, s1, s2)
    Next
    getSameStartStrFromArray = s1
End Function

Function getSameStartStr(key, s1, s2)
    Dim minLen
    minLen = Len(s1)
    If Len(s2) < minLen Then minLen = Len(s2)
    If minLen = Len(key) Then getSameStartStr = key : Exit Function

    Dim str, c1, c2, i
    str = key
    For i = Len(key) + 1 To minLen
        c1 = Mid(s1, i, 1)
        c2 = Mid(s2, i, 1)
        If c1 = c2 Then
            str = str & c1
        Else
            Exit For
        End If
    Next
    getSameStartStr = str
End Function

Function getDriverProjectName(mmiFolderName)
    Dim str : str = mmiFolderName

    'M863Y_YUKE_066-MMI
    'm863ur200_64-SBYH_A8005A-Nitro_8_MMI
    If InStr(str, "MMI") > 0 And InStr(str, "-") > 0 Then
        str = Left(str, InStrRev(str, "-") - 1)
    Else
        str = mmiFolderName
    End If
    getDriverProjectName = str
End Function

Function isPictureFilePath(path)
    If isEndWith(path, ".bmp") Then
        isPictureFilePath = True
    ElseIf isEndWith(path, ".png") Then
        isPictureFilePath = True
    ElseIf isEndWith(path, ".jpg") Then
        isPictureFilePath = True
    ElseIf isEndWith(path, ".jpeg") Then
        isPictureFilePath = True
    Else
        isPictureFilePath = False
    End If
End Function

Function isCompressFilePath(path)
    If isEndWith(path, ".zip") Then
        isCompressFilePath = True
    ElseIf isEndWith(path, ".rar") Then
        isCompressFilePath = True
    ElseIf isEndWith(path, ".7z") Then
        isCompressFilePath = True
    ElseIf isEndWith(path, ".tar.gz") Then
        isCompressFilePath = True
    Else
        isCompressFilePath = False
    End If
End Function

Sub runPath(path)
    path = Replace(path, "/", "\")
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

Sub runTextPath(path)
    path = Replace(path, "/", "\")
    If oFso.FileExists(path) Then
        oWs.Run mTextEditorPath & " " & path
    Else
        MsgBox("not found :" & Vblf & path)
    End If
End Sub

Sub runFolderPath(path)
    path = Replace(path, "/", "\")
    If oFso.FolderExists(path) Then
        oWs.Run "explorer.exe " & path
    Else
        MsgBox("not found :" & Vblf & path)
    End If
End Sub

Sub runWebsite(path)
    oWs.Run mBrowserPath & " " & path
End Sub

Sub runBeyondCompare(leftPath, rightPath)
    oWs.Run mBeyondComparePath & " " & leftPath & " " & rightPath
End Sub

Sub CopyString(str)
    If Len(str) > 452 Then
        MsgBox("String is too long!(max length 452)")
        setOpenPath(str)
        Exit Sub
    End If
    oWs.Run "MsHta vbscript:ClipBoardData.setData(""Text"",""" & str & """)(Window.Close)"
End Sub

Sub CopyQuoteString(str)
    If Len(str) > 452 Then
        MsgBox("String is too long!(max length 452)")
        setOpenPath(str)
        Exit Sub
    End If
    oWs.Run "MsHta vbscript:ClipBoardData.setData(""Text""," & str & ")(Window.Close)"
End Sub

Sub onInputListClick(divId, str)
    Dim path
    If InStr(divId, "outfile") > 0 Then
        path = getOutListPath(str)
        If path <> "" Then runPath(path)

    ElseIf InStr(divId, "openbutton") > 0 Then
        path = getOpenButtonListPath(str)
        If path <> "" Then runPath(path)

    ElseIf InStr(divId, "filebutton") > 0 Then
        Call vaFilePathList.ResetArray()
        Call mFileButtonList.removeList()
        Call setOpenPath(str)
    End If
End Sub

Const KEYCODE_ENTER = 13
Const KEYCODE_SPACE = 32
Const KEYCODE_TAB = 9
Function onKeyDown(keyCode)
    If keyCode = KEYCODE_ENTER Then
        Call onOpenButtonClick()

    ElseIf keyCode = KEYCODE_SPACE Then
        Call pasteAndOpenPath()
    
    ElseIf keyCode = KEYCODE_TAB Then
        Call tabOpenPath()

    Else
        onKeyDown = True : Exit Function
    End If
    
    onKeyDown = False
End Function
