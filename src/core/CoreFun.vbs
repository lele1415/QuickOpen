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

Function searchFolder(path, str, searchType, searchWhere, searchMode, searchTimes, returnType)
    Dim pRootFolder : pRootFolder = checkDriveSdkPath(path)
    If Not isFolderExists(pRootFolder) Then searchFolder = "" : Exit Function
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

Sub hideElement(elementId)
    document.getElementById(elementId).style.display = "none"
End Sub

Sub showElement(elementId)
    document.getElementById(elementId).style.display = "block"
End Sub

Function initTxtFile(FilePath)
    If isFileExists(FilePath) Then
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

Sub startCmdMode()
    Call hideElement("work_fieldset")
    Call hideElement("project_fieldset")
    Call hideElement("openpath_fieldset")
    Call hideElement("openpath2_fieldset")
    Call hideElement("explorer_fieldset")
    Call hideElement("commands_fieldset")
    Call hideElement("br1")
    Call hideElement("br2")
    Call hideElement("br3")
    Call hideElement("br4")
    Call setCmdSmallWindow()
    Call setCmdTextClass(getCmdInputId(), getOpenPathInputId(), "cmd_text")
End Sub

Sub exitCmdMode()
    Call showElement("work_fieldset")
    Call showElement("project_fieldset")
    Call showElement("openpath_fieldset")
    Call showElement("openpath2_fieldset")
    Call showElement("explorer_fieldset")
    Call showElement("commands_fieldset")
    Call showElement("br1")
    Call showElement("br2")
    Call showElement("br3")
    Call showElement("br4")
    Call setDefaultWindow()
    Call setCmdTextClass(getCmdInputId(), getOpenPathInputId(), "textarea_text")
End Sub

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
    If Not isFileExists(path) Then Exit Function
    
    Dim oText, sReadLine, flag
    flag = False
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
            flag = True
            Exit Do
        End If
    Loop

    oText.Close
    Set oText = Nothing
    If Not flag Then readTextAndGetValue = ""
End Function

Function relpaceSlashInPath(path)
    relpaceSlashInPath = Replace(path, "\", "/")
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
    str = relpaceSlashInPath(path)
    If InStr(str, "/") > 0 Then
        str = Replace(str, Left(str, InStrRev(str, "/")), "")
    Else
        str = path
    End If
    getFileNameFromPath = str
End Function

Function getFolderPath(filePath)
    Dim str
    str = relpaceSlashInPath(filePath)
    If InStr(str, "/") > 0 Then
        str = Left(str, InStrRev(str, "/") - 1)
    Else
        str = ""
    End If
    getFolderPath = str
End Function

Function getTabStr()
    Dim folderPath
    folderPath = getFolderPath(getOpenPath())
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

Function getSameStartStrFromArray(folderPath, input, vaStr)
    If vaStr.Bound = -1 Then getSameStartStrFromArray = "" : Exit Function
    If vaStr.Bound = 0 Then
        If isFolderExists(folderPath & "/" & vaStr.V(0)) Then
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

Sub saveHistoryPath(path)
    Dim index
    index = vaPathHistory.GetIndexIfExist(path)
    If index > -1 Then Call vaPathHistory.MoveToEnd(index) : Exit Sub

    If vaPathHistory.Bound = 9 Then Call vaPathHistory.PopBySeq(9)
    Call vaPathHistory.Append(path)
End Sub

Sub showHistoryPath(keyCode)
    If vaPathHistory.Bound = -1 Then Exit Sub

    Dim index
    index = vaPathHistory.GetIndexIfExist(getOpenPath())
    If index = -1 Then mCurrentPath = getOpenPath()

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
            Call setOpenPath(mCurrentPath)
        End If
    End If
End Sub

Function checkDriveSdkPath(path)
    If InStr(path, ":\") = 0 And InStr(path, "\\192.168") = 0 Then
        checkDriveSdkPath = mIp.Infos.getPathWithDriveSdk(path)
    Else
        checkDriveSdkPath = path
    End If
End Function

Sub runPath(path)
    Dim p : p = checkDriveSdkPath(path)
    Dim success : success = False
    p = Replace(p, "/", "\")
    If isFolderExists(p) Then
        oWs.Run "explorer.exe " & p
        success = True
    ElseIf isFileExists(p) Then
        If isPictureFilePath(p) Or isCompressFilePath(p) Then
            oWs.Run "explorer.exe " & p
        Else
            oWs.Run mTextEditorPath & " " & p
        End If
        success = True
    Else
        MsgBox("not found :" & Vblf & p)
    End If
    If success And InStr(p, mIp.Infos.DriveSdk) > 0 Then Call saveHistoryPath(mIp.cutProject(p))
End Sub

Sub runTextPath(path)
    Dim p : p = checkDriveSdkPath(path)
    p = Replace(p, "/", "\")
    If isFileExists(p) Then
        oWs.Run mTextEditorPath & " " & p
    Else
        MsgBox("not found :" & Vblf & p)
    End If
End Sub

Sub runFolderPath(path)
    Dim p : p = checkDriveSdkPath(path)
    p = Replace(p, "/", "\")
    If isFolderExists(p) Then
        oWs.Run "explorer.exe " & p
    Else
        MsgBox("not found :" & Vblf & p)
    End If
End Sub

Sub runWebsite(path)
    oWs.Run mBrowserPath & " " & path
End Sub

Sub runBeyondCompare(leftPath, rightPath)
    Dim lp : lp = checkDriveSdkPath(leftPath)
    Dim rp : rp = checkDriveSdkPath(rightPath)
    oWs.Run mBeyondComparePath & " " & lp & " " & rp
End Sub

Sub CopyString(str)
    If Len(str) > 452 Then
        'MsgBox("String is too long!(max length 452)")
        setOpenPath(Replace(str, "&Chr(34)&""", ""))
        Call CopyOpenPathAllText()
        Exit Sub
    End If
    oWs.Run "MsHta vbscript:ClipBoardData.setData(""Text"",""" & str & """)(Window.Close)"
End Sub

Sub CopyOpenPathAllText()
    oWs.SendKeys "+{TAB}"
    oWs.SendKeys "^a"
    oWs.SendKeys "^x"
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
        Call mFileButtonList.removeList()
        Call setOpenPath(str)
    End If
End Sub

Function changeFocus(keyCode)
    If mFileButtonList.changeListFocus(keyCode) Then changeFocus = True : Exit Function
    If mOpenButtonList.changeListFocus(keyCode) Then changeFocus = True : Exit Function
    changeFocus = False
End Function

Const KEYCODE_ENTER = 13
Const KEYCODE_SPACE = 32
Const KEYCODE_TAB = 9
Const KEYCODE_UP = 38
Const KEYCODE_DOWN = 40
Const KEYCODE_ESC = 27
Function onKeyDown(keyCode)
    If keyCode = KEYCODE_ENTER Then
        If mFileButtonList.isShowing() Then
            Call mFileButtonList.clickFocusedLi()
        ElseIf mOpenButtonList.isShowing() Then
            Call mOpenButtonList.clickFocusedLi()
        Else
            Call onOpenButtonClick()
        End If

    ElseIf keyCode = KEYCODE_ESC Then
        Call mFileButtonList.hideListIfShowing()
        Call mOpenButtonList.hideListIfShowing()

    ElseIf keyCode = KEYCODE_SPACE Then
        Call pasteAndOpenPath()
    
    ElseIf keyCode = KEYCODE_TAB Then
        Call tabOpenPath()
    
    ElseIf keyCode = KEYCODE_UP Or keyCode = KEYCODE_DOWN Then
        If changeFocus(keyCode) Then onKeyDown = False : Exit Function
        Call showHistoryPath(keyCode)

    Else
        onKeyDown = True : Exit Function
    End If
    
    onKeyDown = False
End Function

Sub setDrive(drive)
    If (drive = "z") Then
        mDrive = "Z:\work05\"
    ElseIf (drive = "z6") Then
        mDrive = "Z:\work06\"
    ElseIf (drive = "x") Then
        mDrive = "X:\work2\"
    End If
    
    document.title = getElementValue(getWorkInputId()) & " " & mDrive
End Sub

Sub setSdk(sdk)
    If sdk = "8766s" Then
        mIp.Sdk = "mt8766_s\alps"
        Call onSdkPathChange()
    ElseIf sdk = "8168s" Then
        mIp.Sdk = "mt8168_s\alps"
        Call onSdkPathChange()
    ElseIf sdk = "8766r" Then
        If mDrive = "X:\work2\" Then
            mIp.Sdk = "mt8766_r\alps"
            Call onSdkPathChange()
        ElseIf mDrive = "Z:\work05\" Then
            mIp.Sdk = "mt8766_r\alps2"
            Call onSdkPathChange()
        End If
    ElseIf sdk = "8766r" Then
        mIp.Sdk = "mt8168_r\alps"
        Call onSdkPathChange()
    End If
End Sub

Sub applyProjectPath()
    Dim aInfos : aInfos = Split(getOpenPath(), "/")
    If UBound(aInfos) < 2 Then MsgBox("Not valid project path!") : Exit Sub

    If isFolderExists("weibu/" & aInfos(1) & "/" & aInfos(2)) Then
        mIp.Product = aInfos(1)
        Call onProductChange()
        mIp.Project = aInfos(2)
        Call onProjectChange()
    End if
End Sub

Function strExistInFile(filePath, str)
    Dim oText, path, sLine
    path = checkDriveSdkPath(filePath)
    Set oText = oFso.OpenTextFile(path, FOR_READING)

    Do Until oText.AtEndOfStream
        sLine = oText.ReadLine
        If InStr(sLine, str) > 0 Then
            strExistInFile = True : Exit Function
        End If
    Loop

    strExistInFile = False
End Function

Function checkProjectExist(sdk, product, project)
    If Not checkDrive(sdk, product, project) Then
        Call setDrive("x")
        If Not checkDrive(sdk, product, project) Then
            Call setDrive("z")
            If Not checkDrive(sdk, product, project) Then
                Call setDrive("z6")
                If Not checkDrive(sdk, product, project) Then 
                    MsgBox("Not exist: " & path)
                    checkProjectExist = False
                    Exit Function
                End If
            End If
        End If
    End If

    checkProjectExist = True
End Function

Function checkDrive(sdk, product, project)
    Dim path
	path = mDrive & sdk
	If isFolderExists(path) Then
	    path = path & "\weibu\" & product & "\" & project
		If isFolderExists(path) Then
		    checkDrive = True
			Exit Function
		ENd If
	End If

    checkDrive = False
End Function

Function getTaskNum(project)
    Dim arr, str
    If project <> "" Then
        arr = Split(Replace(project, "-", "_"), "_")
        For Each str In arr
            If isNumeric(str) And Len(str) < 4 Then
                getTaskNum = str
                Exit Function
            End If
        Next
    End If
    getTaskNum = "1"
End Function

Function isT0Sdk()
    If InStr(mIp.Infos.Sdk, "_t0\") > 0 Then
        isT0Sdk = True
    Else
        isT0Sdk = False
    End If
End Function

Function isT0SdkSys()
    If InStr(mIp.Infos.Sdk, "_t0\sys") > 0 Then
        isT0SdkSys = True
    Else
        isT0SdkSys = False
    End If
End Function

Function isT0SdkVnd()
    If InStr(mIp.Infos.Sdk, "_t0\vnd") > 0 Then
        isT0SdkVnd = True
    Else
        isT0SdkVnd = False
    End If
End Function

Sub setT0SdkSys()
    mIp.Sdk = Replace(mIp.Infos.Sdk, "vnd", "sys")
    mIp.Product = mIp.Infos.SysTarget
    If InStr(mIp.Infos.Project, "-MMI") = 0 Then mIp.Project = mIp.Infos.Project & "-MMI"
    Call createWorkName()
End Sub

Sub setT0SdkVnd()
    mIp.Sdk = Replace(mIp.Infos.Sdk, "sys", "vnd")
    mIp.Product = mIp.Infos.VndTarget
    mIp.Project = Replace(mIp.Infos.Project , "-MMI", "")
    Call createWorkName()
End Sub

Function checkBackslash(str)
    str = Replace(str, "/", "\/")
    str = Replace(str, "[", "\[")
    str = Replace(str, "]", "\]")
    str = Replace(str, ".", "\.")
    str = Replace(str, "\.*", ".*")
    checkBackslash = str
End Function
