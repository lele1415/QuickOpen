Const WINDOW_WIDTH = 460
Const WINDOW_HEIGHT = 850
Sub Window_OnLoad
    Dim ScreenWidth : ScreenWidth = CreateObject("HtmlFile").ParentWindow.Screen.AvailWidth
    Dim ScreenHeight : ScreenHeight = CreateObject("HtmlFile").ParentWindow.Screen.AvailHeight
    Window.MoveTo ScreenWidth - WINDOW_WIDTH ,(ScreenHeight - WINDOW_HEIGHT) \ 3
    Window.ResizeTo WINDOW_WIDTH, WINDOW_HEIGHT
End Sub

Const FOR_READING = 1
Const FOR_APPENDING = 8

Set oWs=CreateObject("wscript.shell")
Set oFso=CreateObject("Scripting.FileSystemObject")



Const SEARCH_FILE = 0
Const SEARCH_FOLDER = 1
Const SEARCH_ROOT = 0
Const SEARCH_SUB = 1
Const SEARCH_WHOLE_NAME = 0
Const SEARCH_PART_NAME = 1
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

        Case searchType = SEARCH_FILE And searchWhere = SEARCH_SUB
            For Each Folder In oRootFolder.SubFolders
                sTmp = startSearch(Folder.Files, pRootFolder & "\" & Folder.Name, str, searchMode, SEARCH_ONE, returnType)
                If sTmp <> "" Then
                    searchFolder = sTmp
                    Exit Function
                End If
            Next
            searchFolder = ""

        Case searchType = SEARCH_FILE And searchWhere = SEARCH_SUB
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
    Call disableElement(ID_SHOW_OR_HIDE_SHORTCUTS)
End Sub

Sub unfreezeAllInput()
    Call enableElement(getWorkInputId())
    Call enableElement(getSdkPathInputId())
    Call enableElement(getProductInputId())
    Call enableElement(getProjectInputId())
    Call enableElement(ID_CREATE_SHORTCUTS)
    Call enableElement(ID_SHOW_OR_HIDE_SHORTCUTS)
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
