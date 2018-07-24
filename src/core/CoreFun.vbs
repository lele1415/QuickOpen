Const WINDOW_WIDTH = 420
Const WINDOW_HEIGHT = 800
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

Function initTxtFile(FilePath)
    If oFso.FileExists(FilePath) Then
        Dim TxtFile
        Set TxtFile = oFso.getFile(FilePath)
        TxtFile.Delete
        Set TxtFile = Nothing
    End If    
    oFso.CreateTextFile FilePath, True
End Function