Option Explicit

Function relpaceSlashInPath(path)
    relpaceSlashInPath = Replace(path, "\", "/")
End Function

Function relpaceBackSlashInPath(path)
    relpaceBackSlashInPath = Replace(path, "/", "\")
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

Function trimStr(str)
    trimStr = Replace(Replace(Trim(str), VbCr, ""), VbLf, "")
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

Function getParentPath(filePath)
    Dim str, index
    str = relpaceSlashInPath(filePath)
    index = InStrRev(str, "/")
    Do While index > 0 And index = Len(str)
        str = Left(str, Len(str) - 1)
        index = InStrRev(str, "/")
    Loop

    If index > 0 And index < Len(str) Then
        str = Left(str, index - 1)
    Else
        str = ""
    End If
    getParentPath = str
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
    If InStr(str, "-MMI") > 0 Then
        str = Replace(str, "-MMI", "")
    ElseIf InStr(str, "MMI") > 0 And InStr(str, "-") > 0 Then
        str = Left(str, InStrRev(str, "-") - 1)
    Else
        str = mmiFolderName
    End If
    getDriverProjectName = str
End Function

Function getOriginPathFromOverlayPath(path)
    If InStr(path, "/alps/") > 0 Then
        getOriginPathFromOverlayPath = Right(path, Len(path) - InStr(path, "/alps/") - Len("/alps/") + 1)
    Else
        getOriginPathFromOverlayPath = path
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

Function getTaskNum(project)
    Dim arr, str, taskNum
    taskNum = 1
    If project <> "" Then
        arr = Split(Replace(project, "-", "_"), "_")
        For Each str In arr
            If isNumeric(str) And str > taskNum Then
                taskNum = str
            End If
        Next
    End If
    getTaskNum = trimStr(taskNum)
End Function

Function checkBackslash(str)
    str = Replace(str, "/", "\/")
    str = Replace(str, "[", "\[")
    str = Replace(str, "]", "\]")
    str = Replace(str, ".", "\.")
    str = Replace(str, "\.*", ".*")
    checkBackslash = str
End Function
