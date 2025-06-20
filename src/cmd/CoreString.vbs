Option Explicit

Function isEq(str1, str2)
    If str1 = str2 Then
        isEq = True
    Else
        isEq = False
    End If
End Function

Function isCt(str1, str2)
    If InStr(str1, str2) > 0 Then
        isCt = True
    Else
        isCt = False
    End If
End Function

Function isInStr(str1, str2)
    If InStr(str1, str2) > 0 Then
        isInStr = True
    Else
        isInStr = False
    End If
End Function

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

Function isTaskNum(taskNum)
    If isNumeric(taskNum) And Len(taskNum) < 5 Then
        isTaskNum = True
    Else
        isTaskNum = False
    End If
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

Function getParentPath(path)
    Dim str, index
    str = relpaceSlashInPath(path)
    index = InStrRev(str, "/")
    If index > 0 And index < Len(str) Then
        str = Left(str, index)
    End If
    getParentPath = str
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
    If InStr(path, "weibu/") = 1 Then
        If InStr(path, "/alps/") Then
            path = Split(path, "/alps/")(1)
        Else
            path = Right(path, Len(path) - InStr(path, "/"))
            path = Right(path, Len(path) - InStr(path, "/"))
            path = Right(path, Len(path) - InStr(path, "/"))
        End If
    End If
    getOriginPathFromOverlayPath = path
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

Function getPathWithDriveSdk(path)
    Dim newPath, driveSdk
    newPath = relpaceBackSlashInPath(path)
    driveSdk = relpaceBackSlashInPath(mDrive & mBuild.Sdk)
    If InStr(newPath, "..\") = 1 Then
        getPathWithDriveSdk = Left(driveSdk, InStrRev(driveSdk, "\")) & Right(newPath, Len(newPath) - 3)
    Else
        getPathWithDriveSdk = driveSdk & "\" & newPath
    End If
End Function

Function checkDriveSdkPath(path)
    Dim newPath
    newPath = relpaceBackSlashInPath(path)
    If InStr(newPath, ":\") = 0 And InStr(newPath, "\\192.168") = 0 Then
        checkDriveSdkPath = getPathWithDriveSdk(newPath)
    Else
        checkDriveSdkPath = newPath
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

Function checkBackslash(str)
    str = Replace(str, "/", "\/")
    str = Replace(str, "[", "\[")
    str = Replace(str, "]", "\]")
    str = Replace(str, ".", "\.")
    str = Replace(str, "\.*", ".*")
    checkBackslash = str
End Function

Function replaceProjectInfoStr(path)
    If InStr(path, "[vnd]") > 0 Then
        If mBuild.Infos.is8781() Then
            path = Replace(path, "[vnd]", "vext")
        Else
            path = Replace(path, "[vnd]", "vnd")
        End If
    End If

    If InStr(path, "[product]") > 0 Then
        path = Replace(path, "[product]", mBuild.Product)
    End If

    If InStr(path, "[project]") > 0 Then
        path = Replace(path, "[project]", mBuild.Project)
    End If

    If InStr(path, "[boot_logo]") > 0 Then
        path = Replace(path, "[boot_logo]", mBuild.Infos.getBootLogo())
    End If

    replaceProjectInfoStr = path
End Function
