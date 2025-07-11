Option Explicit

Dim oWs
Set oWs=CreateObject("wscript.shell")

Sub runPath(path)
    Dim p : p = relpaceBackSlashInPath(checkDriveSdkPath(path))
    Dim success : success = False
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
    If success And p <> relpaceBackSlashInPath(path) Then Call saveHistoryPath(relpaceSlashInPath(path))
End Sub

Sub runTextPath(path)
    Dim p : p = checkDriveSdkPath(path)
    If isFileExists(p) Then
        oWs.Run mTextEditorPath & " " & p
    Else
        MsgBox("not found :" & Vblf & p)
    End If
End Sub

Sub runFolderPath(path)
    Dim p : p = checkDriveSdkPath(path)
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
    Dim result, count
    result = -1
    count = 0
    Do Until result = 0 Or count > 5
        result = oWs.Run("MsHta vbscript:ClipBoardData.setData(""Text"",""" & str & """)(close)",0,True)
        count = count + 1
    Loop
End Sub

Sub CopyOpenPathAllText()
    oWs.SendKeys "+{TAB}"
    oWs.SendKeys "^a"
    oWs.SendKeys "^x"
End Sub
