Option Explicit

Dim oWs
Set oWs=CreateObject("wscript.shell")

Sub runPath(path)
    Dim p : p = checkDriveSdkPath(path)
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
    If success And InStr(p, mIp.Infos.DriveSdk) > 0 Then Call saveHistoryPath(mIp.cutProject(p))
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
    oWs.Run "MsHta vbscript:ClipBoardData.setData(""Text"",""" & str & """)(Window.Close)"
End Sub

Sub CopyOpenPathAllText()
    oWs.SendKeys "+{TAB}"
    oWs.SendKeys "^a"
    oWs.SendKeys "^x"
End Sub
