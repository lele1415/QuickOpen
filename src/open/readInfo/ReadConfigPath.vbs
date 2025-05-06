Option Explicit

Dim mTextEditorPath, mBeyondComparePath, mBrowserPath

Sub readConfigText()
    If Not isFileExists(pConfigText) Then Exit Sub
    
    Dim oText, sReadLine
    Set oText = oFso.OpenTextFile(pConfigText, FOR_READING, False, True)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        If InStr(sReadLine, "TextEditorPath") > 0 And InStr(sReadLine, "=") > 0 Then
            mTextEditorPath = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
        ElseIf InStr(sReadLine, "BeyondComparePath") > 0 And InStr(sReadLine, "=") > 0 Then
            mBeyondComparePath = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
        ElseIf InStr(sReadLine, "BrowserPath") > 0 And InStr(sReadLine, "=") > 0 Then
            mBrowserPath = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
        End If
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub checkConfigInfos()
    Dim oText
    Dim count : count = 0
    If Not isFileExists(Replace(mTextEditorPath, """", "")) Then
        Do Until (isFileExists(mTextEditorPath) Or count > 5)
            mTextEditorPath = InputBox("Text editor path : ", "Please input")
            count = count + 1
        Loop

        mTextEditorPath = """" & mTextEditorPath & """"

        Call initTxtFile(pConfigText)
        Set oText = oFso.OpenTextFile(pConfigText, FOR_APPENDING, False, True)
        oText.WriteLine("TextEditorPath=" & mTextEditorPath)
        oText.WriteLine("BeyondComparePath=" & mBeyondComparePath)
        oText.WriteLine("BrowserPath=" & mBrowserPath)

        oText.Close
    Set oText = Nothing
    End If
End Sub
