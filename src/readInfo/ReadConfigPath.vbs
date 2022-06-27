Option Explicit

Dim mTextEditorPath, mBeyondComparePath, mBrowserPath

Sub readConfigText(DictPath)
    If Not oFso.FileExists(DictPath) Then Exit Sub
    
    Dim oText, sReadLine
    Set oText = oFso.OpenTextFile(DictPath, FOR_READING)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        If InStr(sReadLine, "TextEditorPath") > 0 Then
            mTextEditorPath = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
        ElseIf InStr(sReadLine, "BeyondComparePath") > 0 Then
            mBeyondComparePath = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
        ElseIf InStr(sReadLine, "BrowserPath") > 0 Then
            mBrowserPath = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
        End If
    Loop

    oText.Close
    Set oText = Nothing
End Sub
