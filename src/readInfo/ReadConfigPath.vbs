Dim pConfigText : pConfigText = oWs.CurrentDirectory & "\res\config.ini"

Dim mTextEditorPath, mBeyondComparePath

Call readConfigText(pConfigText)

Sub readConfigText(DictPath)
    If Not oFso.FileExists(DictPath) Then Exit Sub
    
    Dim oText
    Set oText = oFso.OpenTextFile(DictPath, FOR_READING)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        If InStr(sReadLine, "TextEditorPath") > 0 Then getTextEditor(sReadLine)
        If InStr(sReadLine, "BeyondComparePath") > 0 Then getBeyondCompare(sReadLine)
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub getTextEditor(sReadLine)
    mTextEditorPath = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
End Sub

Sub getBeyondCompare(sReadLine)
    mBeyondComparePath = Trim(Mid(sReadLine, InStr(sReadLine, "=") + 1))
End Sub
