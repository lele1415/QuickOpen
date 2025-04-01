Option Explicit

Dim pProjectText : pProjectText = oWs.CurrentDirectory & "\res\project.ini"

Dim vaWorksInfo : Set vaWorksInfo = New VariableArray

Sub readWorksInfoText()
    If Not isFileExists(pProjectText) Then Exit Sub
    
    Dim oText, sReadLine
    Set oText = oFso.OpenTextFile(pProjectText, FOR_READING, False, True)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        If InStr(sReadLine, "########") > 0 Then
            Call handleForWorksInfo(oText, sReadLine)
        End If
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub applyLastWorkInfo()
    Dim oInfos
    If vaWorksInfo.Bound > -1 Then
        Set oInfos = vaWorksInfo.V(vaWorksInfo.Bound)
        if Not checkProjectExist(oInfos.Sdk, oInfos.Product, oInfos.Project) Then
            Call mIp.clearSdkInfos()
            Call mIp.clearWorkInfos()
            Exit Sub
        End If
        Call mIp.setProjectInputs(oInfos)
    End If
End Sub

Sub handleForWorksInfo(oText, sReadLine)
    Dim i, oInfos
    i = 0
    Set oInfos = New ProjectInfos

    oInfos.Work = Trim(oText.ReadLine)
    oInfos.Sdk = Trim(oText.ReadLine)
    oInfos.Product = Trim(oText.ReadLine)
    oInfos.Project = Trim(oText.ReadLine)
    If InStr(oInfos.Sdk, "_t0") > 0 Then
        oInfos.SysSdk = Trim(oText.ReadLine)
        oInfos.SysTarget = Trim(oText.ReadLine)
        oInfos.SysProject = Trim(oText.ReadLine)
    End If
    oInfos.Firmware = Trim(oText.ReadLine)
    oInfos.Requirements = Trim(oText.ReadLine)
    oInfos.Zentao = Trim(oText.ReadLine)

    vaWorksInfo.Append(oInfos)
End Sub

Sub showWorkInfo(taskNum)
    If isNumeric(taskNum) And Len(taskNum) < 5 Then
		Dim i, obj, infos : For i = vaWorksInfo.Bound To 0 Step -1
		    Set obj = vaWorksInfo.V(i)
		    If taskNum = obj.TaskNum Then
		    	infos = obj.Work & VbLf &_
                        obj.Sdk & VbLf &_
                        obj.Product & VbLf &_
                        obj.project
                If InStr(obj.Sdk, "_t0") > 0 Then
                    infos = infos & VbLf &_
                            obj.SysSdk & VbLf &_
                            obj.SysProject
                End If
                infos = infos & VbLf &_
                        obj.Firmware & VbLf &_
                        obj.Requirements & VbLf &_
                        obj.Zentao
                Call setOpenPath(infos)
		    	Exit For
		    End If
		Next
    End If
End Sub
