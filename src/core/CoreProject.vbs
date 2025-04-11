Option Explicit

Sub setDrive(drive)
    'If (drive = "z") Then
    '    mDrive = "Z:\work05\"
    'Else
    If (drive = "z6") Then
        mDrive = "Z:\work06\"
    ElseIf (drive = "x1") Then
        mDrive = "X:\work1\"
    ElseIf (drive = "x2") Then
        mDrive = "X:\work2\"
    End If
    
    Call updateTitle()
End Sub

Sub setSdk(sdk)
    If sdk = "8766s" Then
        mIp.Sdk = "mt8766_s\alps"
    ElseIf sdk = "8168s" Then
        mIp.Sdk = "mt8168_s\alps"
    ElseIf sdk = "8766r" Then
        If mDrive = "X:\work2\" Then
            mIp.Sdk = "mt8766_r\alps"
        'ElseIf mDrive = "Z:\work05\" Then
        '    mIp.Sdk = "mt8766_r\alps2"
        End If
    ElseIf sdk = "8766r" Then
        mIp.Sdk = "mt8168_r\alps"
    End If
End Sub

Sub applyProjectPath()
    Dim aInfos : aInfos = Split(getOpenPath(), "/")
    If UBound(aInfos) < 2 Then MsgBox("Not valid project path!") : Exit Sub

    Dim cleanPath : cleanPath = "weibu/" & aInfos(1) & "/" & aInfos(2)
    If isFolderExists(cleanPath) Then
        mIp.Product = aInfos(1)
        mIp.Project = aInfos(2)
    Else
        MsgBox("Path does not exist!" & Vblf & checkDriveSdkPath(cleanPath))
    End if
End Sub

Function checkProjectExist(sdk, product, project)
    If Not checkDrive(sdk, product, project) Then
        Call setDrive("x1")
        If Not checkDrive(sdk, product, project) Then
            Call setDrive("x2")
            If Not checkDrive(sdk, product, project) Then
                Call setDrive("z6")
                If Not checkDrive(sdk, product, project) Then 
                    checkProjectExist = False
                    Exit Function
                End If
            End If
        End If
    End If

    checkProjectExist = True
End Function

Function checkDrive(sdk, product, project)
    Dim path
	path = mDrive & sdk
	If isFolderExists(path) Then
	    path = path & "\weibu\" & product & "\" & project
		If isFolderExists(path) Then
		    checkDrive = True
			Exit Function
		ENd If
	End If

    checkDrive = False
End Function

Function isT0Sdk()
    If InStr(mIp.Infos.Sdk, "_t0") > 0 Then
        isT0Sdk = True
    Else
        isT0Sdk = False
    End If
End Function

Function isSplitSdkSys()
    If isV0SysSdk() Then
        If Not is8781Vnd() And InStr(mIp.Infos.Product, "mssi") > 0 Then
            isSplitSdkSys = True
        ElseIf is8781Vnd() And InStr(mIp.Infos.Sdk, "sys") > 0 Then
            isSplitSdkSys = True
        Else
            isSplitSdkSys = False
        End If
    Else
        If InStr(mIp.Infos.Sdk, "sys") > 0 Then
            isSplitSdkSys = True
        Else
            isSplitSdkSys = False
        End If
    End If
End Function

Function isSplitSdkVnd()
    If InStr(mIp.Infos.Sdk, "\vnd") > 0 Then
        isSplitSdkVnd = True
    ElseIf isV0SysSdk() And InStr(mIp.Infos.Product, "tb87") > 0 Then
        isSplitSdkVnd = True
    Else
        isSplitSdkVnd = False
    End If
End Function

Function isT08168Sdk()
    If InStr(mIp.Infos.Sdk, "t0_816x") > 0 Then
        isT08168Sdk = True
    Else
        isT08168Sdk = False
    End If
End Function

Function isT08168SdkVnd()
    If InStr(mIp.Infos.Sdk, "t0_816x\vnd") > 0 Then
        isT08168SdkVnd = True
    Else
        isT08168SdkVnd = False
    End If
End Function

Function is8781Vnd()
    is8781Vnd = InStr(mIp.Infos.VndTarget, "8781") > 0
End Function

Function is8791Vnd()
    is8791Vnd = InStr(mIp.Infos.VndTarget, "8791") > 0
End Function

Function isT0SysSdk()
    If InStr(mIp.Infos.SysSdk, "\sys") > 0 Then
        isT0SysSdk = True
    Else
        isT0SysSdk = False
    End If
End Function

Function isU0SysSdk()
    If InStr(mIp.Infos.SysSdk, "\u_sys") > 0 Then
        isU0SysSdk = True
    Else
        isU0SysSdk = False
    End If
End Function

Function isV0SysSdk()
    If InStr(mIp.Infos.SysSdk, "\v_sys") > 0 Then
        isV0SysSdk = True
    Else
        isV0SysSdk = False
    End If
End Function

Function checkWifiProduct(project)
    If isSplitSdkSys() And Not isFolderExists(getProjectPath(mIp.Infos.Product, project)) And _
            isFolderExists(getProjectPath(mIp.Infos.Product & "_wifi", project)) Then
        mIp.Infos.SysTarget = mIp.Infos.Product & "_wifi"
        mIp.Product = mIp.Infos.SysTarget
        checkWifiProduct = True
    Else
        checkWifiProduct = False
    End If
End Function

'Function getT0SysProjectFromVnd(vndProject)
'    If vndProject = "M101TB_DG_PT2_531" Then
'        getT0SysProjectFromVnd = "M101TB_DG_PT1_532-MMI-PT2_531"
'        Exit Function
'    End If
'
'    Dim mmiProject
'    If InStr(vndProject, "-") > 0 Then
'        Dim i, arrStr, str
'        arrStr = Split(vndProject, "-")
'        For i = 0 To UBound(arrStr)
'            If i = 0 Then
'                str = arrStr(i) & "-MMI"
'            Else
'                str = str & "-" & arrStr(i)
'            End If
'        Next
'        mmiProject = str
'    Else
'        mmiProject = vndProject & "-MMI"
'    End If
'
'    If checkWifiProduct(mmiProject) Then
'        getT0SysProjectFromVnd = mmiProject
'    ElseIf checkWifiProduct(vndProject) Then
'        getT0SysProjectFromVnd = vndProject
'    ElseIf isFolderExists(getProjectPath(mIp.Infos.Product, mmiProject)) Then
'        getT0SysProjectFromVnd = mmiProject
'    ElseIf isFolderExists(getProjectPath(mIp.Infos.Product, vndProject)) Then
'        getT0SysProjectFromVnd = vndProject
'    Else
'        getT0SysProjectFromVnd = ""
'    End If
'End Function

Sub setT0SdkSys()
    If isSplitSdkVnd() Then
        mIp.T0InnerSwitch = True
        mIp.Infos.VndSdk = mIp.Infos.Sdk
        mIp.Sdk = mIp.Infos.SysSdk
        mIp.T0InnerSwitch = True
        mIp.Product = mIp.Infos.SysTarget
        mIp.T0InnerSwitch = True
        mIp.Project = mIp.Infos.SysProject
        Call createWorkName()
    End If
End Sub

Sub setT0SdkVnd()
    If isSplitSdkSys() Then
        mIp.T0InnerSwitch = True
        mIp.Sdk = mIp.Infos.VndSdk
        mIp.T0InnerSwitch = True
        mIp.Product = mIp.Infos.VndTarget
        mIp.T0InnerSwitch = True
        mIp.Project = mIp.Infos.DriverProject
        Call createWorkName()
    End If
End Sub

Sub findProjectWithTaskNum(taskNum)
    Dim vaProduct
    Set vaProduct = searchFolder("weibu", "_", _
			SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)

    Dim i, product, projectArr
    For i = 0 To vaProduct.Bound
        product = vaProduct.V(i)
        Set projectArr = searchFolder(getProductPath(product), "_" & taskNum, _
                SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
        If projectArr.Bound < 0 Then
            Set projectArr = searchFolder(getProductPath(product), "-" & taskNum, _
                SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
        End If
        If projectArr.Bound >= 0 Then
            Call setOpenPath("weibu/" & product & "/" & projectArr.V(0))
            'mIp.Product = product
            'If projectArr.Bound = 0 Then
            '    mIp.Project = projectArr.V(0)
            'Else
            '    mFindProjectButtonList.VaArray = projectArr
            '    Call mFindProjectButtonList.addList()
            '    Call mFindProjectButtonList.toggleButtonList()
            'End If
            Exit For
        End If
    Next
End Sub

Function getWorkInfoWithTaskNum(taskNum, info)
    If isNumeric(taskNum) And Len(taskNum) < 5 Then
		Dim i, obj : For i = vaWorksInfo.Bound To 0 Step -1
		    Set obj = vaWorksInfo.V(i)
		    If taskNum = obj.TaskNum Then
                If info = "obj" Then
		    	    Set getWorkInfoWithTaskNum = obj
                ElseIf info = "work" Then
                    getWorkInfoWithTaskNum = obj.Work
                End if
		    	Exit Function
		    End If
		Next
    End If
    If info = "obj" Then
        Set getWorkInfoWithTaskNum = New ProjectInfos
    Else
        getWorkInfoWithTaskNum = ""
    End If
End Function
