Option Explicit

Class BuildInfos
    Private mVos, mSdk, mProduct, mProject, mVersion, mPath

    Public Function isB0() : isB0 = isInStr(mSdk, "/b0") : End Function
    Public Function isV0() : isV0 = isInStr(mSdk, "/v_sys") : End Function
    Public Function isU0() : isU0 = isInStr(mSdk, "/u_sys") : End Function
    Public Function isT0() : isT0 = isInStr(mSdk, "/sys") : End Function
    Public Function isS0() : isS0 = isInStr(mSdk, "/vnd") : End Function
    Public Function isR0() : isR0 = isInStr(mSdk, "_r/") : End Function
    Public Function atLeast(version) : atLeast = isGe(mVersion, version) : End Function
    Public Function atMost(version) : atMost = isLe(mVersion, version) : End Function
    Public Function isSdkT0() : isSdkT0 = isInStr(mSdk, "_t0") : End Function
    Public Function isVnd() : isVnd = isInStr(mVos, "vnd") : End Function
    Public Function isSys() : isSys = isInStr(mVos, "sys") : End Function
    Public Function is8168() : is8168 = isInStr(mSdk, "8168") : End Function
    Public Function is8791() : is8791 = isInStr(mProduct, "8791") : End Function
    Public Function is8781()
        If isVnd() Then
            is8781 = isEq(mProduct, "tb8781p1_64")
        ElseIf isSys() Then
            is8781 = isCt(mProduct, "mssi_t_64_cn_armv82")
        Else
            is8781 = False
        End If
    End Function

    Private Sub getVersion()
        If isR0() Then
            mVersion = 11
        ElseIf isS0() Then
            mVersion = 12
        ElseIf isT0() Then
            mVersion = 13
        ElseIf isU0() Then
            mVersion = 14
        ElseIf isV0() Then
            mVersion = 15
        ElseIf isB0() Then
            mVersion = 16
        End If
    End Sub

    Public Function hasAlps() : hasAlps = isFolderExists(ProjectPath & "/alps") : End Function
    Public Function hasConfig() : hasConfig = isFolderExists(ProjectPath & "/config") : End Function

    Public Property Get Version : Version = mVersion : End Property
    Public Property Get ProductPath : ProductPath = "weibu/" & mProduct : End Property
    Public Property Get ProjectPath : ProjectPath = "weibu/" & mProduct & "/" & mProject : End Property
    Public Property Get OriginProjectConfigMk : OriginProjectConfigMk = "device/mediateksample/" & mProduct & "/ProjectConfig.mk" : End Property

    Public Property Get ProductFile
        If isVnd() Then
            If is8781() Then
                ProductFile = "device/mediateksample/" & mProduct & "/vext_" & mProduct & ".mk"
            Else
                ProductFile = "device/mediateksample/" & mProduct & "/vnd_" & mProduct & ".mk"
            End If
        Else
            ProductFile = "device/mediatek/system/" & mProduct & "/sys_" & mProduct & ".mk"
        End If
    End Property

    Public Property Get ProjectConfigMk
        If hasConfig() Then
            ProjectConfigMk = ProjectPath  & "/config/ProjectConfig.mk"
        Else
            ProjectConfigMk = getOverlayPath(OriginProjectConfigMk)
        End If
    End Property

    Public Property Get Out
        'If isSys() And (atLeast(15) Or is8781()) Then
        If isSys() And atLeast(13) Then
            Out = "out_sys"
        Else
            Out = "out"
        End If
    End Property
    Public Property Get OutPath : OutPath = Out & "/target/product/" & mProduct : End Property

    Public Property Get OutSystemExtPath
        If atLeast(14) Then
            OutSystemExtPath = OutPath & "/system_ext/priv-app"
        Else
            OutSystemExtPath = OutPath & "/system/system_ext/priv-app"
        End If
    End Property

    Public Property Get  OutSystemBuildProp : OutSystemBuildProp = OutPath & "/system/build.prop" : End Property
    Public Property Get  OutVendorBuildProp : OutVendorBuildProp = OutPath & "/vendor/build.prop" : End Property
    Public Property Get  OutProductBuildProp
        If atLeast(12) Then
            OutProductBuildProp = OutPath & "/product/etc/build.prop"
        Else
            OutProductBuildProp = OutPath & "/product/build.prop"
        End If
    End Property

    Public Property Get DownloadOutPath
        If isSdkT0() Then
            If is8781() Then
                DownloadOutPath = "..\merged\download_agent"
            Else
                DownloadOutPath = "..\merged"
            End If
        Else
            DownloadOutPath = OutPath
        End If
    End Property

    Public Property Get BootLogo
        Dim logo, csciPath
        logo = readTextAndGetValue("BOOT_LOGO", ProjectConfigMk)
        csciPath = ProjectPath & "/config/csci.ini"
        If isFileExists(csciPath) Then
            Dim oText, sReadLine
            Set oText = oFso.OpenTextFile(checkDriveSdkPath(csciPath), FOR_READING)
            Do Until oText.AtEndOfStream
                sReadLine = oText.ReadLine
                If InStr(sReadLine, "ro.vendor.fake.orientation") > 0 And InStr(sReadLine, " 1 ") > 0 Then
                    logo = logo & "nl"
                    Exit Do
                End If
            Loop
            oText.Close
            Set oText = Nothing
        End If
        BootLogo = logo
    End Property

    Public Property Get LogoPath
        If is8781() Then
            LogoPath = "vendor/mediatek/proprietary/external/BootLogo/logo/" & BootLogo
        Else
            LogoPath = "vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/" & BootLogo
        End If
    End Property

    Public Function getPowerProfilePath()
        If atLeast(14) Then
            getPowerProfilePath = "device/mediatek/system/common/overlay/power/frameworks/base/core/res/res/xml/power_profile.xml"
        Else
            getPowerProfilePath = "vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/power/res/xml/power_profile.xml"
        End If
    End Function

    Public Function getBluetoothfilePath()
        If atLeast(13) Then
            getBluetoothfilePath = "vendor/mediatek/proprietary/packages/modules/Bluetooth/system/btif/src/btif_dm.cc"
        Else
            getBluetoothfilePath = "system/bt/btif/src/btif_dm.cc"
        End If
    End Function

    Public Function getPlatform()
        getPlatform = LCase(readTextAndGetValue("MTK_PLATFORM", OriginProjectConfigMk))
    End Function

    Public Function getOverlayPath(path)
        If atMost(11) And Not is8168() Then
            getOverlayPath = ProjectPath & "/" & path
        Else
            getOverlayPath = ProjectPath & "/alps/" & path
        End if
    End Function

    Public Function getOutProp(prop)
        Dim value
        value = readTextAndGetValue(prop, OutSystemBuildProp)
        If value = "" Then value = readTextAndGetValue(prop, OutVendorBuildProp)
        If value = "" Then value = readTextAndGetValue(prop, OutProductBuildProp)
        getOutProp = value
    End Function

    Public Function getPathWithDriveSdk(path)
        path = relpaceBackSlashInPath(path)
        If InStr(path, "..\") = 1 Then
            getPathWithDriveSdk = relpaceBackSlashInPath(mDrive & getParentPath(mSdk) & "\" & Right(path, Len(path) - 3))
        Else
            getPathWithDriveSdk = relpaceBackSlashInPath(mDrive & mSdk & "\" & path)
        End If
    End Function

    Public Default Function Constructor(vos, sdk, product, project)
        mVos = vos
        mSdk = sdk
        mProduct = product
        mProject = project
        Call getVersion()
        Set Constructor = Me
    End Function
End Class

Class BaseBuild
    Private mSdk, mProduct, mProject, mInfos

    Public Property Get Sdk : Sdk = mSdk : End Property
    Public Property Get Product : Product = mProduct : End Property
    Public Property Get Project : Project = mProject : End Property
    Public Property Get Infos : Set Infos = mInfos : End Property

    Public Default Function Constructor(vos, sdk, product, project)
        mSdk = sdk
        mProduct = product
        mProject = project
        Set mInfos = (New BuildInfos)(vos, sdk, product, project)
        Set Constructor = Me
    End Function

    Public Function toString(sp)
        toString = mSdk & sp & mProduct & sp & mProject
    End Function

    
End Class

Class TaskInfos
    Private mTaskNum, mTaskName, mCustomName
    Public Property Get TaskNum : TaskNum = mTaskNum : End Property
    Public Property Get TaskName : TaskName = mTaskName : End Property
    Public Property Get CustomName : CustomName = mCustomName : End Property

    Public Default Function Constructor(taskNum, taskName, customName)
        mTaskNum = taskNum
        mTaskName = taskName
        mCustomName = customName
        Set Constructor = Me
    End Function

    Public Function toString(regex)
        toString = mTaskNum & regex & mTaskName & regex & mCustomName
    End Function
End Class

Class TaskBuild
    Private mInfos, mVnd, mSys
    Public Property Get Infos : Set Infos = mInfos : End Property
    Public Property Get Vnd : Set Vnd = mVnd : End Property
    Public Property Get Sys : Set Sys = mSys : End Property

    Public Default Function Constructor(infos, vnd, sys)
        Set mInfos = infos
        Set mVnd = vnd
        Set mSys = sys
        Set Constructor = Me
    End Function
End Class


