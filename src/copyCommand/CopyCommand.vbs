Option Explicit

Const ID_COMMAND_ENG = "command_eng"
Const ID_COMMAND_USERDEBUG = "command_userdebug"
Const ID_COMMAND_USER = "command_user"

Const ID_COMMAND_RM_OUT = "command_rm_out"
Const ID_COMMAND_RM_BUILDPROP = "command_rm_buildprop"
Const ID_COMMAND_BUILD_OTA = "command_build_ota"

Dim commandFinal

Sub CommandOfMake()
    Dim rmOut, rmBuildprop, ota
    rmOut = element_isChecked(ID_COMMAND_RM_OUT)
    rmBuildprop = element_isChecked(ID_COMMAND_RM_BUILDPROP)
    ota = element_isChecked(ID_COMMAND_BUILD_OTA)
    Call getMakeCommand(rmOut, rmBuildprop, ota)
End Sub

Sub getMakeCommand(rmOut, rmBuildprop, ota)
    Dim commandOta
    commandFinal = "make -j36 2>&1 | tee build.log"
    commandOta = "make -j36 otapackage 2>&1 | tee build_ota.log"
    
    If rmOut Then
        commandFinal = "rm -rf out/ && " & commandFinal
    ElseIf rmBuildprop Then
        commandFinal = "find " & mIp.Infos.OutPath & " -type f -name build*.prop | xargs rm -v && " & commandFinal
    End If

    If ota Then commandFinal = commandFinal & " && " & commandOta

    Call CopyString(commandFinal)
End Sub

Sub CommandOfLunch()
    If Not mIp.hasProjectInfos() Then Exit Sub
    Dim buildType

    Select Case True
        Case element_isChecked(ID_COMMAND_ENG)
            buildType = "eng"
        Case element_isChecked(ID_COMMAND_USERDEBUG)
            buildType = "userdebug"
        Case element_isChecked(ID_COMMAND_USER)
            buildType = "user"
    End Select

    Call getLunchCommand(buildType)
End Sub

Sub getLunchCommand(buildType)
    Dim comboName
    If InStr(mIp.Infos.Sdk, "8168") > 0 Then
        commandFinal = getLunchItemInSplitBuild(buildType)
        Call CopyQuoteString(commandFinal)
        Exit Sub
    Else
        comboName = "full_" & mIp.Infos.Product & "-" & buildType
        commandFinal = "source build/envsetup.sh ; lunch " & comboName & " " & mIp.Infos.Project
    End If
    Call CopyString(commandFinal)
End Sub

Function getLunchItemInSplitBuild(buildType)
    Dim sysStr, vndStr, commandStr
    sysStr = "sys_" & mIp.Infos.SysTarget & "-" & buildType
    vndStr = "vnd_" & mIp.Infos.Product & "-" & buildType
    commandStr = sysStr & " " & vndStr & " " & mIp.Infos.Project
    commandStr = """lunch_item=""&Chr(34)&""" & commandStr & """&Chr(34)"

    Dim keyStr
    keyStr = "##Cusomer Settings"
    commandStr = """sed -i '/" & keyStr & "/i\""&" & commandStr & "&""' split_build.sh;git diff split_build.sh"""
    getLunchItemInSplitBuild = commandStr
End Function

Sub CommandOfOut()
    If Not mIp.hasProjectInfos() Then Exit Sub
    Call CopyString(mIp.Infos.DownloadOutPath)
End Sub

Sub CopyCleanCommand()
	commandFinal = "git checkout .;git clean -df"
	Call CopyString(commandFinal)
End Sub

Sub CopyCommitInfo(what)
    If Not mIp.hasProjectInfos() Then Exit Sub
    If what = "" Then
	    commandFinal = "[" & mIp.Infos.Project & "] : "
        Call CopyString(commandFinal)
        Exit Sub

    ElseIf what = "lg" Then
        commandFinal = "Logo [" & mIp.Infos.Project & "] : ���ƿ���logo"
    ElseIf what = "ani" Then
        commandFinal = "BootAnimation [" & mIp.Infos.Project & "] : ���ƿ�������"
    ElseIf what = "wp" Then
        commandFinal = "Wallpaper [" & mIp.Infos.Project & "] : ����Ĭ�ϱ�ֽ"
    ElseIf what = "loc" Then
        commandFinal = "Locale [" & mIp.Infos.Project & "] : Ĭ������"
    ElseIf what = "tz" Then
        commandFinal = "Timezone [" & mIp.Infos.Project & "] : Ĭ��ʱ��"
    ElseIf what = "di" Then
        commandFinal = "DisplayId [" & mIp.Infos.Project & "] : �汾��"
    ElseIf what = "bn" Then
        commandFinal = "BuildNumber [" & mIp.Infos.Project & "] : build number "
    ElseIf what = "bm" Then
        commandFinal = "MMI [" & mIp.Infos.Project & "] : Ʒ�ƣ��ͺ�"
    ElseIf what = "bwm" Then
        commandFinal = "MMI [" & mIp.Infos.Project & "] : ������WiFi�ȵ㡢WiFiֱ�����̷�"
    ElseIf what = "mmi" Then
        commandFinal = "MMI [" & mIp.Infos.Project & "] : "
    ElseIf what = "st" Then
        commandFinal = "Settings [" & mIp.Infos.Project & "] : "    
    ElseIf what = "su" Then
        commandFinal = "SystemUI [" & mIp.Infos.Project & "] : "
    ElseIf what = "lac" Then
        commandFinal = "Launcher [" & mIp.Infos.Project & "] : "
    ElseIf what = "bt" Then
        commandFinal = "Bluetooth [" & mIp.Infos.Project & "] : Ĭ����������"
    ElseIf what = "wfap" Then
        commandFinal = "WiFi [" & mIp.Infos.Project & "] : Ĭ��WiFi�ȵ�����"
    ElseIf what = "wfdrt" Then
        commandFinal = "WiFi [" & mIp.Infos.Project & "] : Ĭ��WiFiֱ������"
    ElseIf what = "mtp" Then
        commandFinal = "MTP [" & mIp.Infos.Project & "] : Ĭ���̷�����"
    ElseIf what = "brt" Then
        commandFinal = "Brightness [" & mIp.Infos.Project & "] : Ĭ������"
    ElseIf what = "ad" Then
        commandFinal = "Audio [" & mIp.Infos.Project & "] : Ĭ������"
    ElseIf what = "slp" Then
        commandFinal = "Settings [" & mIp.Infos.Project & "] : Ĭ������ʱ��"
    ElseIf what = "hp" Then
        commandFinal = "Browser [" & mIp.Infos.Project & "] : Ĭ����ַ"
    ElseIf what = "bat" Then
        commandFinal = "Battery [" & mIp.Infos.Project & "] : ��ؼ������"
    ElseIf what = "app" Then
        commandFinal = "App [" & mIp.Infos.Project & "] : "
    End If
	Call CopyQuoteString("""git add weibu;git commit -m ""&Chr(34)&""" & commandFinal & """&Chr(34)")
End Sub

Sub CopyDriverCommitInfo()
    If Not mIp.hasProjectInfos() Then Exit Sub
	commandFinal = "[" & mIp.Infos.DriverProject & "] : "
	Call CopyString(commandFinal)
End Sub

Sub CopyBuildOtaUpdate()
    If InStr(mIp.Infos.Sdk, "_s") > 0 Then
        commandFinal = "./out/host/linux-x86/bin/ota_from_target_files -i old.zip new.zip update.zip"
    Else
        commandFinal = "./build/tools/releasetools/ota_from_target_files -i old.zip new.zip update.zip"
    End If
    Call CopyString(commandFinal)
End Sub

Sub MkdirWeibuFolderPath()
    If Not mIp.hasProjectInfos() Then Exit Sub
    Dim path : path = getOpenPath()
    Dim folderPath
    Dim mkdirCmd, cpCmd
    commandFinal = ""

    If isFileExists(path) Or isFolderExists(path) Then
        Dim index : index = InStrRev(path, "/")
        folderPath = Left(path, index)

        If isFolderExists(folderPath) Then

            If Not isFolderExists(mIp.Infos.getOverlayPath(folderPath)) Then
                mkdirCmd = "mkdir -p " & mIp.Infos.getOverlayPath(folderPath) & ";"
            End If

            commandFinal = mkdirCmd
        End If
    End If

    If isFileExists(path) Then
        If Not isFileExists(mIp.Infos.getOverlayPath(path)) Then
            cpCmd = "cp " & path & " " & mIp.Infos.getOverlayPath(folderPath)
        Else
            MsgBox("File exist!")
        End If
            commandFinal = mkdirCmd & cpCmd
    End If

    commandFinal = Replace(commandFinal, "\", "/")
    Call CopyString(commandFinal)
End Sub

Sub copyExportToolsPathCmd()
    commandFinal = "export PATH=$HOME/Tools:$PATH"
    Call CopyString(commandFinal)
End Sub

Sub CopyPathInWindows()
    If oFso.FileExists(mIp.Infos.getOverlaySdkPath(getOpenPath())) Then
        commandFinal = mIp.Infos.getPathWithDriveSdk(Replace(mIp.Infos.getOverlayPath(getOpenPath()), "/", "\"))
    Else
        commandFinal = mIp.Infos.getPathWithDriveSdk(Replace(getOpenPath(), "/", "\"))
    End If
    Call CopyString(commandFinal)
End Sub

Sub CopyPathInLinux()
    If oFso.FileExists(mIp.Infos.getOverlaySdkPath(getOpenPath())) Then
        commandFinal = mIp.Infos.getOverlaySdkPath(getOpenPath())
    Else
        commandFinal = mIp.Infos.Sdk & "\" & getOpenPath()
    End If
    Call CopyString(commandFinal)
End Sub
