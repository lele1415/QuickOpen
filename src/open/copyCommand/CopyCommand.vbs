Option Explicit

Const ID_COMMAND_ENG = "command_eng"
Const ID_COMMAND_USERDEBUG = "command_userdebug"
Const ID_COMMAND_USER = "command_user"

Const ID_COMMAND_RM_OUT = "command_rm_out"
Const ID_COMMAND_RM_BUILDPROP = "command_rm_buildprop"
Const ID_COMMAND_BUILD_OTA = "command_build_ota"

Dim commandFinal

Sub copyStrAndPasteInXshell(cmdStr)
    If cmdStr = "" Then MsgBox("Empty!") : Exit Sub
    Call CopyString(cmdStr)
    Call pasteCmdInXshell()
End Sub

Sub copyStrAndPasteInPowerShell(cmdStr)
    If cmdStr = "" Then MsgBox("Empty!") : Exit Sub
    Call CopyString(cmdStr)
    'Call pasteCmdInPowerShell()
End Sub

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

    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Function getSplitBuildCommand(opts)
    If Not checkBuildType() Then getSplitBuildCommand = "" : Exit Function
    Dim buildsh, params, commandStr
    If is8781Vnd() Then
        buildsh = "./split_build_v2.sh"
    Else
        buildsh = "./split_build.sh"
    End If
    If opts = "a" Then
        If is8781Vnd() Then
            params = " vext krn hal sys m p"
        ElseIf isV0SysSdk() Then
            params = " krn vnd sys m p"
        Else
            params = " vnd krn sys m p"
        End If
    Else
        If InStr(opts, "v") Then
            If is8781Vnd() Then
                params = params & " vext"
            Else
                params = params & " vnd"
            End If
        End If
        If InStr(opts, "k") Then params = params & " krn"
        If is8781Vnd() And InStr(opts, "h") Then params = params & " hal"
        If InStr(opts, "s") Then params = params & " sys"
        If InStr(opts, "m") Then params = params & " m"
        If InStr(opts, "p") Then params = params & " p"
    End if

    commandStr = buildsh & params

    if Not isV0SysSdk() And InStr(params, " p") And InStr(params, " vnd") = 0 And InStr(params, " krn") = 0 And InStr(params, " hal") = 0 Then
        Call setT0SdkVnd()
        commandStr = getCustomModemSedStr() & commandStr
    End If
    getSplitBuildCommand = commandStr & ";"
End Function

Sub copySplitBuildCommand(opts)
    Call copyStrAndPasteInXshell(getSplitBuildCommand(opts))
End Sub

Function getCustomModemSedStr()
    Dim customModem, deviceModem
    customModem = getMMIProjectConfigValue("CUSTOM_MODEM")
    deviceModem = getDeviceProjectConfigValue("CUSTOM_MODEM")
    if customModem <> deviceModem Then
        getCustomModemSedStr = getSedCmd("", "CUSTOM_MODEM", "=.*$", "= " & customModem, "vnd/" & pDeviceProjectConfigMk)
    Else
        getCustomModemSedStr = ""
    End If
End Function

Sub getSplitTestOTABuildCommand(opts)
    Dim cmdStr
    cmdStr = getSplitBuildCommand(opts)
    
    If isV0SysSdk() And Not is8781Vnd() Then
        cmdStr = cmdStr & "mkdir -p ../OTA/" & mIp.Infos.TaskNum & ";"
        cmdStr = cmdStr & "mv merged/target_files.zip ../OTA/" & mIp.Infos.TaskNum & "/target_files_s.zip;"
        cmdStr = cmdStr & getOTATestSedStr(False) & ";"
        cmdStr = cmdStr & getSplitBuildCommand("sm")
    Else
        cmdStr = cmdStr & "mkdir -p OTA/" & mIp.Infos.TaskNum & ";"
        cmdStr = cmdStr & "mv merged/target_files.zip OTA/" & mIp.Infos.TaskNum & "/target_files_s.zip;"
        If isSplitSdkVnd() Then Call setT0SdkSys()
        cmdStr = cmdStr & "cd " & getFileNameFromPath(mIp.Infos.SysSdk) & ";"
        cmdStr = cmdStr & getOTATestSedStr(False) & ";"
        cmdStr = cmdStr & "cd ..;"
        cmdStr = cmdStr & getSplitBuildCommand("sm")
    End If
    Call copyStrAndPasteInXshell(cmdStr)
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

Sub getLunchCommand(buildType, taskNum)
    Dim comboName
    If isT0Sdk() Or (InStr(mIp.Infos.Sdk, "8168") > 0 And Not InStr(mIp.Infos.Sdk, "_r") > 0) Then
        commandFinal = getLunchCommandInSplitBuild(buildType, taskNum)
    Else
        comboName = "full_" & mIp.Infos.Product & "-" & buildType
        commandFinal = "source build/envsetup.sh ; lunch " & comboName & " " & mIp.Infos.Project
    End If
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Function getLunchCommandInSplitBuild(buildType, taskNum)
    Dim obj, lunchStr, commandStr
    If isNumeric(taskNum) And Len(taskNum) < 5 Then
		Set obj = getWorkInfoWithTaskNum(taskNum, "obj")

        If isT0Sdk() Then
            lunchStr = getLunchStrFromWSavedWork(buildType, obj)
            If lunchStr = "" Then getLunchCommandInSplitBuild = "" : Exit Function
            commandStr = "sed -i 's/^.*$/" & lunchStr & "/' lunch_item"
            If InStr(obj.Product, "tb8781") Then commandStr = commandStr & "_v2"
        Else
            lunchStr = getLunchStrFromWSavedWork(buildType, obj)
            If lunchStr = "" Then getLunchCommandInSplitBuild = "" : Exit Function
            Dim keyStr
            keyStr = "##Cusomer Settings"
            commandStr = "sed -i '/" & keyStr & "/i\" & lunchStr & "' split_build.sh;git diff split_build.sh"
        End If
    End If
    getLunchCommandInSplitBuild = commandStr
End Function

Function getLunchStrFromWSavedWork(buildType, obj)
    Dim lunchStr
    If obj.SysSdk <> "" Then
        If InStr(obj.SysSdk, "\v_sys") Then
            If InStr(obj.Sdk, "\v_sys") Then
                lunchStr = getVVLunchStr(obj, buildType)
            ElseIf InStr(obj.Sdk, "\vnd") > 0 And InStr(obj.Product, "tb8781") > 0 Then
                lunchStr = get8781LunchStr(obj, buildType, "-next-", " V")
            Else
                lunchStr = ""
            End If
        ElseIf InStr(obj.SysSdk, "\u_sys") Then
            If InStr(obj.Product, "tb8781") Then
                lunchStr = get8781LunchStr(obj, buildType, "-", " U")
            Else
                lunchStr = getSULunchStr(obj, buildType, " U")
            End If
        ElseIf InStr(obj.SysSdk, "\sys") Then
            If InStr(obj.Product, "tb8781") Then
                lunchStr = get8781LunchStr(obj, buildType, "-", " T")
            Else
                lunchStr = getSULunchStr(obj, buildType, " T")
            End If
        ElseIf InStr(obj.Sdk, "mt8168_r") Then
            Dim sysStr, vndStr
            sysStr = "sys_" & mIp.Infos.SysTarget & "-" & buildType
            vndStr = "vnd_" & mIp.Infos.VndTarget & "-" & buildType
            lunchStr = sysStr & " " & vndStr & " " & mIp.Infos.Project
            lunchStr = "lunch_item=""&Chr(34)&""" & lunchStr & """&Chr(34)&"""
        Else
            lunchStr = ""
        End If
    Else
        lunchStr = ""
    End if

    getLunchStrFromWSavedWork = lunchStr
End Function

Function getVVLunchStr(obj, buildType)
    If obj.Product <> "" And obj.Project <> "" And obj.SysTarget <> "" And obj.SysProject <> "" Then
        getVVLunchStr = "vnd_" & obj.Product & "-next-" & buildType & " " & obj.Project &_
                " sys_" & obj.SysTarget & "-next-" & buildType & " " & obj.SysProject
    Else
        getVVLunchStr = ""
    End If
End Function

Function getSULunchStr(obj, buildType, androidVer)
    If obj.Product <> "" And obj.Project <> "" And obj.SysTarget <> "" And obj.SysProject <> "" Then
        getSULunchStr = " vnd_" & obj.Product & "-" & buildType & " " & obj.Project &_
                " sys_" & obj.SysTarget & "-" & buildType & " " & obj.SysProject &_
                androidVer
    Else
        getSULunchStr = ""
    End If
End Function

Function get8781LunchStr(obj, buildType, releseStr, androidVer)
    If obj.Product <> "" And obj.Project <> "" And obj.SysTarget <> "" And obj.SysProject <> "" Then
        get8781LunchStr = "hal_" & obj.HalTarget & "-" & buildType &_
                " krn_" & obj.KrnTarget & "-" & buildType &_
                " vext_" & obj.Product & "-" & buildType & " " & obj.Project &_
                " sys_" & obj.SysTarget & releseStr & buildType & " " & obj.SysProject &_
                androidVer
    Else
        get8781LunchStr = ""
    End If
End Function

Sub getT0SysLunchCommand(buildType)
    If isV0SysSdk() THen
        commandFinal = "source build/envsetup.sh && lunch sys_" & mIp.Infos.SysTarget & "-next-" & buildType & " " & mIp.Infos.SysProject
    Else
        commandFinal = "source build/envsetup.sh && lunch sys_" & mIp.Infos.SysTarget & "-" & buildType & " " & mIp.Infos.SysProject
    End If
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub getT0VndLunchCommand(buildType)
    If is8781Vnd() Then
        commandFinal = "source build/envsetup.sh && lunch vext_" & mIp.Infos.VndTarget & "-" & buildType & " " & mIp.Infos.DriverProject
    ELse
        commandFinal = "source build/envsetup.sh && lunch vnd_" & mIp.Infos.VndTarget & "-" & buildType & " " & mIp.Infos.DriverProject
    End If
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub findProjectPathWithTaskNum(taskNum)
    Dim commandFinal
    If isNumeric(taskNum) And Len(taskNum) < 5 Then
        commandFinal = "find weibu -maxdepth 2 -name ""&Chr(34)&""*" & taskNum & "*""&Chr(34)&"""
    End If
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub getProjectPathWithTaskNum(taskNum, which)
    Dim obj
    If isNumeric(taskNum) And Len(taskNum) < 5 Then
		Set obj = getWorkInfoWithTaskNum(taskNum, "obj")
        If which = "s" Then
            If obj.SysTarget <> "" And obj.SysProject <> "" Then
                Call setOpenPath("weibu/" & obj.SysTarget & "/" & obj.SysProject)
            Else
                Call setOpenPath("weibu/" & obj.Product & "/" & obj.Project)
            End If
        Else
            Call setOpenPath("weibu/" & obj.Product & "/" & obj.Project)
        End If
    End If
End Sub

Function getRomPath()
    getRomPath = getParentPath(mIp.Infos.DriveSdk) & "\ROM"
End Function

Sub CommandOfOut()
    If Not mIp.hasProjectInfos() Then Exit Sub
    Call CopyString(mIp.Infos.DownloadOutPath)
End Sub

Sub CopyCleanCommand(sysAndVnd)
    if sysAndVnd Then
        commandFinal = "cd sys/;git checkout .;git clean -df;cd ../vnd/;git checkout .;git clean -df;cd ../"
    Else
	    commandFinal = "git checkout .;git clean -df"
    End If
	Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub CopyDriverCommitInfo()
    If Not mIp.hasProjectInfos() Then Exit Sub
	commandFinal = "[" & mIp.Infos.DriverProject & "] : "
	Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub CopyBuildOtaUpdate()
    If InStr(mIp.Infos.Sdk, "_r") > 0 Then
        commandFinal = "./build/tools/releasetools/ota_from_target_files -i old.zip new.zip update.zip"
    Else
        commandFinal = "./out/host/linux-x86/bin/ota_from_target_files -i target_files_.zip target_files.zip update__to_.zip"
        If InStr(mIp.Infos.OutPath, "out_sys") > 0 Then commandFinal = Replace(commandFinal, "/out/", "/out_sys/")
    End If
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Function getParentProject()
    Dim vaParents, project, index, folderPath
    Set vaParents = New VariableArray
    project = mIp.Infos.Project
    index = InStrRev(project, "-")
    Do While index > 0
        project = Left(project, index - 1)
        folderPath = "weibu/" & mIp.Infos.Product & "/" & project
        If isFolderExists(folderPath) Then
            vaParents.Append(folderPath)
        End If
        index = InStrRev(project, "-")
    Loop
    Set getParentProject = vaParents
End Function

Sub MkdirWeibuFolderPath()
    If Not mIp.hasProjectInfos() Then Exit Sub
    Dim path : path = getOpenPath()
    If Not isFileExists(path) Then Exit Sub
    Dim folderPath
    Dim fileName
    Dim filePath
    Dim overlayFilePath
    Dim overlayFolderPath
    Dim mkdirCmd, cpCmd
    commandFinal = ""

    folderPath = Left(path, InStrRev(path, "/"))
    fileName = Replace(path, folderPath, "")
    If InStr(path, "weibu/") = 1 Then
        folderPath = Right(folderPath, Len(folderPath) - InStr(folderPath, "alps/") - Len("alps/") + 1)
    End If
    filePath = folderPath & fileName
    overlayFilePath = mIp.Infos.getOverlayPath(filePath)
    overlayFolderPath = mIp.Infos.getOverlayPath(folderPath)

    If isFileExists(overlayFilePath) Then MsgBox("File exist! " & overlayFilePath) : Exit Sub
    If Not isFolderExists(overlayFolderPath) Then
        mkdirCmd = "mkdir -p " & overlayFolderPath & ";"
    End If

    If InStr(path, "weibu/") <> 1 And InStr(mIp.Infos.Project, "-") > 0 Then
        Dim vaParents, i
        Set vaParents = getParentProject()
        If vaParents.Bound > -1 Then
            For i = 0 To vaParents.Bound
                If isFileExists(vaParents.V(i) & "/alps/" & path) Then
                    path = vaParents.V(i) & "/alps/" & path
                    Exit For
                End If
            Next
        End If
    End If

    cpCmd = "cp " & path & " " & mIp.Infos.getOverlayPath(folderPath)
    commandFinal = mkdirCmd & cpCmd

    commandFinal = relpaceSlashInPath(commandFinal)
    If InStr(path, "weibu/") = 1 Then Call setOpenPath(filePath)
    Call addProjectPath()
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub copyExportToolsPathCmd()
    commandFinal = "export PATH=$HOME/Tools:$PATH"
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub CopyPathInWindows()
    If isFileExists(mIp.Infos.getOverlayPath(getOpenPath())) Then
        commandFinal = mIp.Infos.getPathWithDriveSdk(mIp.Infos.getOverlayPath(getOpenPath()))
    Else
        commandFinal = mIp.Infos.getPathWithDriveSdk(getOpenPath())
    End If
    Call CopyString(commandFinal)
End Sub

Sub CopyPathInLinux()
    If isFileExists(mIp.Infos.getOverlayPath(getOpenPath())) Then
        commandFinal = mIp.Infos.Sdk & "\" & mIp.Infos.getOverlayPath(getOpenPath())
    Else
        commandFinal = mIp.Infos.Sdk & "\" & getOpenPath()
    End If
    Call CopyString(commandFinal)
End Sub

Sub copyCdSdkCommand()
    Dim path, arr
    arr = Split(mIp.Infos.DriveSdk, ":\")
    path = relpaceSlashInPath(arr(1))
    Call copyStrAndPasteInXshell("cd " & path)
End Sub

Function getSedCmd(cmdStr, searchStr, replaceStr, newStr, filePath)
    If isArray(filePath) Then
        Dim i, str
        str = cmdStr & "sed -i '/" & checkBackslash(searchStr) & "/s/" & checkBackslash(replaceStr) & "/" & checkBackslash(newStr) & "/'"
        For i = 0 To UBound(filePath)
            str = str & " " & filePath(i)
        Next
        getSedCmd = str & ";"
    Else
        getSedCmd = cmdStr & "sed -i '/" & checkBackslash(searchStr) & "/s/" & checkBackslash(replaceStr) & "/" & checkBackslash(newStr) & "/' " & filePath & ";"
    End If
End Function

Function getSedAddCmd(cmdStr, searchStr, addStr, filePath)
    getSedAddCmd = cmdStr & "sed -i '/" & checkBackslash(searchStr) & "/a\" & checkBackslash(addStr) & "' " & filePath & ";"
End Function

Function getGitDiffCmd(cmdStr, filePath)
    If isFileExists(mIp.Infos.getOverlayPath(filePath)) Then
        getGitDiffCmd = cmdStr & "git diff " & mIp.Infos.getOverlayPath(filePath) & ";"
    Else
        getGitDiffCmd = cmdStr & "git diff --no-index " & filePath & " " & mIp.Infos.getOverlayPath(filePath) & ";"
    End If
End Function

Function GetDiffCmdFromOverlayPath()
    Dim overlayPath, originPath
    overlayPath = getOpenPath()
    originPath = getOriginPathFromOverlayPath(getOpenPath())
    Call copyStrAndPasteInXshell("git diff --no-index " & originPath & " " & overlayPath)
End Function

Function getMultiMkdirStr(arr, what)
    Dim str, path, ovlFolder, ovlFile
    For Each path In arr
	    ovlFolder = mIp.Infos.getOverlayPath(getParentPath(path))
	    ovlFile = mIp.Infos.getOverlayPath(path)
	    If (Not (what = "lg" And InStr(path, "_kernel.bmp") > 0)) And (Not isFolderExists(ovlFolder)) Then
	        str =  str & "mkdir -p " & ovlFolder & ";"
		End If

	    If what = "lg" Then
	        str =  str & "cp ../File/logo.bmp " & ovlFile & ";"
		ElseIf what = "ani" Then
		    If InStr(path, "bootanimation.zip") > 0 Then
	            str =  str & "cp ../File/bootanimation.zip " & ovlFolder & ";"
			ElseIf InStr(path, "products.mk") > 0 Then
			    If Not isFileExists(ovlFile) Then str =  str & "cp " & path & " " & ovlFolder & ";"
				str =  getSedCmd(str, "bootanimation", "#", "", mIp.Infos.getOverlayPath(path))
				str =  getGitDiffCmd(str, path)
			End If
		ElseIf what = "wp" Then
		    If InStr(path, "default_wallpaper.png") > 0 Then
	            str =  str & "cp ../File/default_wallpaper.png " & ovlFolder & ";"
			ElseIf InStr(path, "default_wallpaper.jpg") > 0 Then
	            str =  str & "cp ../File/default_wallpaper.jpg " & ovlFolder & ";"
			End If
		Else
		    str =  str & "cp " & path & " " & ovlFolder & ";"
		End If
	Next
	getMultiMkdirStr = str
End Function

Function getLogoPath()
    If is8781Vnd() Then
        getLogoPath = "vendor/mediatek/proprietary/external/BootLogo/logo/[boot_logo]/"
    Else
        getLogoPath = "vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]/"
    End If
End Function

Function getPowerProfilePath()
    If isU0SysSdk() Or isV0SysSdk() Then
        getPowerProfilePath = "device/mediatek/system/common/overlay/power/frameworks/base/core/res/res/xml/power_profile.xml"
    Else
        getPowerProfilePath = "vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/power/res/xml/power_profile.xml"
    End If
End Function

Function getOutSystemExtPrivAppPath()
    If isU0SysSdk() Or isV0SysSdk() Then
        getOutSystemExtPrivAppPath = "/system_ext/priv-app"
    Else
        getOutSystemExtPrivAppPath = "/system/system_ext/priv-app"
    End If
End Function

Sub mkdirLogo()
    Dim lg_fd, lg_u, lg_k
    lg_fd = getLogoPath()
    lg_u = Replace(lg_fd & "[boot_logo]_uboot.bmp", "[boot_logo]", mIp.Infos.BootLogo)
    lg_k = Replace(lg_fd & "[boot_logo]_kernel.bmp", "[boot_logo]", mIp.Infos.BootLogo)

    Dim arr, finalStr
    arr = Array(lg_u, lg_k)
    finalStr = getMultiMkdirStr(arr, "lg")
    Call copyStrAndPasteInXshell(finalStr)
End Sub

Sub mkdirBootAnimation()
    Dim ani_media, ani_product
    ani_media = "vendor/weibu_sz/media/bootanimation.zip"
    ani_product = "vendor/weibu_sz/products/products.mk"

    Dim arr, finalStr
    arr = Array(ani_media, ani_product)
    finalStr = getMultiMkdirStr(arr, "ani")
    Call copyStrAndPasteInXshell(finalStr)
End Sub

Sub mkdirWallpaper(go)
    Dim wp_gms, wp_go1, wp_go2, wp1, wp2, wp3
    If isT0Sdk() Then
        wp_gms = "vendor/partner_gms/overlay/AndroidGmsBetaOverlay/res/drawable-nodpi/default_wallpaper.png"
    Else
        wp_gms = "vendor/partner_gms/overlay/AndroidSGmsBetaOverlay/res/drawable-nodpi/default_wallpaper.png"
    End If
    wp_go1 = "device/mediatek/common/overlay/ago/frameworks/base/core/res/res/drawable-nodpi/default_wallpaper.jpg"
    wp_go2 = "device/mediatek/system/common/overlay/ago/frameworks/base/core/res/res/drawable-nodpi/default_wallpaper.jpg"
    wp1 = "frameworks/base/core/res/res/drawable-nodpi/default_wallpaper.png"
    wp2 = "frameworks/base/core/res/res/drawable-sw600dp-nodpi/default_wallpaper.png"
    wp3 = "frameworks/base/core/res/res/drawable-sw720dp-nodpi/default_wallpaper.png"

    Dim arr, finalStr
    If Not go Then
        arr = Array(wp_gms, wp1, wp2, wp3)
    Else
        arr = Array(wp_go1, wp_go2, wp1, wp2, wp3)
    End If
    finalStr = getMultiMkdirStr(arr, "wp")
    Call copyStrAndPasteInXshell(finalStr)
End Sub

Sub mkdirTee()
    Call setT0SdkVnd()
    Dim teeOverlayPath, finalStr
    teeOverlayPath = mIp.Infos.getOverlayPath("vendor/mediatek/proprietary/trustzone/trustkernel/source/build/" & mIp.Infos.Product)
    If isFolderExists(teeOverlayPath) And isFileExists(teeOverlayPath & "cert.dat") And isFileExists(teeOverlayPath & "array.c") Then
        MsgBox("tee files exist!")
        Exit Sub
    End If
    If Not isFolderExists(teeOverlayPath) Then
        finalStr = "mkdir -p " & teeOverlayPath & ";"
    End If
    If Not isFileExists(teeOverlayPath & "cert.dat") Then
        finalStr = finalStr & "cp ../File/cert.dat " & teeOverlayPath & ";"
    End If
    If Not isFileExists(teeOverlayPath & "array.c") Then
        finalStr = finalStr & "cp ../File/array.c " & teeOverlayPath & ";"
    End If
    Call copyStrAndPasteInXshell(finalStr)
End Sub

Sub mkdirProductInfo(where)
    If Not isFileExists("../File/product.txt") Then MsgBox("product.txt does not exist!") : Exit Sub
    If where = "sys" Then Call setT0SdkSys()
    If where = "vnd" Then Call setT0SdkVnd()
    Dim info, infoArr, infoDict, cmdStr, finalStr
    infoArr = Array("brand", "manufacturer", "model", "name", "device")
    Set infoDict = CreateObject("Scripting.Dictionary")
    For Each info In infoArr
        Call infoDict.Add(info, readTextAndGetValue(info, "../File/product.txt"))
    Next
    For Each info In infoArr
        If infoDict.Item(info) <> "" Then
            cmdStr = cmdStr & getCmdStrForCpFileAndSetValue(Array(info, infoDict.Item(info)))
        End If
    Next
    Dim cmd, cmdArr, mkdirFlag, cpFlag, diffStr
    cmdArr = Split(cmdStr, ";")
    mkdirFlag = False
    cpFlag = False
    For Each cmd In cmdArr
        If InStr(cmd, "mkdir") = 1 Then
            If Not mkdirFlag Then
                finalStr = finalStr & cmd & ";"
                mkdirFlag = True
            End If
        ElseIf InStr(cmd, "cp") = 1 Then
            If Not cpFlag Then
                finalStr = finalStr & cmd & ";"
                cpFlag = True
            End If
        ElseIf InStr(cmd, "git diff") = 1 Then
            diffStr = cmd
        Else
            If cmd <> "" Then finalStr = finalStr & cmd & ";"
        End If
    Next
    finalStr = finalStr & diffStr
    Call copyStrAndPasteInXshell(finalStr)
End Sub

Sub CopyCommitInfo(what)
    If Not mIp.hasProjectInfos() Then Exit Sub
    If what = "" Then
	    commandFinal = "[" & mIp.Infos.Project & "] : "
        Call copyStrAndPasteInXshell(commandFinal)
        Exit Sub

    ElseIf what = "lg" Then
        commandFinal = "Logo [" & mIp.Infos.Project & "] : 客制开机logo"
    ElseIf what = "ani" Then
        commandFinal = "BootAnimation [" & mIp.Infos.Project & "] : 客制开机动画"
    ElseIf what = "wp" Then
        commandFinal = "Wallpaper [" & mIp.Infos.Project & "] : 客制默认壁纸"
    ElseIf what = "loc" Then
        commandFinal = "Locale [" & mIp.Infos.Project & "] : 默认语言"
    ElseIf what = "tz" Then
        commandFinal = "Timezone [" & mIp.Infos.Project & "] : 默认时区"
    ElseIf what = "di" Then
        commandFinal = "DisplayId [" & mIp.Infos.Project & "] : 版本号"
    ElseIf InStr(what, "bn=") = 1 Then
        commandFinal = "BuildNumber [" & mIp.Infos.Project & "] : 固定指纹信息（build number " & Replace(what, "bn=", "") & "）"
    ElseIf InStr(what, "sp=") = 1 Then
        If InStr(what, "-bn=") > 0 Then
            commandFinal = "GMS [" & mIp.Infos.Project & "] : 固定GMS信息（安全补丁日期 " & Replace(Split(what, "-bn=")(0), "sp=", "") & "、build number " & Split(what, "-bn=")(1) & "）"
        Else
            commandFinal = "GMS [" & mIp.Infos.Project & "] : 固定安全补丁日期 " & Replace(what, "sp=", "")
        End If
    ElseIf what = "bm" Then
        commandFinal = "MMI [" & mIp.Infos.Project & "] : 品牌，型号"
    ElseIf what = "bwm" Then
        commandFinal = "MMI [" & mIp.Infos.Project & "] : 蓝牙、WiFi热点、WiFi直连、盘符"
    ElseIf what = "mmi" Then
        commandFinal = "MMI [" & mIp.Infos.Project & "] : "
    ElseIf what = "st" Then
        commandFinal = "Settings [" & mIp.Infos.Project & "] : "    
    ElseIf what = "su" Then
        commandFinal = "SystemUI [" & mIp.Infos.Project & "] : "
    ElseIf what = "lc" Then
        commandFinal = "Launcher [" & mIp.Infos.Project & "] : "
    ElseIf what = "cam" Then
        commandFinal = "Camera [" & mIp.Infos.Project & "] : "
    ElseIf what = "bt" Then
        commandFinal = "Bluetooth [" & mIp.Infos.Project & "] : 默认蓝牙名称"
    ElseIf what = "wfap" Then
        commandFinal = "WiFi [" & mIp.Infos.Project & "] : 默认WiFi热点名称"
    ElseIf what = "wfdrt" Then
        commandFinal = "WiFi [" & mIp.Infos.Project & "] : 默认WiFi直连名称"
    ElseIf what = "mtp" Then
        commandFinal = "MTP [" & mIp.Infos.Project & "] : 默认盘符名称"
    ElseIf what = "brt" Then
        commandFinal = "Brightness [" & mIp.Infos.Project & "] : 默认亮度"
    ElseIf what = "ad" Then
        commandFinal = "Audio [" & mIp.Infos.Project & "] : 默认音量"
    ElseIf what = "slp" Then
        commandFinal = "Settings [" & mIp.Infos.Project & "] : 默认休眠时间"
    ElseIf what = "hp" Then
        commandFinal = "Browser [" & mIp.Infos.Project & "] : 默认网址"
    ElseIf what = "bat" Then
        commandFinal = "Battery [" & mIp.Infos.Project & "] : 电池检测容量"
    ElseIf what = "ft" Then
        commandFinal = "FactoryTest [" & mIp.Infos.Project & "] : "
    Else
        commandFinal = what & " [" & mIp.Infos.Project & "] : "
    End If
	Call copyStrAndPasteInXshell("git add weibu;git commit -m ""&Chr(34)&""" & commandFinal & """&Chr(34)&""")
End Sub

Sub CopyAdbPushCmd(which)
    Dim outPath, sourcePath, targetPath, finalStr
	outPath = mIp.Infos.getPathWithDriveSdk(mIp.Infos.OutPath)
    If which = "su" Then
	    sourcePath = outPath & "\system\system_ext\priv-app\MtkSystemUI"
		targetPath = "/system/system_ext/priv-app/"
		finalStr = "adb push " & sourcePath & " " & targetPath
	ElseIf which = "st" Then
	    sourcePath = outPath & "\system\system_ext\priv-app\MtkSettings"
		targetPath = "/system/system_ext/priv-app/"
		finalStr = "adb push " & sourcePath & " " & targetPath
	ElseIf which = "sl" Then
	    sourcePath = outPath & "\system\system_ext\priv-app\SearchLauncherQuickStep"
		targetPath = "/system/system_ext/priv-app/"
		finalStr = "adb push " & sourcePath & " " & targetPath
	ElseIf which = "cam" Then
	    sourcePath = outPath & "\system\system_ext\app\Camera"
		targetPath = "/system/system_ext/app/"
		finalStr = "adb push " & sourcePath & " " & targetPath
    ElseIf which = "fws" Then
	    sourcePath = outPath & "\system\framework\services.jar"
		targetPath = "/system/framework"
		finalStr = "adb push " & sourcePath & " " & targetPath
    ElseIf which = "ft" Then
	    sourcePath = outPath & "\system\priv-app\FactoryTest"
		targetPath = "/system/priv-app/"
		finalStr = "adb push " & sourcePath & " " & targetPath
    ElseIf which = "sr" Then
	    sourcePath = outPath & "\system\app\SoundRecorder_old"
		targetPath = "/system/app/"
		finalStr = "adb push " & sourcePath & " " & targetPath
    ElseIf which = "fwr" Then
	    sourcePath = outPath & "\system\framework\framework-res.apk"
		targetPath = "/system/framework/"
		finalStr = "adb push " & sourcePath & " " & targetPath
	End If

    If isU0SysSdk() Or isV0SysSdk() Then
        finalStr = Replace(finalStr, "\system\system_ext\", "\system_ext\")
        finalStr = Replace(finalStr, "/system/system_ext/", "/system_ext/")
    End If
	Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbClearCmd(which)
    Dim finalStr
    If which = "su" Then
		finalStr = "adb shell stop;adb shell start"
	ElseIf which = "st" Then
		finalStr = "adb shell pm clear com.android.settings"
	ElseIf which = "lc" Then
		finalStr = "adb shell pm clear com.android.launcher3"
	ElseIf which = "cam" Then
		finalStr = "adb shell pm clear com.mediatek.camera"
    ElseIf which = "ft" Then
		finalStr = "adb shell pm clear com.weibu.factorytest"
    ElseIf which = "sr" Then
		finalStr = "adb shell pm clear com.android.soundrecorder"
	End If
	Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbStartCmd(which)
    Dim finalStr
    If which = "ab" Then
		finalStr = "adb shell am start -a android.settings.DEVICE_INFO_SETTINGS"
	ElseIf which = "log" Then
		finalStr = "adb shell am start -n com.debug.loggerui/.MainActivity"
    ElseIf which = "ft" Then
		finalStr = "adb shell am start -n com.weibu.factorytest/.FactoryTest"
	End If
	Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbDumpsysCmd(which)
    Dim finalStr
    If which = "a" Then
		finalStr = "adb shell ""&Chr(34)&""dumpsys activity top | grep ACTIVITY | tail -n 1""&Chr(34)&"""
	ElseIf which = "f" Then
		finalStr = "adb shell ""&Chr(34)&""dumpsys activity top | grep '#[0-9]: ' | tail -n 1""&Chr(34)&"""
	ElseIf which = "r" Then
		finalStr = "adb shell ""&Chr(34)&""dumpsys activity activities | grep '* ActivityRecord{'""&Chr(34)&"""
	ElseIf which = "temp" Then
		finalStr = "adb shell dumpsys battery set temp"
	ElseIf which = "level" Then
		finalStr = "adb shell dumpsys battery set level"
	ElseIf which = "su" Then
		finalStr = "adb shell ""&Chr(34)&""dumpsys activity service com.android.systemui | grep --color""&Chr(34)&"""
	End If
	Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbLogcatCmd(which)
    Dim finalStr
    If which = "as" Then
        finalStr = "adb shell ""&Chr(34)&""logcat -s  ActivityTaskManager | grep START""&Chr(34)&"""
    End If
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbSettingsCmd(which)
    Dim finalStr
    If which = "sec" Then
        finalStr = "adb shell settings put secure clock_seconds 1"
    ElseIf which = "brt" Then
        finalStr = "adb shell settings get system screen_brightness"
    End If
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbInstallCmd(which)
    Dim finalStr
    If which = "att" Then
        finalStr = "adb install D:\APK\Antutu\antutu-benchmark-v10.apk"
    ElseIf which = "aida" Then
        finalStr = "adb install D:\APK\Antutu\aida64-v198.apk"
    ElseIf which = "dvc" Then
        finalStr = "adb install D:\APK\Antutu\DevCheck_v5.11_Mod.apk"
    ElseIf which = "z" Then
        finalStr = "adb install D:\APK\Antutu\CPU-Z-1.43.apk"
    ElseIf which = "hw" Then
        finalStr = "adb install D:\APK\Antutu\DeviceInfoHW-v5.20.1.apk"
    End If
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbGetGmsPropCmd()
    Dim finalStr
    finalStr = "adb shell ""&Chr(34)&""getprop | grep fingerprint | grep -v ro.bootimage.build.fingerprint | grep -v preview_sdk_fingerprint""&Chr(34)&"";adb shell ""&Chr(34)&""getprop | grep -E 'security_patch|gmsversion|base_os|first_api_level|clientidbase'""&Chr(34)&"""
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyQmakeCmd(which)
    Dim cmdStr
    If which = "sl" Then
        cmdStr = "qmake SearchLauncherQuickStep"
    ElseIf which = "st" Then
        cmdStr = "qmake MtkSettings"
    ElseIf which = "su" Then
        cmdStr = "qmake MtkSystemUI"
    ElseIf which = "ft" Then
        cmdStr = "qmake FactoryTest"
    ElseIf which = "fws" Then
        cmdStr = "mmm -j32 frameworks/base/services:services"
    ElseIf which = "fwr" Then
        cmdStr = "mmm -j32 frameworks/base/core/res"
    ElseIf which = "lot" Then
        cmdStr = "qmake GmsSampleIntegration"
    End If
    Call copyStrAndPasteInXshell(cmdStr)
End Sub

Sub modDisplayIdForOtaTest()
    Dim sedStr
    sedStr = getOTATestSedStr(True)
	If sedStr <> "" Then Call copyStrAndPasteInXshell(sedStr)
End Sub

Function getOTATestSedStr(showDiff)
    Dim buildinfo, keyStr, sedStr
	buildinfo = mIp.Infos.ProjectPath & "/config/buildinfo.sh"
	If Not isFileExists(buildinfo) Then buildinfo = mIp.Infos.DriverProjectPath & "/config/buildinfo.sh"
	If Not isFileExists(buildinfo) Then buildinfo = mIp.Infos.getOverlayPath("build/make/tools/buildinfo.sh")
	If Not isFileExists(buildinfo) Then buildinfo = "build/make/tools/buildinfo.sh"

	keyStr = "ro.build.display.id"
	If InStr(buildinfo, "/config/") Then
	    sedStr = "sed -i '/" & keyStr & "/s/$/-OTA_test/' " & buildinfo
	Else
	    sedStr = "sed -i '/" & keyStr & "/s/""&Chr(34)&""$/-OTA_test""&Chr(34)&""/' " & buildinfo
	End If
    If showDiff Then
        sedStr = sedStr & "; git diff " & buildinfo
    End If
    getOTATestSedStr = sedStr
End Function

Sub modSystemprop(whatArr)
    Dim systempropPath, cmdStr, keyStr, valueStr
	systempropPath = mIp.Infos.ProjectPath & "/config/system.prop"

    valueStr = whatArr(1)
	If whatArr(0) = "tz" Then
	    keyStr = "persist.sys.timezone"
		valueStr = Replace(valueStr, "/", "\/")
	ElseIf whatArr(0) = "loc" Then
	    keyStr = "persist.sys.locale"
	ElseIf whatArr(0) = "ftd" Then
	    keyStr = "ro.weibu.factorytest.disable_" & whatArr(1)
		valueStr = "1"
	Else
	    Exit Sub
	End If

    If Not isFileExists(systempropPath) Then
	    cmdStr = "echo -e ""&Chr(34)&""" & keyStr & "=" & valueStr & """&Chr(34)&""\n > " & systempropPath
		cmdStr = cmdStr & ";git diff " & systempropPath
	Else
	    If strExistInFile(systempropPath, keyStr) Then
		    cmdStr = "sed -i '/" & keyStr & "/s/.*/" & keyStr & "=" & valueStr & "/' " & systempropPath
		Else
	        cmdStr = "sed -i '$a " & keyStr & "=" & valueStr & "' " & systempropPath
		End If
		cmdStr = cmdStr & ";git diff " & systempropPath
	End If
	Call copyStrAndPasteInXshell(cmdStr)
End Sub

Sub modFingerprintInfos(version)
    
End Sub

Sub cpFileAndSetValue(whatArr)
    Dim cmdStr
    cmdStr = getCmdStrForCpFileAndSetValue(whatArr)
    If cmdStr = "" Then Exit Sub
    Call copyStrAndPasteInXshell(cmdStr)
End Sub

Function getCmdStrForCpFileAndSetValue(whatArr)
    Dim filePath, folderPath, keyStr, startStr, searchStr, valueStr, cmdStr
    If whatArr(0) = "gmsv" Then
	    filePath = "vendor/partner_gms/products/gms_package_version.mk"
		keyStr = "GMS_PACKAGE_VERSION_ID"
		startStr = " := "
		searchStr = keyStr & startStr
		valueStr = whatArr(1)
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

	ElseIf whatArr(0) = "sp" Then
        If InStr(mIp.Infos.Sdk, "\v_sys") Then
            filePath = "build/make/core/version_util.mk"
        Else
            filePath = "build/make/core/version_defaults.mk"
        End If
		keyStr = "PLATFORM_SECURITY_PATCH"
		startStr = " := "
		searchStr = keyStr & startStr
		valueStr = whatArr(1)
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

	    filePath = "vendor/mediatek/proprietary/buildinfo_vnd/device.mk"
		keyStr = "VENDOR_SECURITY_PATCH"
		searchStr = keyStr & startStr
		cmdStr = cmdStr & getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

	ElseIf whatArr(0) = "bn" Then
        Dim weibuConfig : weibuConfig = "build/make/core/weibu_config.mk"
        keyStr = "WEIBU_BUILD_NUMBER"
		startStr = " := "
        If isFileExists(weibuConfig) And isSplitSdkVnd() Then
            filePath = weibuConfig
            startStr = " ?= "
        Else
	        filePath = "device/mediatek/system/common/BoardConfig.mk"
            If InStr(mIp.Infos.Sdk, "_r") > 0 Then
                keyStr = "BUILD_NUMBER_WEIBU"
            End If
        End If
		searchStr = keyStr & startStr
		valueStr = whatArr(1)
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

	ElseIf whatArr(0) = "bn2" Then
	    filePath = "device/mediatek/system/common/BoardConfig.mk"
		keyStr = "WEIBU_BUILD_NUMBER"
		startStr = " := "
		searchStr = keyStr & startStr
		valueStr = whatArr(1)
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

		filePath = "device/mediatek/vendor/common/BoardConfig.mk"
		cmdStr = cmdStr & getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")
    
    ElseIf whatArr(0) = "fp" Then
        Dim version, buildId
        version = whatArr(1)
        If version = "15" Then
            buildId="AP3A.240905.015.A2"
        ElseIf version = "14" Then
            buildId="UP1A.231005.007"
        ElseIf version = "13" Then
            buildId="TP1A.220624.014"
        Else
            MsgBox("Unknown fp version: " & version)
            getCmdStrForCpFileAndSetValue = ""
            Exit Function
        End If
        cmdStr = getCpAndSedCmdStr("build/make/core/build_id.mk", "BUILD_ID", "=", buildId, "s")
        cmdStr = cmdStr & getCpAndSedCmdStr("build/make/core/sysprop.mk", "BUILD_FINGERPRINT := $(PRODUCT_BRAND)", "$(PLATFORM_VERSION)", version, "ss")

	ElseIf whatArr(0) = "bt" Then
	    If isSplitSdkSys() Then
	        filePath = "vendor/mediatek/proprietary/packages/modules/Bluetooth/system/btif/src/btif_dm.cc"
		Else
		    filePath = "system/bt/btif/src/btif_dm.cc"
		End If
		keyStr = "static char btif_default_local_name"
		startStr = " = "
		searchStr = keyStr
		valueStr = """&Chr(34)&""" & whatArr(1) & """&Chr(34)&"";"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

	ElseIf whatArr(0) = "mtp" Then
	    filePath = "frameworks/base/media/java/android/mtp/MtpDatabase.java"
		startStr = " = "
		searchStr = "mDeviceProperties.getString"
		valueStr = """&Chr(34)&""" & whatArr(1) & """&Chr(34)&"";"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")
		searchStr = "Build.MODEL"
		cmdStr = cmdStr & getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")
    
	ElseIf whatArr(0) = "wfap" Then
	    filePath = "packages/modules/Wifi/service/java/com/android/server/wifi/WifiApConfigStore.java"
		startStr = "("
		searchStr = "configBuilder.setSsid(Build.MODEL)"
		valueStr = """&Chr(34)&""" & whatArr(1) & """&Chr(34)&"");"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

	ElseIf whatArr(0) = "wfdrt" Then
	    filePath = "packages/modules/Wifi/service/java/com/android/server/wifi/p2p/WifiP2pServiceImpl.java"
		searchStr = "String getPersistedDeviceName()"
		valueStr = "            if (true) return ""&Chr(34)&""" & whatArr(1) & """&Chr(34)&"";"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "a")

	ElseIf whatArr(0) = "brand" Or whatArr(0) = "model" Or whatArr(0) = "manufacturer" Then
	    If isSplitSdkSys() Then
            filePath = "device/mediatek/system/" & mIp.Infos.SysTarget & "/sys_" & mIp.Infos.SysTarget & ".mk"
        Else
            If is8781Vnd() Then
	            filePath = "device/mediateksample/" & mIp.Infos.Product & "/vext_" & mIp.Infos.Product & ".mk"
            Else
	            filePath = "device/mediateksample/" & mIp.Infos.Product & "/vnd_" & mIp.Infos.Product & ".mk"
            End If
        End If
        keyStr = "PRODUCT_" & UCase(whatArr(0))
        startStr = " := "
		searchStr = keyStr & startStr
		valueStr = whatArr(1)
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

    ElseIf whatArr(0) = "name" Or whatArr(0) = "device" Then
        If Not isSplitSdkVnd() Then
            filePath = "device/mediatek/system/" & mIp.Infos.SysTarget & "/sys_" & mIp.Infos.SysTarget & ".mk"
        Else
            If is8781Vnd() Then
                filePath = "device/mediateksample/" & mIp.Infos.Product & "/vext_" & mIp.Infos.Product & ".mk"
            Else
                filePath = "device/mediateksample/" & mIp.Infos.Product & "/vnd_" & mIp.Infos.Product & ".mk"
            End If
        End If
        keyStr = "PRODUCT_SYSTEM_" & UCase(whatArr(0))
        startStr = " := "
        If mIp.Infos.Product = "tb8765ap1_bsp_1g_k419" Or _
                mIp.Infos.Product = "tb8766p1_64_bsp" Or _
                mIp.Infos.Product = "tb8788p1_64_bsp_k419" Or _
                mIp.Infos.Product = "tb8321p3_bsp" Or _
                mIp.Infos.Product = "tb8768p1_64_bsp"  Then
            searchStr = keyStr & startStr
            valueStr = whatArr(1)
            cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")
        Else
            searchStr = "PRODUCT_BRAND"
            valueStr = "PRODUCT_SYSTEM_" & UCase(whatArr(0)) & startStr & whatArr(1)
            cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "a")
        End If

    ElseIf whatArr(0) = "brt" Then
        If isSplitSdkSys() Then Call setT0SdkVnd()
	    filePath = "vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/res/values/config.xml"
        startStr = ">"
		searchStr = "config_screenBrightnessSettingDefaultFloat"
		valueStr = whatArr(1) & "</item>"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

    ElseIf whatArr(0) = "bat" Then
        filePath = getPowerProfilePath()
        startStr = ">"
		searchStr = "battery.capacity"
		valueStr = whatArr(1) & "</item>"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")
    
    ElseIf whatArr(0) = "sdk" Then
	    mIp.Sdk = Trim(whatArr(1))
	ElseIf whatArr(0) = "ssdk" Then
	    mIp.Infos.SysSdk = Trim(whatArr(1))
	ElseIf whatArr(0) = "spj" Then
	    mIp.Infos.SysProject = Trim(whatArr(1))
	ElseIf whatArr(0) = "fmw" Then
	    mIp.Firmware = Trim(whatArr(1))
	ElseIf whatArr(0) = "req" Then
	    mIp.Requirements = Trim(whatArr(1))
	ElseIf whatArr(0) = "zt" Then
	    mIp.Zentao = Trim(whatArr(1))

	End If
	getCmdStrForCpFileAndSetValue = cmdStr
End Function

Function getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, mode)
    Dim folderPath, cmdStr
	folderPath = getParentPath(filePath)

	If Not isFileExists(mIp.Infos.getOverlayPath(filePath)) Then
		cmdStr = cmdStr & "mkdir -p " & mIp.Infos.getOverlayPath(folderPath) & ";"
		cmdStr = cmdStr & "cp " & filePath & " " & mIp.Infos.getOverlayPath(folderPath) & ";"
	End If
    cmdStr = getSedStr(cmdStr, mIp.Infos.getOverlayPath(filePath), searchStr, startStr, valueStr, mode)
    cmdStr = getGitDiffCmd(cmdStr, filePath)

	getCpAndSedCmdStr = cmdStr
End Function

Function getSedStr(cmdStr, filePath, searchStr, startStr, valueStr, mode)
    If mode = "s" Then
        getSedStr = getSedCmd(cmdStr, searchStr, startStr & ".*$", startStr & valueStr, filePath)
    ElseIf mode = "ss" Then
        getSedStr = getSedCmd(cmdStr, searchStr, startStr, valueStr, filePath)
    ElseIf mode = "a" Then
        getSedStr = getSedAddCmd(cmdStr, searchStr, valueStr, filePath)
    End If
End Function

Function checkAndMkdir(folderPath)
    Dim cmdStr
    If Not isFolderExists(folderPath) Then
        if (isT0Sdk()) Then folderPath = Replace(folderPath, "../", "")
        cmdStr = "mkdir -p " & folderPath & ";"
    End If
    checkAndMkdir = cmdStr
End Function

Function checkMvOut(outPath, folders)
    Dim cmdStr, folder, outFolder, parentFolder, tmpFolder
    For Each folder In folders
        If isT0Sdk() Then folder = "../" & folder
        If Not isFolderExists(folder) Then
            MsgBox(folder & " does not exist!")
            checkMvOut = ""
            Exit Function
        End If
    Next

    For Each folder In folders
        outFolder = outPath & "/" & folder
        If isFolderExists(outFolder) Then
            MsgBox(outFolder & " already exist!")
            checkMvOut = ""
            Exit Function
        Else
            tmpFolder = getParentPath(outFolder)
            If parentFolder <> tmpFolder Then
                parentFolder = tmpFolder
                IF Not isFolderExists(parentFolder) Then
                    cmdStr = cmdStr & "mkdir -p " & parentFolder & ";"
                End If
            End If
            cmdStr = cmdStr & "mv " & folder & " " & parentFolder & ";"
        End If
    Next

    if isT0Sdk() Then cmdStr = Replace(cmdStr, "../", "")
    checkMvOut = cmdStr
End Function

Function checkMvIn(outPath, folders)
    Dim cmdStr, folder, outFolder, parentFolder
    For Each folder In folders
        outFolder = outPath & "/" & folder
        If Not isFolderExists(outFolder) Then
            MsgBox(outFolder & " does not exist!")
            checkMvIn = ""
            Exit Function
        End If
    Next

    For Each folder In folders
        Dim checkFd : checkFd = folder
        If isT0Sdk() Then checkFd = "../" & folder
        If isFolderExists(checkFd) Then
            MsgBox(folder & " already exist!")
            checkMvIn = ""
            Exit Function
        Else
            outFolder = outPath & "/" & folder
            parentFolder = getParentPath(folder)
            if parentFolder = "" Then parentFolder = "./"
            cmdStr = cmdStr & "mv " & outFolder & " " & parentFolder & ";"
        End If
    Next

    if isT0Sdk() Then cmdStr = Replace(cmdStr, "../", "")
    checkMvIn = cmdStr
End Function

Function getOutFoldersForMvIn()
    If isT0Sdk() Then
        If isT08168Sdk() Then
            getOutFoldersForMvIn = Array("merged", "sys/out", "vnd/out")
        ElseIf is8781Vnd() Then
            If isV0SysSdk() Then
                getOutFoldersForMvIn = Array("merged", "v_sys/out_sys", "vnd/out", "vnd/out_hal", "vnd/out_krn")
            ElseIf isU0SysSdk() Then
                getOutFoldersForMvIn = Array("merged", "u_sys/out_sys", "vnd/out", "vnd/out_hal", "vnd/out_krn")
            Else
                getOutFoldersForMvIn = Array("merged", "sys/out", "vnd/out", "vnd/out_hal", "vnd/out_krn")
            End If
        Else
            If isV0SysSdk() Then
                getOutFoldersForMvIn = Array("v_sys/merged", "v_sys/out_sys", "v_sys/out")
            ElseIf isU0SysSdk() Then
                getOutFoldersForMvIn = Array("merged", "u_sys/out", "vnd/out")
            Else
                getOutFoldersForMvIn = Array("merged", "sys/out", "vnd/out")
            End If
        End If
    ElseIf InStr(mIp.Infos.Sdk, "8168") > 0 Then
        getOutFoldersForMvIn = Array("out", "out_sys")
    Else
        getOutFoldersForMvIn = Array("out")
    End If
End Function

Function getOutFoldersForMvOut()
    If isT0Sdk() Then
        'v+v
        If isV0SysSdk() And isFolderExists("../v_sys/merged") And isFolderExists("../v_sys/out_sys") And isFolderExists("../v_sys/out") Then
            If isFolderExists("../v_sys/out_krn") Then
                getOutFoldersForMvOut = Array("v_sys/merged", "v_sys/out_sys", "v_sys/out", "v_sys/out_krn")
            Else
                getOutFoldersForMvOut = Array("v_sys/merged", "v_sys/out_sys", "v_sys/out")
            End If
        '8781
        ElseIf isFolderExists("../vnd/out_hal") Then
            '8781 s+v
            If isV0SysSdk() And isFolderExists("../v_sys/out_sys") Then
                If isFolderExists("../v_sys/out") Then
                    MsgBox("There are two sys out folders: out/ out_sys/")
                    getOutFoldersForMvOut = Array("")
                Else
                    getOutFoldersForMvOut = Array("merged", "v_sys/out_sys", "vnd/out", "vnd/out_hal", "vnd/out_krn")
                End If
            '8781 s+u
            ElseIf isU0SysSdk() And isFolderExists("../u_sys/out_sys") Then
                If isFolderExists("../u_sys/out") Then
                    MsgBox("There are two sys out folders: out/ out_sys/")
                    getOutFoldersForMvOut = Array("")
                Else
                    getOutFoldersForMvOut = Array("merged", "u_sys/out_sys", "vnd/out", "vnd/out_hal", "vnd/out_krn")
                End If
            '8781 s+t
            ElseIf isFolderExists("../sys/out_sys") Then
                getOutFoldersForMvOut = Array("merged", "sys/out_sys", "vnd/out", "vnd/out_hal", "vnd/out_krn")
            Else
                MsgBox("No sys out_sys!")
                getOutFoldersForMvOut = Array("")
            End If
        ElseIf isFolderExists("../vnd/out") Then
            If isU0SysSdk() And isFolderExists("../u_sys/out") Then
                getOutFoldersForMvOut = Array("merged", "u_sys/out", "vnd/out")
            ElseIf isFolderExists("../sys/out") Then
                getOutFoldersForMvOut = Array("merged", "sys/out", "vnd/out")
            Else
                MsgBox("No sys out!")
                getOutFoldersForMvOut = Array("")
            End If
        Else
            MsgBox("No vnd out!")
            getOutFoldersForMvOut = Array("")
        End If
    Else
        If isFolderExists("out_sys") And isFolderExists("out") Then
            getOutFoldersForMvOut = Array("out", "out_sys")
        ElseIf isFolderExists("out") Then
            getOutFoldersForMvOut = Array("out")
        Else
            MsgBox("No out!")
            getOutFoldersForMvOut = Array("")
        End If
    End If
End Function

Sub moveOutFoldersOut()
    If isSplitSdkVnd() Then Call setT0SdkSys()
    Dim innerId, idArr, taskNum, taskNumArr, buildType, workName, outName, outPath, outFolders, cmdStr
    innerId = getSysOutInfo("ro.build.display.inner.id")
    If innerId = "" Then MsgBox("Empty inner id!") : Exit Sub
    buildType = getSysOutInfo("ro.build.type")
    If buildType = "userdebug" Then buildType = "debug"
    idArr = Split(Split(innerId, ".")(4), "-")
    If isNumeric(idArr(UBound(idArr))) Then
        taskNum = idArr(UBound(idArr))
    Else
        taskNum = idArr(UBound(idArr) - 1)
    End If
    If Not isNumeric(taskNum) Then MsgBox("Invalid taskNum! " & taskNum) : Exit Sub
    workName = getWorkInfoWithTaskNum(taskNum, "work")
    if workName = "" Then Exit Sub
    outName = Split(workName, " ")(0)
    outPath = "../OUT/" & outName & "_" & buildType
    outFolders = getOutFoldersForMvOut()
    If outFolders(0) <> "" Then
        cmdStr = checkMvOut(outPath, outFolders)
        Call copyStrAndPasteInXshell(cmdStr)
    End If
End Sub

Sub moveOutFoldersIn(buildType)
    Dim outName, outPath, outFolders, cmdStr
    outName = Split(mIp.Infos.Work, " ")(0)
    outPath = "../OUT/" & outName & "_" & buildType
    outFolders = getOutFoldersForMvIn()
    cmdStr = checkMvIn(outPath, outFolders)
    Call copyStrAndPasteInXshell(cmdStr)
End Sub
