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
    If isT0Sdk() Or InStr(mIp.Infos.Sdk, "8168") > 0 Then
        commandFinal = getLunchItemInSplitBuild(buildType)
        Call CopyString(commandFinal)
        Exit Sub
    Else
        comboName = "full_" & mIp.Infos.Product & "-" & buildType
        commandFinal = "source build/envsetup.sh ; lunch " & comboName & " " & mIp.Infos.Project
    End If
    Call CopyString(commandFinal)
End Sub

Function getLunchItemInSplitBuild(buildType)
    Dim sysStr, vndStr, lunchStr, commandStr
    sysStr = "sys_" & mIp.Infos.SysTarget & "-" & buildType
    vndStr = "vnd_" & mIp.Infos.VndTarget & "-" & buildType

    If isT0Sdk() Then
        If isT0SdkVnd() Then Call setT0SdkSys()
        lunchStr = vndStr & " " & mIp.Infos.DriverProject & " " &  sysStr & " " & mIp.Infos.Project
        commandStr = "sed -i 's/^.*$/" & lunchStr & "/' lunch_item"
    Else
        lunchStr = sysStr & " " & vndStr & " " & mIp.Infos.Project
        lunchStr = "lunch_item=""&Chr(34)&""" & lunchStr & """&Chr(34)&"""

        Dim keyStr
        keyStr = "##Cusomer Settings"
        commandStr = "sed -i '/" & keyStr & "/i\" & lunchStr & "' split_build.sh;git diff split_build.sh"
    End If
    getLunchItemInSplitBuild = commandStr
End Function

Sub CommandOfOut()
    If Not mIp.hasProjectInfos() Then Exit Sub
    Call CopyString(mIp.Infos.DownloadOutPath)
End Sub

Sub CopyCleanCommand()
    if isT0Sdk() Then
        commandFinal = "cd sys/;git checkout .;git clean -df;cd ../vnd/;git checkout .;git clean -df;cd ../"
    Else
	    commandFinal = "git checkout .;git clean -df"
    End If
	Call CopyString(commandFinal)
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
        If isFileExists(path) Then
            Dim index : index = InStrRev(path, "/")
            folderPath = Left(path, index)
        Else
            folderPath = path
        End If

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

    commandFinal = relpaceSlashInPath(commandFinal)
    Call CopyString(commandFinal)
End Sub

Sub copyExportToolsPathCmd()
    commandFinal = "export PATH=$HOME/Tools:$PATH"
    Call CopyString(commandFinal)
End Sub

Sub CopyPathInWindows()
    If isFileExists(mIp.Infos.getOverlayPath(getOpenPath())) Then
        commandFinal = mIp.Infos.getPathWithDriveSdk(Replace(mIp.Infos.getOverlayPath(getOpenPath()), "/", "\"))
    Else
        commandFinal = mIp.Infos.getPathWithDriveSdk(Replace(getOpenPath(), "/", "\"))
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
    Call CopyString("cd " & path)
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
        getSedCmd = cmdStr & "sed -i '/" & checkBackslash(searchStr) & "/s/" & checkBackslash(replaceStr) & "/" & checkBackslash(newStr) & "/' " & mIp.Infos.getOverlayPath(filePath) & ";"
    End If
End Function

Function getSedAddCmd(cmdStr, searchStr, addStr, filePath)
    getSedAddCmd = cmdStr & "sed -i '/" & checkBackslash(searchStr) & "/a\" & checkBackslash(addStr) & "' " & mIp.Infos.getOverlayPath(filePath) & ";"
End Function

Function getGitDiffCmd(cmdStr, filePath)
    If isFileExists(mIp.Infos.getOverlayPath(filePath)) Then
        getGitDiffCmd = cmdStr & "git diff " & mIp.Infos.getOverlayPath(filePath) & ";"
    Else
        getGitDiffCmd = cmdStr & "git diff --no-index " & filePath & " " & mIp.Infos.getOverlayPath(filePath) & ";"
    End If
End Function

Function getMultiMkdirStr(arr, what)
    Dim str, path, ovlFolder, ovlFile
    For Each path In arr
	    ovlFolder = mIp.Infos.getOverlayPath(getFolderPath(path))
	    ovlFile = mIp.Infos.getOverlayPath(path)
	    If (Not (what = "lg" And InStr(path, "_kernel.bmp") > 0)) And (Not isFolderExists(ovlFolder)) Then
	        str =  str & "mkdir -p " & ovlFolder & ";"
		End If

	    If what = "lg" Then
	        str =  str & "cp ../File/logo.bmp " & ovlFolder & ";"
		ElseIf what = "ani" Then
		    If InStr(path, "bootanimation.zip") > 0 Then
	            str =  str & "cp ../File/bootanimation.zip " & ovlFolder & ";"
			ElseIf InStr(path, "products.mk") > 0 And Not isFileExists(ovlFile) Then
			    str =  str & "cp " & path & " " & ovlFolder & ";"
				str =  getSedCmd(str, "bootanimation", "#", "", path)
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

Sub mkdirLogo()
    Dim lg_fd, lg_u, lg_k
    lg_fd = "vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]/"
    lg_u = Replace(lg_fd & "[boot_logo]_uboot.bmp", "[boot_logo]", mIp.Infos.BootLogo)
    lg_k = Replace(lg_fd & "[boot_logo]_kernel.bmp", "[boot_logo]", mIp.Infos.BootLogo)

    Dim arr, finalStr
    arr = Array(lg_u, lg_k)
    finalStr = getMultiMkdirStr(arr, "lg")
    Call CopyString(finalStr)
End Sub

Sub mkdirBootAnimation()
    Dim ani_media, ani_product
    ani_media = "vendor/weibu_sz/media/bootanimation.zip"
    ani_product = "vendor/weibu_sz/products/products.mk"

    Dim arr, finalStr
    arr = Array(ani_media, ani_product)
    finalStr = getMultiMkdirStr(arr, "ani")
    Call CopyString(finalStr)
End Sub

Sub mkdirWallpaper(go)
    Dim wp_gms, wp_go1, wp_go2, wp1, wp2, wp3
    wp_gms = "vendor/partner_gms/overlay/AndroidSGmsBetaOverlay/res/drawable-nodpi/default_wallpaper.png"
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
    Call CopyString(finalStr)
End Sub

Sub CopyCommitInfo(what)
    If Not mIp.hasProjectInfos() Then Exit Sub
    If what = "" Then
	    commandFinal = "[" & mIp.Infos.Project & "] : "
        Call CopyString(commandFinal)
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
    ElseIf what = "bn" Then
        commandFinal = "BuildNumber [" & mIp.Infos.Project & "] : build number "
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
    ElseIf what = "lac" Then
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
    ElseIf what = "app" Then
        commandFinal = "App [" & mIp.Infos.Project & "] : "
    Else
        commandFinal = what & " [" & mIp.Infos.Project & "] : "
    End If
	Call CopyString("git add weibu;git commit -m ""&Chr(34)&""" & commandFinal & """&Chr(34)&""")
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
    ElseIf which = "sv" Then
	    sourcePath = outPath & "\system\framework\services.jar"
		targetPath = "/system/framework"
		finalStr = "adb push " & sourcePath & " " & targetPath
        sourcePath = outPath & "\system\framework\services.jar.bprof"
		finalStr = finalStr & ";adb push " & sourcePath & " " & targetPath
        sourcePath = outPath & "\system\framework\services.jar.prof"
		finalStr = finalStr & ";adb push " & sourcePath & " " & targetPath
        
        sourcePath = outPath & "\system\framework\oat\arm64\services.art"
        targetPath = "/system/framework/oat/arm64"
		finalStr = finalStr & ";adb push " & sourcePath & " " & targetPath
        sourcePath = outPath & "\system\framework\oat\arm64\services.odex"
		finalStr = finalStr & ";adb push " & sourcePath & " " & targetPath
        sourcePath = outPath & "\system\framework\oat\arm64\services.vdex"
		finalStr = finalStr & ";adb push " & sourcePath & " " & targetPath
	End If
	Call CopyString(finalStr)
End Sub

Sub modDisplayIdForOtaTest()
	Dim buildinfo, keyStr, sedStr
	buildinfo = mIp.Infos.ProjectPath & "/config/buildinfo.sh"
	If Not isFileExists(buildinfo) Then buildinfo = mIp.Infos.DriverProjectPath & "/config/buildinfo.sh"
	If Not isFileExists(buildinfo) Then buildinfo = mIp.Infos.getOverlayPath("build/make/tools/buildinfo.sh")
	If Not isFileExists(buildinfo) Then MsgBox("No buildinfo.sh found in overlay") : Exit Sub

	keyStr = "ro.build.display.id"
	If InStr(buildinfo, "/config/") Then
	    sedStr = "sed -i '/" & keyStr & "/s/$/-OTA_test/' " & buildinfo
	    Call CopyString(sedStr)
	Else
	    sedStr = "sed -i '/" & keyStr & "/s/""&Chr(34)&""$/-OTA_test""&Chr(34)&""/' " & buildinfo
	    Call CopyString(sedStr)
	End If
End Sub

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
	Call CopyString(cmdStr)
End Sub

Sub cpFileAndSetValue(whatArr)
    Dim filePath, folderPath, keyStr, eqStr, searchStr, valueStr, cmdStr
    If whatArr(0) = "gmsv" Then
	    filePath = "vendor/partner_gms/products/gms_package_version.mk"
		keyStr = "GMS_PACKAGE_VERSION_ID"
		eqStr = " := "
		searchStr = keyStr & eqStr
		valueStr = whatArr(1)
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

	ElseIf whatArr(0) = "sp" Then
	    filePath = "build/make/core/version_defaults.mk"
		keyStr = "PLATFORM_SECURITY_PATCH"
		eqStr = " := "
		searchStr = keyStr & eqStr
		valueStr = whatArr(1)
		cmdStr = getCpAndSedCmdStr(filePath, eqStr, valueStr, "s")

	    filePath = "vendor/mediatek/proprietary/buildinfo_vnd/device.mk"
		keyStr = "VENDOR_SECURITY_PATCH"
		searchStr = keyStr & eqStr
		cmdStr = cmdStr & getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

	ElseIf whatArr(0) = "bn" Then
	    filePath = "device/mediatek/system/common/BoardConfig.mk"
		If InStr(mIp.Infos.Sdk, "_r") > 0 Then
		    keyStr = "BUILD_NUMBER_WEIBU"
		Else
	        keyStr = "WEIBU_BUILD_NUMBER"
		End If
		eqStr = " := "
		searchStr = keyStr & eqStr
		valueStr = whatArr(1)
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

	ElseIf whatArr(0) = "bn2" Then
	    filePath = "device/mediatek/system/common/BoardConfig.mk"
		keyStr = "WEIBU_BUILD_NUMBER"
		eqStr = " := "
		searchStr = keyStr & eqStr
		valueStr = whatArr(1)
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

		filePath = "device/mediatek/vendor/common/BoardConfig.mk"
		cmdStr = cmdStr & getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

	ElseIf whatArr(0) = "bt" Then
	    If isT0SdkSys() Then
	        filePath = "vendor/mediatek/proprietary/packages/modules/Bluetooth/system/btif/src/btif_dm.cc"
		Else
		    filePath = "system/bt/btif/src/btif_dm.cc"
		End If
		keyStr = "static char btif_default_local_name"
		eqStr = " = "
		searchStr = keyStr
		valueStr = """&Chr(34)&""" & whatArr(1) & """&Chr(34)&"";"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

	ElseIf whatArr(0) = "mtp" Then
	    filePath = "frameworks/base/media/java/android/mtp/MtpDatabase.java"
		eqStr = " = "
		searchStr = "mDeviceProperties.getString"
		valueStr = """&Chr(34)&""" & whatArr(1) & """&Chr(34)&"";"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

	ElseIf whatArr(0) = "wfap" Then
	    filePath = "packages/modules/Wifi/service/java/com/android/server/wifi/WifiApConfigStore.java"
		eqStr = "("
		searchStr = "configBuilder.setSsid(Build.MODEL)"
		valueStr = """&Chr(34)&""" & whatArr(1) & """&Chr(34)&"");"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

	ElseIf whatArr(0) = "wfdrt" Then
	    filePath = "packages/modules/Wifi/service/java/com/android/server/wifi/p2p/WifiP2pServiceImpl.java"
		searchStr = "String getPersistedDeviceName()"
		valueStr = "            if (true) return ""&Chr(34)&""" & whatArr(1) & """&Chr(34)&"";"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "a")

	ElseIf whatArr(0) = "brand" Or whatArr(0) = "model" Or whatArr(0) = "manufacturer" Then
	    filePath = "device/mediateksample/" & mIp.Infos.Product & "/vnd_" & mIp.Infos.Product & ".mk"
        keyStr = "PRODUCT_" & UCase(whatArr(0))
        eqStr = " := "
		searchStr = keyStr & eqStr
		valueStr = whatArr(1)
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

    ElseIf whatArr(0) = "name" Or whatArr(0) = "device" Then
	    filePath = "device/mediateksample/" & mIp.Infos.Product & "/vnd_" & mIp.Infos.Product & ".mk"
        keyStr = "PRODUCT_SYSTEM_" & UCase(whatArr(0))
        eqStr = " := "
		searchStr = keyStr & eqStr
		valueStr = whatArr(1)
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

    ElseIf whatArr(0) = "brt" Then
        If isT0SdkSys() Then Call setT0SdkVnd()
	    filePath = "vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/res/values/config.xml"
        eqStr = ">"
		searchStr = "config_screenBrightnessSettingDefaultFloat"
		valueStr = whatArr(1) & "</item>"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

    ElseIf whatArr(0) = "bat" Then
        If isT0SdkSys() Then Call setT0SdkVnd()
	    filePath = "vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/power/res/xml/power_profile.xml"
        eqStr = ">"
		searchStr = "battery.capacity"
		valueStr = whatArr(1) & "</item>"
		cmdStr = getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, "s")

	End If
	Call CopyString(cmdStr)
End Sub

Function getCpAndSedCmdStr(filePath, searchStr, eqStr, valueStr, mode)
    Dim folderPath, cmdStr
	folderPath = getFolderPath(filePath)

	If Not isFileExists(mIp.Infos.getOverlayPath(filePath)) Then
		cmdStr = cmdStr & "mkdir -p " & mIp.Infos.getOverlayPath(folderPath) & ";"
		cmdStr = cmdStr & "cp " & filePath & " " & mIp.Infos.getOverlayPath(folderPath) & ";"
	End If
    cmdStr = getSedStr(cmdStr, filePath, searchStr, eqStr, valueStr, mode)
    cmdStr = getGitDiffCmd(cmdStr, filePath)

	getCpAndSedCmdStr = cmdStr
End Function

Function getSedStr(cmdStr, filePath, searchStr, eqStr, valueStr, mode)
    if mode = "s" Then
        getSedStr = getSedCmd(cmdStr, searchStr, eqStr & ".*$", eqStr & valueStr, filePath)
    ElseIf mode = "a" Then
        getSedStr = getSedAddCmd(cmdStr, searchStr, valueStr, filePath)
    End If
End Function

Sub mvOut(buildType, where)
    Dim cmdStr, outName, outPath
    outName = getProjectSimpleName() & "_" & buildType
    If isT0Sdk() Then
        outPath = "OUT/" & outName
        If where = "out" Then
            If Not isFolderExists("../" & outPath) Then
                cmdStr = cmdStr & "mkdir " & outPath & ";"
            End If
            If Not isFolderExists("../" & outPath & "/sys") Then
                cmdStr = cmdStr & "mkdir " & outPath & "/sys;"
            ElseIf isFolderExists("../" & outPath & "/sys/out") Then
                MsgBox("sys out/ is exist!")
                Exit Sub
            End If
            If Not isFolderExists("../" & outPath & "/vnd") Then
                cmdStr = cmdStr & "mkdir " & outPath & "/vnd;"
            ElseIf isFolderExists("../" & outPath & "/vnd/out") Or _
                    isFolderExists("../" & outPath & "/vnd/out_krn")Then
                MsgBox("vnd out/ is exist!")
                Exit Sub
            End If
            If isFolderExists("../" & outPath & "/merged") Then
                MsgBox("merged/ is exist!")
                Exit Sub
            End If
            If isFolderExists("../sys/out") And _
                    isFolderExists("../vnd/out") And _
                    isFolderExists("../vnd/out_krn") And _
                    isFolderExists("../merged") Then
                cmdStr = cmdStr & "mv sys/out " & outPath & "/sys/;"
                cmdStr = cmdStr & "mv vnd/out " & outPath & "/vnd/;"
                cmdStr = cmdStr & "mv vnd/out_krn " & outPath & "/vnd/;"
                cmdStr = cmdStr & "mv merged " & outPath & "/;"
            Else
                MsgBox("SDK out/ or merged/ is not exist!")
            End If
        ElseIf where = "in" Then
            If Not isFolderExists("../sys/out") And _
                    Not isFolderExists("../vnd/out") And _
                    Not isFolderExists("../vnd/out_krn") And _
                    Not isFolderExists("../merged") Then
                If isFolderExists("../" & outPath & "/sys/out") And _
                        isFolderExists("../" & outPath & "/vnd/out") And _
                        isFolderExists("../" & outPath & "/vnd/out_krn") And _
                        isFolderExists("../" & outPath & "/merged") Then
                    cmdStr = cmdStr & "mv " & outPath & "/sys/out sys/;"
                    cmdStr = cmdStr & "mv " & outPath & "/vnd/out vnd/;"
                    cmdStr = cmdStr & "mv " & outPath & "/vnd/out_krn vnd/;"
                    cmdStr = cmdStr & "mv " & outPath & "/merged ./;"
                Else
                    MsgBox("OUT out/ or merged/ is not exist!")
                End If
            Else
                MsgBox("SDK out/ or merged/ is exist!")
            End If
        End If
    ElseIf Not InStr(mIp.Infos.Sdk, "8168") > 0 Then
        outPath = "../OUT/" & outName
        If where = "out" Then
            If Not isFolderExists(outPath) Then
                cmdStr = cmdStr & "mkdir " & outPath & ";"
            ElseIf isFolderExists(outPath & "/out") Then
                MsgBox("out/ is exist!")
                Exit Sub
            End If
            If isFolderExists("out") Then
                cmdStr = cmdStr & "mv out " & outPath & "/;"
            Else
                MsgBox("SDK out/ is not exist!")
            End If
        ElseIf where = "in" Then
            If Not isFolderExists("out") Then
                If isFolderExists(outPath & "/out") Then
                    cmdStr = cmdStr & "mv " & outPath & "/out ./;"
                Else
                    MsgBox("OUT out/ is not exist!")
                End If
            ELse
                MsgBox("SDK out/ is exist!")
            End If
        End If
    Else
        outPath = "../OUT/" & outName
        If where = "out" Then
            If Not isFolderExists(outPath) Then
                cmdStr = cmdStr & "mkdir " & outPath & ";"
            ElseIf isFolderExists(outPath & "/out") Or _
                    isFolderExists(outPath & "/out_sys") Then
                MsgBox("OUT out/ is exist!")
                Exit Sub
            End If
            If isFolderExists("out") And isFolderExists("out_sys") Then
                cmdStr = cmdStr & "mv out " & outPath & "/;"
                cmdStr = cmdStr & "mv out_sys " & outPath & "/;"
            ELse
                MsgBox("SDK out/ or out_sys/ is not exist!")
            End If
        ElseIf where = "in" Then
            If Not isFolderExists("out") Or _
                    Not isFolderExists("out_sys") Then
                If isFolderExists(outPath & "/out") Or _
                    isFolderExists(outPath & "/out_sys") Then
                    cmdStr = cmdStr & "mv " & outPath & "/out ./;"
                    cmdStr = cmdStr & "mv " & outPath & "/out_sys ./;"
                Else
                    MsgBox("OUT out/ or out_sys/ is not exist!")
                End If
            ELse
                MsgBox("SDK out/ or out_sys/ is exist!")
            End If
        End If
    End If
    Call CopyString(cmdStr)
End Sub
